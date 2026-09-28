import json
import logging
import tempfile
import threading
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
from types import SimpleNamespace

from lxml import etree as ET
from openpyxl import load_workbook
from PIL import Image

import analyzer
import pair_runner
from pair_analysis import read_snapshot, compare_snapshots, tracked_advice, bookmark_copy, validate_paths, W


def docx(path, body, extra=None):
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", f'<w:document xmlns:w="{W}"><w:body>{body}<w:sectPr/></w:body></w:document>')
        for name, content in (extra or {}).items():
            archive.writestr(name, content)
    return path


def p(text):
    from xml.sax.saxutils import escape
    return '<w:p><w:r><w:t>'+escape(text)+'</w:t></w:r></w:p>'


class PairTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def snapshots(self, a, b):
        return read_snapshot(docx(self.root/'before.docx', a)), read_snapshot(docx(self.root/'after.docx', b))

    def test_identical_has_no_changes(self):
        a, b = self.snapshots(p('같은 내용'), p('같은 내용'))
        self.assertEqual(compare_snapshots(a, b)[0], [])

    def test_insert_delete_modify_and_repeated_paragraphs(self):
        a, b = self.snapshots(p('처음')+p('중복')+p('이전')+p('중복')+p('끝'), p('처음')+p('중복')+p('이후')+p('중복')+p('끝')+p('추가'))
        changes, _ = compare_snapshots(a,b)
        self.assertEqual([x['change_type'] for x in changes], ['modify','insert'])
        self.assertEqual(changes[0]['before'].text, '이전')
        reverse, _ = compare_snapshots(b,a)
        self.assertEqual(reverse[-1]['change_type'], 'delete')

    def test_nested_table_kept_as_one_block(self):
        def table(text):
            return '<w:tbl><w:tr><w:tc>'+p('바깥')+'<w:tbl><w:tr><w:tc>'+p(text)+'</w:tc></w:tr></w:tbl>'+p('')+'</w:tc></w:tr></w:tbl>'
        a,b=self.snapshots(table('이전'), table('이후'))
        changes,_=compare_snapshots(a,b)
        self.assertEqual(len(changes),1)
        self.assertEqual(changes[0]['before'].kind,'table')
        self.assertIn('이전',changes[0]['before'].text)

    def test_format_only_and_media_changes(self):
        a,b=self.snapshots(p('same'), '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>same</w:t></w:r></w:p>')
        self.assertEqual(len(compare_snapshots(a,b)[0]),1)
        body='<w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:drawing r:embed="rId1"/></w:r></w:p>'
        rels='<Relationships><Relationship Id="rId1" Type="image" Target="media/image.png"/></Relationships>'
        a=read_snapshot(docx(self.root/'a.docx',body,{'word/_rels/document.xml.rels':rels,'word/media/image.png':b'old'}))
        b=read_snapshot(docx(self.root/'b.docx',body,{'word/_rels/document.xml.rels':rels,'word/media/image.png':b'new'}))
        self.assertEqual(len(compare_snapshots(a,b)[0]),1)

    def test_revision_inputs_rejected(self):
        path=docx(self.root/'tracked.docx','<w:ins w:id="1">'+p('추가')+'</w:ins>')
        with self.assertRaisesRegex(ValueError,'미처리'):
            read_snapshot(path)

    def test_optional_history_match_and_mismatch_do_not_change_pair(self):
        a,b=self.snapshots(p('old'),p('new'))
        path=docx(self.root/'tracked.docx','<w:p><w:del w:id="1"><w:r><w:delText>old</w:delText></w:r></w:del><w:ins w:id="2"><w:r><w:t>new</w:t></w:r></w:ins></w:p>')
        self.assertEqual(tracked_advice(path,a,b)['status'],'text_match_advisory_only')
        docx(path,p('unrelated'))
        self.assertEqual(tracked_advice(path,a,b)['status'],'mismatch_advisory_disabled')
        self.assertEqual(len(compare_snapshots(a,b)[0]),1)

    def test_bookmarks_original_bytes_preserved(self):
        a,b=self.snapshots(p('old'),p('new'))
        original=a.path.read_bytes()
        changes,_=compare_snapshots(a,b)
        out=self.root/'working.docx'
        bookmark_copy(a,changes,'before',out)
        self.assertEqual(a.path.read_bytes(),original)
        with zipfile.ZipFile(out) as archive:
            root=ET.fromstring(archive.read('word/document.xml'))
        names=[n.get(f'{{{W}}}name') for n in root.iter(f'{{{W}}}bookmarkStart')]
        self.assertEqual(names,['BDSG_S_0001','BDSG_E_0001'])

    def test_paths_duplicate_and_output_collision(self):
        a,b=self.snapshots(p('a'),p('b'))
        with self.assertRaises(ValueError):
            validate_paths(a.path,a.path,None,self.root/'out.bdsg')
        with self.assertRaises(ValueError):
            validate_paths(a.path,b.path,None,self.root/'out.docx')
        collision=docx(self.root/'out_변경전_전체.docx',p('a'))
        with self.assertRaises(ValueError):
            validate_paths(collision,b.path,None,self.root/'out.bdsg')

    def test_dependency_and_external_story_warnings(self):
        a=read_snapshot(docx(self.root/'a.docx',p('same'),{'word/header1.xml':f'<w:hdr xmlns:w="{W}">{p("old")}</w:hdr>'}))
        b=read_snapshot(docx(self.root/'b.docx',p('same'),{'word/header1.xml':f'<w:hdr xmlns:w="{W}">{p("new")}</w:hdr>'}))
        changes,warnings=compare_snapshots(a,b)
        self.assertFalse(changes)
        self.assertTrue(any('본문 외' in w for w in warnings))

    def test_preview_illegal_chars_long_text_formula_and_images(self):
        image=self.root/'image.png'
        Image.new('RGB',(300,100),'white').save(image)
        original='=1+1\x00\x01\x0b'+('가'*33000)
        items=[{'index':1,'section':'S\x01','page':'전 1 / 후 2','before':original,'after':'=SUM(A1:A2)',
                'before_image':str(image),'after_image':str(image)}]
        output=self.root/'preview.xlsx'
        analyzer._write_preview_xlsx(output,items,'test',warnings=['test warning'])
        book=load_workbook(output)
        self.addCleanup(book.close)
        sheet=book['페이지별 변경 미리보기']
        self.assertEqual(len(sheet['G4'].value),32767)
        self.assertEqual(sheet['H4'].data_type,'s')
        self.assertNotIn('\x00',sheet['G4'].value)
        self.assertEqual(items[0]['before'],original)
        self.assertEqual(len(sheet._images),2)
        anchor=sheet._images[0].anchor
        self.assertEqual(anchor.editAs,'twoCell')
        self.assertEqual(anchor.to.colOff/anchor.to.rowOff,3)

    def test_no_changes_package_does_not_start_word(self):
        a,b=self.snapshots(p('same'),p('same'))
        output=self.root/'out.bdsg'
        with patch.object(pair_runner,'WordEngine',side_effect=AssertionError('Word not needed')):
            result=analyzer.create_analysis_package(a.path,b.path,None,output,lambda e:None,threading.Event(),logging.getLogger('test'))
        self.assertEqual(result['count'],0)
        with zipfile.ZipFile(output) as archive:
            manifest=json.loads(archive.read('analysis.json'))
            self.assertEqual(archive.read('변경전_전체.docx'),a.path.read_bytes())
            self.assertEqual(manifest['input_mode'],'pair')

    def test_cancel_preserves_previous_package(self):
        a,b=self.snapshots(p('a'),p('b'))
        out=self.root/'out.bdsg';out.write_bytes(b'previous')
        event=threading.Event();event.set()
        from app import OperationCancelled
        with self.assertRaises(OperationCancelled):
            analyzer.create_analysis_package(a.path,b.path,None,out,lambda e:None,event,logging.getLogger('test'))
        self.assertEqual(out.read_bytes(),b'previous')

    def test_word_orchestration_partial_capture_and_cleanup(self):
        a,b=self.snapshots(p('old'),p('new'))
        events=[];closed=[]
        class FakeDoc:
            def __init__(self,name): self.name=name
            def Repaginate(self): pass
            def Range(self,start,end):
                return SimpleNamespace(Information=lambda code: 2 if 'after' in self.name else 1)
        class FakeEngine:
            def __init__(self,**kwargs): self.progress=kwargs['progress']
            def _start_word(self,*args):
                return SimpleNamespace(Documents=SimpleNamespace(Open=lambda name,**kw: FakeDoc(name)))
            def _retry_word_call(self,label,fn): return fn()
            def _close_document(self,doc,label): closed.append(label)
            def _close_word(self,doc,word): self.progress({'kind':'word_closed'})
        def capture(doc,index,engine,path):
            if 'after' in doc.name: raise RuntimeError('simulated capture failure')
            path.parent.mkdir(parents=True,exist_ok=True)
            Image.new('RGB',(120,40),'white').save(path)
        logger=logging.getLogger('mock-word-test');logger.addHandler(logging.NullHandler());logger.propagate=False
        with patch.object(pair_runner,'WordEngine',FakeEngine):
            result=pair_runner.create_package(a.path,b.path,None,self.root/'out.bdsg',events.append,threading.Event(),logger,
                 preview_writer=analyzer._write_preview_xlsx,capture=capture,
                 bookmark_range=lambda *args:SimpleNamespace(Start=0,End=10),version=analyzer.VERSION,package_format=analyzer.PACKAGE_FORMAT)
        self.assertEqual(closed,['변경 전','변경 후'])
        self.assertEqual(result['count'],1)
        with zipfile.ZipFile(self.root/'out.bdsg') as archive:
            item=json.loads(archive.read('analysis.json'))['items'][0]
            self.assertEqual(item['page'],'전 1 / 후 2')
            self.assertEqual(item['before'],'old')
            self.assertEqual(item['after'],'new')
            self.assertIn('images/0001_before.png',archive.namelist())
            self.assertNotIn('after_image',item)
            self.assertEqual(item['image_status'],'텍스트 대체')
        self.assertEqual(events[-1]['stage'],'분석 완료')

    def test_changed_styles_includes_unchanged_text_for_review(self):
        styles=lambda bold:f'<w:styles xmlns:w="{W}"><w:style w:styleId="Normal"><w:rPr>{"<w:b/>" if bold else ""}</w:rPr></w:style></w:styles>'
        a=read_snapshot(docx(self.root/'a.docx',p('same'),{'word/styles.xml':styles(False)}))
        b=read_snapshot(docx(self.root/'b.docx',p('same'),{'word/styles.xml':styles(True)}))
        pairs,warnings=compare_snapshots(a,b)
        self.assertEqual(len(pairs),1)
        self.assertTrue(pairs[0]['needs_review'])
        self.assertTrue(warnings)

    def test_toc_and_appendix_scope_preserved(self):
        body=p('표지')+p('목차')+p('본문 시작')+p('본문')+p('부록 1')+p('제외 내용')
        snapshot=read_snapshot(docx(self.root/'scope.docx',body))
        self.assertEqual([b.text for b in snapshot.blocks],['본문 시작','본문'])


if __name__=='__main__':
    unittest.main()
