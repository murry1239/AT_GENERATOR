import unittest
import test_pair_analysis as fixtures
from test_pair_analysis import docx, p, W
from pair_analysis import read_snapshot, compare_snapshots


def table(value):
    return '<w:tbl><w:tr><w:tc>' + p(value) + '</w:tc></w:tr></w:tbl>'


def toc(value='서론'):
    return ('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText>TO</w:instrText></w:r><w:r><w:instrText>C \\o "1-3"</w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>'+value+'</w:t></w:r></w:p>'
            '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGEREF _Toc1</w:instrText></w:r>'
            '<w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
            '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')


class ScopeTests(unittest.TestCase):
    setUp = fixtures.PairTests.setUp
    snapshots = fixtures.PairTests.snapshots

    def test_front_table_and_body_changed_but_toc_and_appendix_changes_excluded(self):
        a, b = self.snapshots(table('버전 1')+p('목차')+toc('이전 목차')+p('이전 본문')+p('별첨 1')+p('이전 별첨'),
                              table('버전 2')+p('목차')+toc('이후 목차')+p('이후 본문')+p('별첨 1')+p('이후 별첨'))
        pairs, _ = compare_snapshots(a,b)
        self.assertEqual([x['before'].text for x in pairs], ['버전 1', '이전 본문'])
        self.assertEqual(a.blocks[0].region, '문서 앞부분')
        self.assertEqual(a.blocks[-1].region, '본문')
        self.assertTrue(a.scope['toc_detected']); self.assertTrue(a.scope['appendix_detected'])

    def test_front_table_retained_without_toc_and_appendix_title_in_table_is_not_boundary(self):
        a, b = self.snapshots(table('별첨 1 참조')+p('이전 본문')+p('부록 1')+p('이전 제외'),
                              table('별첨 2 참조')+p('이후 본문')+p('부록 1')+p('이후 제외'))
        self.assertEqual(len(compare_snapshots(a,b)[0]), 2)

    def test_manual_toc_leaders_and_tabs_excluded(self):
        entry = '<w:p><w:r><w:t>1. 서론</w:t><w:tab/><w:t>2</w:t></w:r></w:p>'
        a, b = self.snapshots(p('표지')+p('목차')+entry+p('부록 1......7')+p('본문 첫 문단')+p('첨부 1')+p('제외'),
                              p('표지')+p('목차')+entry+p('부록 1......8')+p('본문 첫 문단')+p('첨부 1')+p('제외 변경'))
        self.assertEqual(compare_snapshots(a,b)[0], [])
        self.assertEqual([x.text for x in a.blocks], ['표지', '본문 첫 문단'])

    def test_content_control_toc_excludes_only_that_control(self):
        control = '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr><w:sdtContent>'+p('목차 항목')+'</w:sdtContent></w:sdt>'
        a, b = self.snapshots(table('이전 표')+control+table('본문 표'), table('이후 표')+control+table('본문 표'))
        self.assertEqual([x.text for x in a.blocks], ['이전 표', '본문 표'])
        self.assertEqual(len(compare_snapshots(a,b)[0]), 1)

    def test_multiple_tocs_do_not_remove_content_between_them(self):
        a, b = self.snapshots(p('표지')+p('목차')+toc()+table('이전 안내')+p('표 목차')+toc('표 목록')+p('본문'),
                              p('표지')+p('목차')+toc()+table('이후 안내')+p('표 목차')+toc('표 목록')+p('본문'))
        self.assertEqual(len(compare_snapshots(a,b)[0]), 1)
        self.assertEqual([x.text for x in a.blocks], ['표지', '이전 안내', '본문'])

    def test_ordinary_appendix_reference_sentence_remains(self):
        a, b = self.snapshots(p('본문')+p('부록 1을 참조하십시오.')+p('끝'), p('본문')+p('부록 2를 참조하십시오.')+p('끝'))
        self.assertFalse(a.scope['appendix_detected']); self.assertEqual(len(compare_snapshots(a,b)[0]), 1)

    def test_mixed_toc_and_front_table_is_excluded_with_warning(self):
        body = '<w:tbl><w:tr><w:tc>'+p('목차')+p('서론......1')+'</w:tc><w:tc>'+p('검토 대상 양식')+'</w:tc></w:tr></w:tbl>'+p('본문')
        snap = read_snapshot(docx(self.root/'mixed.docx', body))
        self.assertEqual([x.text for x in snap.blocks], ['본문'])
        self.assertTrue(any('블록 전체를 제외' in w for w in snap.warnings))

    def test_unclosed_toc_field_is_rejected_instead_of_silently_excluding_body(self):
        body = toc().replace('<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>', '')+p('본문')
        with self.assertRaisesRegex(ValueError, '끝 표식'): read_snapshot(docx(self.root/'broken.docx', body))


if __name__ == '__main__': unittest.main()
