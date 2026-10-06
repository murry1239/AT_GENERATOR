import json
import logging
import threading
import unittest
import zipfile
from types import SimpleNamespace
from unittest.mock import Mock, patch

import analyzer
import pair_runner
import word_capture
from app import OperationCancelled
from pair_analysis import compare_snapshots
import test_pair_analysis as fixtures
from test_pair_analysis import p


class ContentTests(unittest.TestCase):
    setUp = fixtures.PairTests.setUp
    snapshots = fixtures.PairTests.snapshots

    def test_empty_paragraph_spacing_ignored_but_drawing_kept(self):
        a, b = self.snapshots(p('same')+p(''), p('same'))
        self.assertEqual(compare_snapshots(a, b)[0], [])
        a, b = self.snapshots(p('same'), p('same')+p('  ')+'<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        self.assertEqual(compare_snapshots(a, b)[0], [])
        a, b = self.snapshots(p('same'), p('same')+'<w:p><w:r><w:drawing/></w:r></w:p>')
        self.assertEqual(len(compare_snapshots(a, b)[0]), 1)

    def test_style_outline_and_run_splitting_ignored(self):
        changed = ('<w:p><w:pPr><w:pStyle w:val="Custom"/><w:outlineLvl w:val="2"/>'
                   '<w:numPr><w:numId w:val="7"/></w:numPr></w:pPr>'
                   '<w:r><w:t>sa</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>me</w:t></w:r></w:p>')
        a, b = self.snapshots(p('same'), changed)
        self.assertEqual(compare_snapshots(a, b)[0], [])

    def test_whitespace_and_break_changes_remain(self):
        for changed in (p('a  b'), '<w:p><w:r><w:t>a</w:t><w:br/><w:t>b</w:t></w:r></w:p>'):
            a, b = self.snapshots(p('a b'), changed)
            self.assertEqual(len(compare_snapshots(a, b)[0]), 1)

    def test_table_decoration_ignored_but_merge_retained(self):
        def table(props):
            return '<w:tbl><w:tr><w:tc><w:tcPr>'+props+'</w:tcPr>'+p('same')+'</w:tc></w:tr></w:tbl>'
        a, b = self.snapshots(table(''), table('<w:shd w:fill="FF0000"/>'))
        self.assertEqual(compare_snapshots(a, b)[0], [])
        a, b = self.snapshots(table(''), table('<w:gridSpan w:val="2"/>'))
        self.assertEqual(len(compare_snapshots(a, b)[0]), 1)

    def test_field_and_hyperlink_changes_remain(self):
        a, b = self.snapshots('<w:p><w:fldSimple w:instr="DATE"/></w:p>', '<w:p><w:fldSimple w:instr="TIME"/></w:p>')
        self.assertEqual(len(compare_snapshots(a, b)[0]), 1)
        a, b = self.snapshots('<w:p><w:hyperlink w:anchor="a">'+p('link')+'</w:hyperlink></w:p>',
                              '<w:p><w:hyperlink w:anchor="b">'+p('link')+'</w:hyperlink></w:p>')
        self.assertEqual(len(compare_snapshots(a, b)[0]), 1)

    def test_capture_circuit_breaker_keeps_all_text(self):
        a, b = self.snapshots(''.join(p(f'old {i}') for i in range(8)), ''.join(p(f'new {i}') for i in range(8)))
        closed = []
        doc = SimpleNamespace(Repaginate=lambda: None, Range=lambda *a: SimpleNamespace(Information=lambda n: 1))
        engine = SimpleNamespace(
            _start_word=lambda *a: SimpleNamespace(Documents=SimpleNamespace(Open=lambda *a, **kw: doc)),
            _retry_word_call=lambda label, fn: fn(),
            _close_document=lambda doc, label: closed.append(label), _close_word=lambda *a: None)
        capture = Mock(side_effect=RuntimeError('no picture'))
        logger = logging.getLogger('circuit-test'); logger.addHandler(logging.NullHandler()); logger.propagate = False
        output = self.root/'out.bdsg'
        with patch.object(pair_runner, 'WordEngine', return_value=engine):
            pair_runner.create_package(a.path, b.path, None, output, lambda e: None, threading.Event(), logger,
                preview_writer=analyzer._write_preview_xlsx, capture=capture,
                bookmark_range=lambda *a: SimpleNamespace(Start=0, End=10), version=analyzer.VERSION,
                package_format=analyzer.PACKAGE_FORMAT, package_version=analyzer.PACKAGE_VERSION)
        self.assertEqual(capture.call_count, 6)  # Three tries per source, not sixteen.
        with zipfile.ZipFile(output) as archive:
            manifest = json.loads(archive.read('analysis.json'))
        self.assertEqual(manifest['item_count'], 8)
        self.assertEqual(manifest['version'], '변분기 Alpha 1.0')
        self.assertEqual(manifest['producer_version'], '변분기 Alpha 1.1')
        self.assertEqual(manifest['schema_version'], 3)
        self.assertEqual(manifest['scope_policy'], 'front_matter_and_body_excluding_toc_and_appendices_v1')
        self.assertEqual(manifest['items'][-1]['after'], 'new 7')
        self.assertEqual(manifest['capture_statistics']['circuit_skipped'], 10)
        self.assertEqual(len(closed), 2)


class CaptureTests(unittest.TestCase):
    def engine(self):
        return SimpleNamespace(_check_cancelled=Mock(), _retry_word_call=lambda label, fn, **kw: fn())

    def test_direct_emf_does_not_touch_clipboard(self):
        with patch.object(word_capture, 'save_emf') as save, patch.object(word_capture, 'clipboard_capture') as clip:
            self.assertEqual(word_capture.capture_range(SimpleNamespace(EnhMetaFileBits=b'emf'), self.engine(), 'out.png'), 'direct_emf')
        save.assert_called_once(); clip.assert_not_called()

    def test_clipboard_fallback_once(self):
        with patch.object(word_capture, 'save_emf', side_effect=ValueError('bad emf')), patch.object(word_capture, 'clipboard_capture') as clip:
            self.assertEqual(word_capture.capture_range(SimpleNamespace(EnhMetaFileBits=b'emf'), self.engine(), 'out.png'), 'clipboard_bitmap')
        clip.assert_called_once()

    def test_failures_bounded_and_both_causes_reported(self):
        with patch.object(word_capture, 'save_emf', side_effect=ValueError('bad emf')), patch.object(word_capture, 'clipboard_capture', side_effect=RuntimeError('no bitmap')) as clip:
            with self.assertRaisesRegex(RuntimeError, 'bad emf.*no bitmap'):
                word_capture.capture_range(SimpleNamespace(EnhMetaFileBits=b'emf'), self.engine(), 'out.png')
        clip.assert_called_once()

    def test_cancellation_never_falls_back(self):
        with patch.object(word_capture, 'save_emf', side_effect=OperationCancelled('stop')), patch.object(word_capture, 'clipboard_capture') as clip:
            with self.assertRaises(OperationCancelled):
                word_capture.capture_range(SimpleNamespace(EnhMetaFileBits=b'emf'), self.engine(), 'out.png')
        clip.assert_not_called()
