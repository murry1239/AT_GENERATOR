import unittest
import test_pair_analysis as fixtures
from test_pair_analysis import p, docx
from test_alpha05 import table, toc
from pair_analysis import read_snapshot, compare_snapshots


def heading(value):
    return p(value).replace('<w:p>', '<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr>', 1)


class InventoryScopeTests(unittest.TestCase):
    setUp = fixtures.PairTests.setUp
    snapshots = fixtures.PairTests.snapshots

    def test_inventory_between_two_front_tocs_keeps_body_and_inventory_changes(self):
        def body(value):
            return (table('양식 '+value)+p('목차')+toc('본문......'+value)+
                    p('[첨부 1] 안내문 '+value)+p('[첨부 2] 서식')+p('')+
                    table('List of Table')+toc('표......'+value)+heading('1. 본문')+
                    p('본문 내용 '+value)+p('별첨 1')+p('제외 내용 '+value))
        a, b = self.snapshots(body('이전'), body('이후'))
        changes, _ = compare_snapshots(a, b)
        self.assertEqual([x['before'].text for x in changes],
                         ['양식 이전', '[첨부 1] 안내문 이전', '본문 내용 이전'])
        self.assertEqual([x['before'].region for x in changes], ['문서 앞부분', '문서 앞부분', '본문'])
        self.assertEqual(a.scope['end_label'], '별첨 1')
        self.assertEqual(len(a.scope['front_matter_attachment_paragraphs']), 2)
        self.assertNotIn('List of Table', [x.text for x in a.blocks])

    def test_later_toc_inside_actual_appendix_does_not_extend_front_matter(self):
        def body(value):
            return (p('목차')+toc()+heading('1. 본문')+p('본문 '+value)+
                    p('부록 1')+p('목차')+toc('부록목차 '+value)+p('부록 내용 '+value))
        a, b = self.snapshots(body('이전'), body('이후'))
        changes, _ = compare_snapshots(a, b)
        self.assertEqual([x['before'].text for x in changes], ['본문 이전'])
        self.assertEqual(changes[0]['before'].region, '본문')
        self.assertEqual(a.scope['end_label'], '부록 1')
        self.assertLess(a.scope['front_matter_end'], a.scope['excluded_toc_paragraph_ranges'][-1][1])

    def test_inventory_after_only_toc_before_body_heading_is_preserved(self):
        def body(value):
            return (p('목차')+toc()+p('[첨부 1] 안내문')+p('[첨부 2] 서식')+
                    heading('1. 본문')+p('내용 '+value)+p('부록 1')+p('제외 '+value))
        a, b = self.snapshots(body('이전'), body('이후'))
        self.assertEqual(len(compare_snapshots(a, b)[0]), 1)
        self.assertIn('[첨부 1] 안내문', [x.text for x in a.blocks])
        self.assertEqual(a.scope['end_label'], '부록 1')
        self.assertEqual(len(a.scope['front_matter_attachment_paragraphs']), 2)

    def test_inventory_before_only_toc_is_not_an_appendix_cutoff(self):
        a, b = self.snapshots(p('[첨부 1] 자료')+p('목차')+toc()+p('본문 이전'),
                              p('[첨부 1] 자료')+p('목차')+toc()+p('본문 이후'))
        self.assertFalse(a.scope['appendix_detected'])
        self.assertEqual([x['before'].text for x in compare_snapshots(a,b)[0]], ['본문 이전'])

    def test_singular_and_plural_list_titles_in_tables_are_excluded(self):
        for title in ['List of Table', 'List of Tables', 'List of Figure', 'List of Figures']:
            with self.subTest(title=title):
                a, b = self.snapshots(table(title)+toc()+p('본문 이전'),
                                      table(title)+toc()+p('본문 이후'))
                self.assertEqual([x.text for x in a.blocks], ['본문 이전'])

    def test_strong_appendix_heading_followed_by_own_toc_is_still_cutoff(self):
        a, b = self.snapshots(p('목차')+toc()+heading('부록 1')+p('목차')+toc()+p('제외 이전'),
                              p('목차')+toc()+heading('부록 1')+p('목차')+toc()+p('제외 이후'))
        self.assertTrue(a.scope['appendix_detected'])
        self.assertEqual(a.scope['end_label'], '부록 1')
        self.assertEqual(compare_snapshots(a,b)[0], [])

    def test_unstyled_appendix_content_after_only_toc_is_not_skipped(self):
        a, b = self.snapshots(p('목차')+toc()+p('첨부 1')+p('실제 첨부 내용 이전'),
                              p('목차')+toc()+p('첨부 1')+p('실제 첨부 내용 이후'))
        self.assertTrue(a.scope['appendix_detected'])
        self.assertEqual(compare_snapshots(a,b)[0], [])

    def test_no_main_toc_with_appendix_own_toc_still_excludes_appendix(self):
        def body(value):
            return (heading('1. 본문')+p('본문 '+value)+heading('부록 1')+
                    p('목차')+toc()+p('부록 내용 '+value))
        a, b = self.snapshots(body('이전'), body('이후'))
        changes, _ = compare_snapshots(a,b)
        self.assertEqual([x['before'].text for x in changes], ['본문 이전'])
        self.assertEqual(changes[0]['before'].region, '본문')
        self.assertEqual(a.scope['end_label'], '부록 1')


if __name__ == '__main__': unittest.main()
