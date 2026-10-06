"""Discontinuous comparison scope: front matter and body, excluding TOCs and appendices."""
import re
from app import _heading_style_ids, _paragraph_heading

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
POLICY = 'front_matter_and_body_excluding_toc_and_appendices_v1'
TITLE = re.compile(r'^(?:목\s*차|차례|표\s*목차|그림\s*목차|표\s*차례|그림\s*차례|table\s+of\s+contents|table\s+of\s+figures|list\s+of\s+(?:tables?|figures?)|contents)\s*[:：]?$', re.I)
APPENDIX = re.compile(r'^\s*(?:\[\s*)?(?:제\s*)?(?:\d+[.\s]+)?(?:부록|붙임|별첨|첨부|appendix|appendices|annex)(?=\s|$|[\]\d:：.\-])', re.I)
FIELD = re.compile(r'^\s*TOC(?:\s|\\|$)', re.I)
ENTRY = re.compile(r'(?:\t|[.·…]{2,}|\s{2,})\s*(?:\d+(?:[-–]\d+)?|[ivxlcdm]+)\s*$', re.I)
INVENTORY_ENTRY = re.compile(r'^\s*(?:\[\s*)?(?:부록|붙임|별첨|첨부|appendix|annex)\s*(?:\d+|[a-z])\s*(?:\]|[.:：\-])?(?=\s|$)', re.I)
INVENTORY_TITLE = re.compile(r'^\s*(?:첨부|별첨|부록|붙임)\s*(?:자료|문서)?\s*목록\s*[:：]?\s*$', re.I)


def text(p):
    parts = []
    for n in p.iter():
        if n.tag == f'{{{W}}}t': parts.append(n.text or '')
        elif n.tag == f'{{{W}}}tab': parts.append('\t')
    return ''.join(parts).strip()


def ranges(indices):
    result = []
    for n in sorted(indices):
        if result and n == result[-1][1]: result[-1][1] += 1
        else: result.append([n, n + 1])
    return result


def leading_toc_ranges(paragraphs, omitted, heading_styles):
    """Join front TOCs separated by a plain attachment inventory, not body text.

    Later TOCs inside actual appendices must not extend this front-matter region.
    A heading or any substantive paragraph in the gap ends the leading cluster.
    """
    toc_ranges = ranges(omitted)
    if not toc_ranges:
        return []
    # A document may have no main TOC and only a TOC within a real appendix.
    # A genuine appendix heading before that TOC prevents it becoming front matter.
    for p in paragraphs[:toc_ranges[0][0]]:
        value = text(p)
        if (APPENDIX.match(value) and not INVENTORY_TITLE.fullmatch(value) and
            not any(parent.tag == f'{{{W}}}tbl' for parent in p.iterancestors()) and
            _paragraph_heading(p, value, heading_styles) is not None):
            return []
    leading = [toc_ranges[0]]
    for current in toc_ranges[1:]:
        gap = paragraphs[leading[-1][1]:current[0]]
        def front_list(p):
            value = text(p)
            return (not value or INVENTORY_TITLE.fullmatch(value) or
                    (INVENTORY_ENTRY.match(value) and
                     _paragraph_heading(p, value, heading_styles) is None))
        if not all(front_list(p) for p in gap):
            break
        leading.append(current)
    return leading


def detect_scope(root, styles=None):
    body = root.find('w:body', NS)
    paragraphs = list(body.iter(f'{{{W}}}p'))
    positions = {p: n for n, p in enumerate(paragraphs)}
    heading_styles = _heading_style_ids(styles)
    toc_styles = set()
    if styles is not None:
        for style in styles.findall('w:style', NS):
            name = style.find('w:name', NS)
            value = name.get(f'{{{W}}}val', '') if name is not None else ''
            sid = style.get(f'{{{W}}}styleId', '')
            if re.search(r'^toc\s*\d*$|table\s*of\s*figures|목차|차례', value, re.I) or re.match(r'^TOC\d+$', sid, re.I):
                toc_styles.add(sid)
    omitted = set()
    stack = []
    warnings = []
    for idx, p in enumerate(paragraphs):
        for node in p.iter():
            if node.tag == f'{{{W}}}fldChar':
                kind = node.get(f'{{{W}}}fldCharType')
                if kind == 'begin': stack.append({'start': idx, 'instruction': ''})
                elif kind == 'end' and stack:
                    current = stack.pop()
                    if FIELD.match(current['instruction']): omitted.update(range(current['start'], idx + 1))
            elif node.tag == f'{{{W}}}instrText' and stack:
                stack[-1]['instruction'] += node.text or ''
            elif node.tag == f'{{{W}}}fldSimple' and FIELD.match(node.get(f'{{{W}}}instr', '')):
                omitted.add(idx)
        style = p.find('w:pPr/w:pStyle', NS)
        if style is not None and style.get(f'{{{W}}}val') in toc_styles:
            omitted.add(idx)
    if any(FIELD.match(current['instruction']) for current in stack):
        raise ValueError('끝 표식이 없는 목차 필드가 있습니다. Word에서 목차를 갱신하고 다시 저장하십시오.')
    for control in body.iter(f'{{{W}}}sdt'):
        gallery = [g.get(f'{{{W}}}val', '') for g in control.iter(f'{{{W}}}docPartGallery')]
        if any(re.search(r'table of contents|table of figures|목차|차례', value, re.I) for value in gallery):
            omitted.update(positions[p] for p in control.iter(f'{{{W}}}p'))
    # A typed TOC needs a title plus page leaders/tab or TOC styles. Stop at real content.
    for idx, p in enumerate(paragraphs):
        if not TITLE.fullmatch(text(p)): continue
        omitted.add(idx)
        for n in range(idx + 1, len(paragraphs)):
            value = text(paragraphs[n])
            if n in omitted or not value or TITLE.fullmatch(value) or ENTRY.search(value):
                omitted.add(n)
            else:
                break
    # An attachment inventory between leading contents/table/figure TOCs is front
    # matter. Treating its first item as a cutoff would silently drop all body text.
    leading = leading_toc_ranges(paragraphs, omitted, heading_styles)
    front_end = leading[-1][1] if leading else 0
    # Some forms place the inventory after their only TOC. A plain inventory
    # followed by a genuine body heading is still front matter. A real appendix
    # heading, or inventory followed by ordinary appendix content, is not skipped.
    if leading:
        n = front_end
        inventory_found = False
        while n < len(paragraphs):
            value = text(paragraphs[n])
            if not value:
                n += 1
                continue
            if (INVENTORY_TITLE.fullmatch(value) or
                (INVENTORY_ENTRY.match(value) and
                 _paragraph_heading(paragraphs[n], value, heading_styles) is None)):
                inventory_found = True
                n += 1
                continue
            if (inventory_found and not APPENDIX.match(value) and
                _paragraph_heading(paragraphs[n], value, heading_styles) is not None):
                front_end = n
            break
    ignored_front_titles = []
    # Never interpret an appendix reference inside a front-matter form/table as its start.
    appendix_start = len(paragraphs)
    end_label = '문서 끝'
    for idx, p in enumerate(paragraphs):
        if idx in omitted or any(parent.tag == f'{{{W}}}tbl' for parent in p.iterancestors()): continue
        value = text(p)
        if not APPENDIX.match(value): continue
        if (idx < front_end and
            (INVENTORY_TITLE.fullmatch(value) or INVENTORY_ENTRY.match(value)) and
            _paragraph_heading(p, value, heading_styles) is None):
            ignored_front_titles.append(idx)
            continue
        heading = _paragraph_heading(p, value, heading_styles)
        # Ordinary sentences referring to attachments remain in the comparison.
        if heading is None and (len(value) > 100 or re.search(r'참조|참고|확인|포함되어|기재되어|수록되어|첨부된', value)):
            continue
        appendix_start = idx; end_label = value[:120]; break
    blocks = []
    excluded = []
    for index, node in enumerate(body):
        ps = list(node.iter(f'{{{W}}}p'))
        ids = [positions[p] for p in ps]
        if not ids: continue
        toc_overlap = any(i in omitted for i in ids)
        appendix_overlap = any(i >= appendix_start for i in ids)
        if toc_overlap or appendix_overlap:
            reason = '목차' if toc_overlap else '별첨·부록'
            excluded.append({'block': index, 'reason': reason})
            if any(i not in omitted and i < appendix_start for i in ids):
                warnings.append(f'블록 {index}: {reason}와 비교 대상 내용이 같은 표/콘텐츠 컨트롤에 있어 블록 전체를 제외했습니다. 원문에서 분리한 뒤 재분석하십시오.')
        else:
            blocks.append(index)
    return dict(policy=POLICY, paragraph_start=0, paragraph_end=appendix_start,
                toc_detected=bool(omitted), appendix_detected=appendix_start < len(paragraphs),
                start_label='문서 시작(앞부분 포함)', end_label=end_label,
                excluded_toc_paragraph_ranges=ranges(omitted), excluded_blocks=excluded,
                included_blocks=blocks, front_matter_end=front_end,
                leading_toc_paragraph_ranges=leading,
                front_matter_attachment_paragraphs=ignored_front_titles), warnings
