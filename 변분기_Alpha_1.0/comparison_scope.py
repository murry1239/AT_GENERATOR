"""Discontinuous comparison scope: front matter and body, excluding TOCs and appendices."""
import re
from app import _heading_style_ids, _paragraph_heading

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
POLICY = 'front_matter_and_body_excluding_toc_and_appendices_v1'
TITLE = re.compile(r'^(?:목\s*차|차례|표\s*목차|그림\s*목차|표\s*차례|그림\s*차례|table\s+of\s+contents|table\s+of\s+figures|list\s+of\s+(?:tables|figures)|contents)\s*[:：]?$', re.I)
APPENDIX = re.compile(r'^\s*(?:\[\s*)?(?:제\s*)?(?:\d+[.\s]+)?(?:부록|붙임|별첨|첨부|appendix|appendices|annex)(?=\s|$|[\]\d:：.\-])', re.I)
FIELD = re.compile(r'^\s*TOC(?:\s|\\|$)', re.I)
ENTRY = re.compile(r'(?:\t|[.·…]{2,}|\s{2,})\s*(?:\d+(?:[-–]\d+)?|[ivxlcdm]+)\s*$', re.I)


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
    # Never interpret an appendix reference inside a front-matter form/table as its start.
    appendix_start = len(paragraphs)
    end_label = '문서 끝'
    for idx, p in enumerate(paragraphs):
        if idx in omitted or any(parent.tag == f'{{{W}}}tbl' for parent in p.iterancestors()): continue
        value = text(p)
        if not APPENDIX.match(value): continue
        heading = _paragraph_heading(p, value, heading_styles)
        # Ordinary sentences referring to attachments remain in the comparison.
        if heading is None and (len(value) > 100 or re.search(r'참조|참고|확인|포함되어|기재되어|수록되어|첨부된', value)):
            continue
        appendix_start = idx; end_label = value[:120]; break
    blocks = []
    excluded = []
    toc_end = max(omitted, default=-1) + 1
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
                included_blocks=blocks, front_matter_end=toc_end), warnings
