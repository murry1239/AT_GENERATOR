"""Two clean DOCX inputs are authoritative; tracked input is advisory only.

Comparison and range preparation do not launch Word. Word is used later for
pagination and pictures. Top-level tables remain intact, including nested tables.
"""
from __future__ import annotations

import copy
import hashlib
import io
import posixpath
import zipfile
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from lxml import etree as ET

from app import XML_REVISION_TAGS, _heading_style_ids, _paragraph_heading
from comparison_scope import detect_scope

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W}
DOC = "word/document.xml"
IGNORE_NODES = {"bookmarkStart", "bookmarkEnd", "proofErr", "lastRenderedPageBreak", "commentRangeStart", "commentRangeEnd", "commentReference"}


def content_signature(node, rels):
    """Ignore text formatting/run splits, but retain content and table structure.

    Drawing/object XML stays conservative: changed media, crops, equations and
    hyperlinks must not disappear just because their plain text is identical.
    """
    tokens = []

    def emit(value):
        if isinstance(value, str) and tokens and isinstance(tokens[-1], str):
            tokens[-1] += value
        elif value != "":
            tokens.append(value)

    def visit(element):
        name = local(element.tag)
        namespace = ET.QName(element).namespace
        if namespace == W:
            if name in IGNORE_NODES or name in {"pPr", "rPr", "tblPr", "tblGrid", "sectPr", "sdtPr", "sdtEndPr"}:
                return
            if name in {"tcPr", "trPr"}:
                # Cell merge/span changes are structural, not decoration.
                for child in element:
                    if local(child.tag) in {"gridSpan", "vMerge", "hMerge", "gridBefore", "gridAfter"}:
                        emit((child.tag, tuple(sorted(child.attrib.items()))))
                return
            if name in {"t", "delText"}:
                emit(element.text or "")
                return
            if name in {"r", "sdt", "sdtContent", "smartTag"}:
                for child in element:
                    visit(child)
                return
        attrs = []
        for key, value in element.attrib.items():
            attr = ET.QName(key)
            if attr.localname.startswith("rsid") or attr.localname in {"paraId", "textId"}:
                continue
            if attr.namespace == R:
                value = repr(rels.get(value, ("unresolved", value)))
            attrs.append((key, value))
        emit(("start", element.tag, tuple(sorted(attrs))))
        if element.text and element.text.strip():
            emit(element.text)
        for child in element:
            visit(child)
        emit(("end", element.tag))

    visit(node)
    return tuple(tokens)


def local(tag):
    return ET.QName(tag).localname


def parse(data):
    return ET.fromstring(data, ET.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False, remove_comments=True))


def digest(data):
    return hashlib.sha256(data).hexdigest()


def text_of(node):
    parts = []
    for event, child in ET.iterwalk(node, events=("start", "end")):
        tag = local(child.tag)
        if event == "start":
            if child.tag in {f"{{{W}}}t", f"{{{W}}}delText"}:
                parts.append(child.text or "")
            elif tag == "tab":
                parts.append("\t")
            elif tag in {"br", "cr"}:
                parts.append("\n")
        elif child.tag == f"{{{W}}}p":
            parts.append("\n")
    return "".join(parts).rstrip("\n")


@dataclass
class Block:
    index: int
    node: object
    text: str
    signature: tuple
    section: str
    kind: str
    region: str = '본문'


@dataclass
class Snapshot:
    path: Path
    content: bytes
    data: dict
    root: object
    blocks: list[Block]
    scope: dict
    warnings: list[str] = field(default_factory=list)


def read_snapshot(path: Path, *, allow_revisions=False) -> Snapshot:
    path = Path(path).resolve()
    if not path.is_file() or path.suffix.lower() != ".docx":
        raise ValueError("존재하는 DOCX 파일을 선택하십시오.")
    try:
        raw_bytes = path.read_bytes()
        with zipfile.ZipFile(io.BytesIO(raw_bytes)) as archive:
            data = {name: archive.read(name) for name in archive.namelist()}
        root = parse(data[DOC])
    except (zipfile.BadZipFile, KeyError, ET.XMLSyntaxError) as exc:
        raise ValueError("DOCX 구조를 읽을 수 없습니다. Word에서 다시 저장한 파일을 사용하십시오.") from exc
    if not allow_revisions:
        for name, content in data.items():
            if name.startswith("word/") and name.endswith(".xml"):
                if any(ET.QName(n).namespace == W and local(n.tag) in XML_REVISION_TAGS for n in parse(content).iter()):
                    raise ValueError("변경 전·후 파일에는 미처리 변경이력이 없어야 합니다. 정리된 파일을 입력하고 이력 파일은 선택 입력란에 지정하십시오.")
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("문서 본문을 찾지 못했습니다.")
    styles = parse(data["word/styles.xml"]) if "word/styles.xml" in data else None
    scope, scope_warnings = detect_scope(root, styles)
    rels = {}
    if "word/_rels/document.xml.rels" in data:
        for rel in parse(data["word/_rels/document.xml.rels"]):
            target = rel.get("Target", "")
            resolved = posixpath.normpath(posixpath.join("word", target)) if not target.startswith("/") else target.lstrip("/")
            rels[rel.get("Id")] = (rel.get("Type"), target if rel.get("TargetMode") == "External" else digest(data[resolved]) if resolved in data else resolved)

    headings = _heading_style_ids(styles)
    section = "본문"
    blocks = []
    warnings = list(scope_warnings)
    included = set(scope['included_blocks'])
    previous_region = None
    paragraph_index = 0
    for index, node in enumerate(body):
        name = local(node.tag)
        if name == "sectPr" or name in IGNORE_NODES:
            continue
        paragraphs = list(node.iter(f"{{{W}}}p"))
        if not paragraphs:
            warnings.append(f"문단이 없는 본문 요소({name})는 이미지 비교 범위 밖입니다. 원본을 확인하십시오.")
        first, end = paragraph_index, paragraph_index + len(paragraphs)
        paragraph_index = end
        if index not in included:
            continue
        region = '문서 앞부분' if scope['toc_detected'] and first < scope['front_matter_end'] else '본문'
        if region != previous_region:
            section = region
            previous_region = region
        value = text_of(node)
        if name == "p":
            section = _paragraph_heading(node, value, headings) or section
        kind = "table" if name == "tbl" or node.find(".//w:tbl", NS) is not None else "paragraph"
        signature = content_signature(node, rels)
        # Empty paragraphs only alter spacing. An image/equation/field without
        # plain text still has content tokens and must remain in the comparison.
        spacing_tags = {f"{{{W}}}{tag}" for tag in ("p", "tab", "br", "cr")}
        spacing_only = all(not token.strip() if isinstance(token, str) else
                           token[0] in {"start", "end"} and token[1] in spacing_tags
                           for token in signature)
        if name == "p" and spacing_only:
            continue
        blocks.append(Block(index, node, value, signature, section, kind, region))
    return Snapshot(path, raw_bytes, data, root, blocks, scope, warnings)


def compare_snapshots(before: Snapshot, after: Snapshot):
    """Content-first block diff. Repeated content uses autojunk=False."""
    warnings = before.warnings + after.warnings
    dependencies = ["word/styles.xml", "word/numbering.xml", "word/theme/theme1.xml", "word/fontTable.xml"]
    changed_dependencies = [name for name in dependencies if before.data.get(name) != after.data.get(name)]
    before_sect = before.root.findall(".//w:sectPr", NS)
    after_sect = after.root.findall(".//w:sectPr", NS)
    if [ET.tostring(n) for n in before_sect] != [ET.tostring(n) for n in after_sect]:
        changed_dependencies.append("구역/페이지 설정")
    if changed_dependencies:
        warnings.append("스타일·번호·테마·페이지 설정 차이가 있습니다. 내용 중심 비교이므로 서식만 다른 본문은 제외했습니다. 자동 문단번호/수준·레이아웃 변화는 원본에서 확인하십시오.")
    extra = sorted(n for n in set(before.data) | set(after.data)
                   if n.startswith(("word/header", "word/footer", "word/footnotes", "word/endnotes"))
                   and before.data.get(n) != after.data.get(n))
    if extra:
        warnings.append("본문 외 영역 차이(이 버전의 이미지 비교 범위 밖): " + ", ".join(extra))
    if before.scope != after.scope:
        warnings.append("변경 전·후 본문 범위 감지 결과가 다릅니다. analysis.json의 scope_before/scope_after를 확인하십시오.")
    matcher = SequenceMatcher(None, [b.signature for b in before.blocks], [b.signature for b in after.blocks], autojunk=False)
    pairs = []
    for tag, a1, a2, b1, b2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        # One row per block keeps captures bounded. Replacement blocks are
        # paired in document order; no semantic one-to-one equivalence claimed.
        for offset in range(max(a2-a1, b2-b1)):
            old = before.blocks[a1+offset] if a1+offset < a2 else None
            new = after.blocks[b1+offset] if b1+offset < b2 else None
            kind = "insert" if old is None else "delete" if new is None else "modify"
            pairs.append({"before": old, "after": new, "change_type": kind,
                          "needs_review": max(a2-a1,b2-b1) > 1 and tag == "replace",
                          "match_method": "content_document_order", "format_review": False})
    if any(x["needs_review"] for x in pairs):
        warnings.append("문단 재배치·분할·병합 검토가 필요합니다. 전후 블록은 문서 순서대로 대응시켰습니다.")
    return pairs, warnings


def tracked_advice(path, before, after):
    if path is None:
        return {"status": "not_supplied", "revision_count": 0, "records": []}
    tracked = read_snapshot(path, allow_revisions=True)

    def projection(side):
        root = copy.deepcopy(tracked.root)
        for node in list(root.iter()):
            if ET.QName(node).namespace != W or node.getparent() is None:
                continue
            name = local(node.tag)
            if name.endswith("PrChange") or name in {"tblGridChange", "numberingChange"}:
                node.getparent().remove(node)
            elif name in ({"ins", "moveTo"} if side == "before" else {"del", "moveFrom"}):
                node.getparent().remove(node)
        return text_of(root.find("w:body", NS))

    records = [{"type": local(n.tag), "id": n.get(f"{{{W}}}id", ""), "author": n.get(f"{{{W}}}author", ""),
                "date": n.get(f"{{{W}}}date", ""), "text": text_of(n)}
               for n in tracked.root.iter() if ET.QName(n).namespace == W and local(n.tag) in XML_REVISION_TAGS]
    matches = (projection("before") == text_of(before.root.find("w:body", NS)) and
               projection("after") == text_of(after.root.find("w:body", NS)))
    return {"status": "text_match_advisory_only" if matches else "mismatch_advisory_disabled",
            "revision_count": len(records), "records": records if matches else [],
            "note": "텍스트 수준의 보조 확인입니다. 서식·표 구조 일치 또는 변경기록의 완전성을 보증하지 않습니다. 전후 비교 결과를 대체하지 않습니다."}


def bookmark_copy(snapshot, pairs, side, target):
    """Add boundary bookmarks only to a temporary copy, without changing runs."""
    root = copy.deepcopy(snapshot.root)
    body = root.find("w:body", NS)
    original_nodes = list(body)
    existing = {n.get(f"{{{W}}}name") for n in root.iter(f"{{{W}}}bookmarkStart")}
    ids = [int(n.get(f"{{{W}}}id")) for n in root.iter(f"{{{W}}}bookmarkStart") if (n.get(f"{{{W}}}id") or "").isdigit()]
    bookmark_id = max(ids, default=0) + 1
    from app import WordEngine
    for i, pair in enumerate(pairs, 1):
        block = pair[side]
        if block is None:
            continue
        node = original_nodes[block.index]
        for boundary in ("start", "end"):
            name = WordEngine._bookmark_name(i, boundary)
            if name in existing:
                raise ValueError("입력 문서에 변분기 내부 북마크가 있습니다. 원본 파일을 다시 선택하십시오.")
            start = ET.Element(f"{{{W}}}bookmarkStart", {f"{{{W}}}id": str(bookmark_id), f"{{{W}}}name": name})
            end = ET.Element(f"{{{W}}}bookmarkEnd", {f"{{{W}}}id": str(bookmark_id)})
            pos = body.index(node) + (1 if boundary == "end" else 0)
            body.insert(pos, start)
            body.insert(pos+1, end)
            bookmark_id += 1
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in snapshot.data.items():
            archive.writestr(name, ET.tostring(root, encoding="UTF-8", xml_declaration=True) if name == DOC else data)


def validate_paths(before, after, tracked, output):
    inputs = [Path(before).resolve(), Path(after).resolve()]
    if tracked is not None:
        inputs.append(Path(tracked).resolve())
    if len(set(inputs)) != len(inputs):
        raise ValueError("변경 전·후·이력에는 서로 다른 파일을 선택하십시오.")
    for path in inputs:
        if not path.is_file() or path.suffix.lower() != ".docx":
            raise ValueError("입력 파일은 존재하는 DOCX여야 합니다.")
    output = Path(output).resolve()
    if output.suffix.lower() != ".bdsg":
        raise ValueError("출력 파일 확장자는 .bdsg여야 합니다.")
    outputs = [output] + [output.with_name(output.stem + suffix) for suffix in
                         ("_변경전_전체.docx", "_변경후_전체.docx", "_변경내용_미리보기.xlsx")]
    if any(p in inputs for p in outputs):
        raise ValueError("출력 경로가 입력 원본과 겹칩니다. 다른 출력 이름을 사용하십시오.")
    return outputs
