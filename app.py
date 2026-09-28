from __future__ import annotations

import gc
import os
import logging
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import uuid
import zipfile
import bisect
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor
from contextlib import suppress
from dataclasses import dataclass
from pathlib import Path
from tkinter import Tk, DoubleVar, StringVar, filedialog, messagebox, simpledialog
from tkinter import ttk
from typing import Any, Callable


APP_TITLE = "변분기 내부 엔진 - Alpha 0.2"
APP_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = APP_DIR / "assets" / "change_table_template.docx"
STALL_WARNING_SECONDS = 120
FORCE_STOP_SECONDS = 8
WORD_CALL_RETRY_ATTEMPTS = 20
WORD_CALL_RETRY_MAX_DELAY = 2.0
WORD_SAVE_RETRY_ATTEMPTS = 80


# Microsoft Word constants. Numeric values avoid a makepy dependency.
WD_ACTIVE_END_ADJUSTED_PAGE_NUMBER = 1
WD_WITHIN_TABLE = 12
WD_OUTLINE_LEVEL_BODY_TEXT = 10
WD_REVISION_INSERT = 1
WD_REVISION_DELETE = 2
WD_REVISION_REPLACE = 9
WD_REVISION_MOVED_FROM = 14
WD_REVISION_MOVED_TO = 15
WD_REVISION_CELL_INSERTION = 16
WD_REVISION_CELL_DELETION = 17
WD_COLOR_RED = 255
WD_COLOR_BLUE = 16711680
WD_UNDERLINE_SINGLE = 1
WD_DO_NOT_SAVE_CHANGES = 0
WD_COLLAPSE_START = 1
WD_ORIENT_LANDSCAPE = 1
WD_AUTOFIT_WINDOW = 2
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_COLOR_GRAY_50 = 8421504

# Transient COM errors returned while Word is busy with pagination/layout work.
RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846
RPC_E_SERVERCALL_REJECTED = -2147417845
RETRYABLE_WORD_HRESULTS = {
    RPC_E_CALL_REJECTED,
    RPC_E_SERVERCALL_RETRYLATER,
    RPC_E_SERVERCALL_REJECTED,
}

OLD_SIDE_REVISIONS = {
    WD_REVISION_DELETE,
    WD_REVISION_MOVED_FROM,
    WD_REVISION_CELL_DELETION,
}
NEW_SIDE_REVISIONS = {
    WD_REVISION_INSERT,
    WD_REVISION_MOVED_TO,
    WD_REVISION_CELL_INSERTION,
}

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = {"w": W_NS}
W_VAL = f"{{{W_NS}}}val"
XML_REVISION_TAGS = {
    "ins",
    "del",
    "moveFrom",
    "moveTo",
    "cellIns",
    "cellDel",
    "cellMerge",
    "rPrChange",
    "pPrChange",
    "tblPrChange",
    "tblGridChange",
    "trPrChange",
    "tcPrChange",
    "sectPrChange",
    "numberingChange",
}


class OperationCancelled(Exception):
    """Raised when the user requests cancellation between Word operations."""


def format_elapsed(seconds: float) -> str:
    total = max(0, int(seconds))
    minutes, seconds = divmod(total, 60)
    return f"{minutes:02d}:{seconds:02d}"


def create_operation_logger(operation: str) -> tuple[logging.Logger, Path]:
    base = Path(os.environ.get("LOCALAPPDATA", str(APP_DIR))) / "변분기" / "logs"
    try:
        base.mkdir(parents=True, exist_ok=True)
    except OSError:
        base = APP_DIR / "logs"
        base.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    path = base / f"변분기_{operation}_{stamp}.log"
    logger = logging.getLogger(f"byeonbungi.{uuid.uuid4().hex}")
    logger.setLevel(logging.INFO)
    logger.propagate = False
    handler = logging.FileHandler(path, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(handler)
    return logger, path


def close_operation_logger(logger: logging.Logger) -> None:
    for handler in list(logger.handlers):
        handler.close()
        logger.removeHandler(handler)


@dataclass
class ChangeGroup:
    section: str
    page_start: int
    page_end: int
    range_start: int
    range_end: int
    summary: str
    context_type: str = "paragraph"
    omit_before: bool = False
    omit_after: bool = False

    @property
    def page_label(self) -> str:
        if self.page_start == self.page_end:
            return str(self.page_start)
        return f"{self.page_start}-{self.page_end}"


@dataclass(frozen=True)
class XmlRevisionRecord:
    revision_index: int
    paragraph_index: int
    section_id: int
    section_hint: str
    change_type: str
    summary: str
    table_id: int = 0
    table_row: int = 0


@dataclass
class XmlChangeGroup:
    section_id: int
    section_hint: str
    revision_start: int
    revision_end: int
    paragraph_start: int
    paragraph_end: int
    summary: str
    context_type: str = "paragraph"
    table_id: int = 0
    table_row_start: int = 0
    table_row_end: int = 0


@dataclass(frozen=True)
class XmlBodyScope:
    paragraph_start: int
    paragraph_end: int
    toc_detected: bool
    appendix_detected: bool
    start_label: str
    end_label: str


def _xml_local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


COMMENT_MARKER_NAMES = {"commentRangeStart", "commentRangeEnd", "commentReference"}
COMMENT_PART_PREFIXES = (
    "word/comments",
    "word/people.xml",
)


def _remove_xml_elements(root: ET.Element, names: set[str]) -> int:
    parent_map = {child: parent for parent in root.iter() for child in parent}
    targets = [node for node in root.iter() if _xml_local_name(node.tag) in names]
    for node in targets:
        parent = parent_map.get(node)
        if parent is not None:
            parent.remove(node)
    return len(targets)


def create_comment_free_copy(input_path: Path, output_path: Path) -> int:
    """Create a comment-free copy without reserializing WordprocessingML.

    Word documents often declare extension namespaces only on the document
    root. ElementTree reserialization can discard those declarations while
    leaving mc:Ignorable references behind, which makes Word report a corrupt
    file. Alpha 0.8 therefore removes only comment-related byte ranges and
    leaves all remaining XML bytes, prefixes and declarations untouched.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    comment_count = 0
    comment_references = 0
    with zipfile.ZipFile(input_path, "r") as source, zipfile.ZipFile(
        output_path, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for info in source.infolist():
            name = info.filename
            lower_name = name.lower()
            data = source.read(name)

            if any(lower_name.startswith(prefix) for prefix in COMMENT_PART_PREFIXES):
                if lower_name == "word/comments.xml":
                    with suppress(Exception):
                        comments_root = ET.fromstring(data)
                        comment_count += sum(
                            1
                            for node in comments_root.iter()
                            if _xml_local_name(node.tag) == "comment"
                        )
                continue

            if lower_name.endswith(".rels") and (
                b"comment" in data.lower() or b"people" in data.lower()
            ):
                data = re.sub(
                    rb"<Relationship\b(?=[^>]*(?:comments?|people\.xml))[^>]*/>",
                    b"",
                    data,
                    flags=re.I,
                )

            elif lower_name == "[content_types].xml" and (
                b"comment" in data.lower() or b"people.xml" in data.lower()
            ):
                data = re.sub(
                    rb"<(?:Override|Default)\b(?=[^>]*(?:comments?|people\.xml))[^>]*/>",
                    b"",
                    data,
                    flags=re.I,
                )

            elif lower_name.startswith("word/") and lower_name.endswith(".xml") and (
                b"commentRange" in data or b"commentReference" in data
            ):
                patterns = (
                    rb"<w:commentRangeStart\b[^>]*/>",
                    rb"<w:commentRangeEnd\b[^>]*/>",
                    rb"<w:commentReference\b[^>]*/>",
                )
                for pattern in patterns:
                    data, removed = re.subn(pattern, b"", data, flags=re.I)
                    comment_references += removed

            target.writestr(info, data)
    return max(comment_count, comment_references // 3)


def _xml_text(element: ET.Element) -> str:
    parts: list[str] = []
    for node in element.iter():
        if _xml_local_name(node.tag) in {"t", "delText", "instrText", "delInstrText"}:
            if node.text:
                parts.append(node.text)
    return clean_word_text("".join(parts))


def _heading_style_ids(styles_root: ET.Element | None) -> set[str]:
    if styles_root is None:
        return set()
    result: set[str] = set()
    for style in styles_root.findall("w:style", W):
        if style.get(f"{{{W_NS}}}type") != "paragraph":
            continue
        style_id = style.get(f"{{{W_NS}}}styleId", "")
        name_node = style.find("w:name", W)
        name = name_node.get(W_VAL, "") if name_node is not None else ""
        outline = style.find("w:pPr/w:outlineLvl", W)
        outline_value = outline.get(W_VAL, "9") if outline is not None else "9"
        is_outline = outline_value.isdigit() and int(outline_value) < 9
        if is_outline or re.search(r"heading|제목|개요", name, re.I):
            result.add(style_id)
    return result


def _style_names(styles_root: ET.Element | None) -> dict[str, str]:
    if styles_root is None:
        return {}
    result: dict[str, str] = {}
    for style in styles_root.findall("w:style", W):
        style_id = style.get(f"{{{W_NS}}}styleId", "")
        name_node = style.find("w:name", W)
        if style_id and name_node is not None:
            result[style_id] = name_node.get(W_VAL, "")
    return result


TOC_TITLE_RE = re.compile(
    r"^(?:목\s*차|차\s*례|table\s+of\s+contents|list\s+of\s+contents|"
    r"list\s+of\s+tables?|contents)$",
    re.I,
)
APPENDIX_TITLE_RE = re.compile(
    r"^\s*(?:\[\s*)?(?:제\s*)?(?:\d+(?:\.\d+)*\s*)?"
    r"(?:부록|붙임|별첨|첨부|appendix|appendices|annex)"
    r"(?=\s|$|[\]\[\dA-Z가-힣:：.\-])",
    re.I,
)


def _detect_body_scope_from_roots(
    document_root: ET.Element,
    styles_root: ET.Element | None,
) -> XmlBodyScope:
    paragraphs = document_root.findall(".//w:body//w:p", W)
    paragraph_count = len(paragraphs)
    if not paragraphs:
        return XmlBodyScope(0, 0, False, False, "문서 시작", "문서 끝")

    heading_styles = _heading_style_ids(styles_root)
    style_names = _style_names(styles_root)
    toc_style_ids = {
        style_id
        for style_id, name in style_names.items()
        if re.search(r"(?:^|\s)toc(?:\s|$)|table of figures|목차|차례", name, re.I)
    }
    paragraph_texts = [_xml_text(paragraph) for paragraph in paragraphs]

    # Detect complete TOC fields, including a separate list of tables. A TOC
    # field can contain many nested PAGEREF fields, so its outer field depth is
    # tracked until the matching end marker.
    toc_ranges: list[tuple[int, int]] = []
    field_depth = 0
    active_toc: tuple[int, int] | None = None
    fld_type_attr = f"{{{W_NS}}}fldCharType"
    for paragraph_index, paragraph in enumerate(paragraphs):
        for node in paragraph.iter():
            local = _xml_local_name(node.tag)
            if local == "fldChar":
                field_type = node.get(fld_type_attr, "")
                if field_type == "begin":
                    field_depth += 1
                elif field_type == "end":
                    if active_toc is not None and field_depth == active_toc[1]:
                        toc_ranges.append((active_toc[0], paragraph_index + 1))
                        active_toc = None
                    field_depth = max(0, field_depth - 1)
            elif local == "instrText" and node.text:
                if re.search(r"(?:^|\s)TOC(?:\s|\\)", node.text, re.I):
                    active_toc = (paragraph_index, max(1, field_depth))
    if active_toc is not None:
        toc_ranges.append((active_toc[0], paragraph_count))

    # Word can also store a TOC in a content control rather than a field.
    paragraph_index_by_id = {id(paragraph): index for index, paragraph in enumerate(paragraphs)}
    for node in document_root.iter():
        if _xml_local_name(node.tag) != "sdt":
            continue
        gallery_values = [
            str(child.get(W_VAL, ""))
            for child in node.iter()
            if _xml_local_name(child.tag) == "docPartGallery"
        ]
        if not any(re.search(r"table of contents|목차|차례", value, re.I) for value in gallery_values):
            continue
        indexes = [
            paragraph_index_by_id[id(child)]
            for child in node.iter()
            if id(child) in paragraph_index_by_id
        ]
        if indexes:
            toc_ranges.append((min(indexes), max(indexes) + 1))

    body_start = 0
    toc_detected = False
    if toc_ranges:
        toc_ranges.sort()
        first_start, cluster_end = toc_ranges[0]
        # Treat TOCs as front matter only when the first one occurs in the first
        # third of the document. Include adjacent lists of tables/figures.
        if first_start <= max(100, paragraph_count // 3):
            toc_detected = True
            for start, end in toc_ranges[1:]:
                if start - cluster_end > 25:
                    break
                cluster_end = max(cluster_end, end)
            body_start = cluster_end

    if not toc_detected:
        # Fallback for manually typed TOCs without a Word field.
        toc_title_index = next(
            (
                index
                for index, text in enumerate(paragraph_texts[: max(100, paragraph_count // 3)])
                if TOC_TITLE_RE.match(clean_word_text(text))
            ),
            None,
        )
        if toc_title_index is not None:
            toc_detected = True
            body_start = toc_title_index + 1
            for index in range(body_start, paragraph_count):
                paragraph = paragraphs[index]
                style_node = paragraph.find("w:pPr/w:pStyle", W)
                style_id = style_node.get(W_VAL, "") if style_node is not None else ""
                instructions = "".join(
                    child.text or "" for child in paragraph.findall(".//w:instrText", W)
                )
                text = paragraph_texts[index]
                if (
                    not text
                    or style_id in toc_style_ids
                    or re.search(r"\bPAGEREF\b|(?:^|\s)TOC(?:\s|\\)", instructions, re.I)
                    or TOC_TITLE_RE.match(clean_word_text(text))
                ):
                    body_start = index + 1
                    continue
                break

    while body_start < paragraph_count and not paragraph_texts[body_start]:
        body_start += 1

    body_end = paragraph_count
    appendix_detected = False
    end_label = "문서 끝"
    for index in range(body_start, paragraph_count):
        text = clean_word_text(paragraph_texts[index])
        if not text or not APPENDIX_TITLE_RE.match(text):
            continue
        heading = _paragraph_heading(paragraphs[index], text, heading_styles)
        if heading is not None or text.startswith("[") or re.match(
            r"^(?:부록|붙임|별첨|appendix|appendices|annex)(?:\s|$)", text, re.I
        ):
            body_end = index
            appendix_detected = True
            end_label = text[:120]
            break

    start_label = (
        paragraph_texts[body_start][:120]
        if body_start < paragraph_count and paragraph_texts[body_start]
        else "문서 시작"
    )
    return XmlBodyScope(
        paragraph_start=body_start,
        paragraph_end=max(body_start, body_end),
        toc_detected=toc_detected,
        appendix_detected=appendix_detected,
        start_label=start_label,
        end_label=end_label,
    )


def detect_xml_body_scope(input_path: Path) -> XmlBodyScope:
    with zipfile.ZipFile(input_path) as archive:
        document_root = ET.fromstring(archive.read("word/document.xml"))
        styles_root = (
            ET.fromstring(archive.read("word/styles.xml"))
            if "word/styles.xml" in archive.namelist()
            else None
        )
    return _detect_body_scope_from_roots(document_root, styles_root)


def _paragraph_heading(paragraph: ET.Element, text: str, heading_styles: set[str]) -> str | None:
    if not text:
        return None
    style = paragraph.find("w:pPr/w:pStyle", W)
    style_id = style.get(W_VAL, "") if style is not None else ""
    outline = paragraph.find("w:pPr/w:outlineLvl", W)
    outline_value = outline.get(W_VAL, "9") if outline is not None else "9"
    if style_id in heading_styles or (outline_value.isdigit() and int(outline_value) < 9):
        return text[:160]
    if len(text) <= 140 and (
        re.match(r"^\d+(?:\.\d+){0,3}\s+\S", text)
        or re.match(r"^(?:제\s*\d+\s*[장절]|Reference\b)", text, re.I)
    ):
        return text[:160]
    return None


def _outer_revision_nodes(paragraph: ET.Element) -> list[ET.Element]:
    result: list[ET.Element] = []

    def visit(node: ET.Element, within_revision: bool = False) -> None:
        local = _xml_local_name(node.tag)
        is_revision = local in XML_REVISION_TAGS
        if is_revision and not within_revision:
            result.append(node)
        for child in node:
            visit(child, within_revision or is_revision)

    visit(paragraph)
    return result


def _classify_xml_change(raw: dict[str, Any]) -> dict[str, Any]:
    labels = {
        "ins": "추가",
        "del": "삭제",
        "moveFrom": "이동 전",
        "moveTo": "이동 후",
    }
    change_type = str(raw["change_type"])
    text = clean_word_text(str(raw.get("text") or ""))
    if not text:
        text = clean_word_text(str(raw.get("paragraph_text") or ""))
    summary = text[:100] if text else f"{labels.get(change_type, '서식/구조')} 변경"
    return {**raw, "summary": summary}


def extract_xml_revisions(
    input_path: Path,
    body_scope: XmlBodyScope | None = None,
) -> list[XmlRevisionRecord]:
    with zipfile.ZipFile(input_path) as archive:
        document_root = ET.fromstring(archive.read("word/document.xml"))
        styles_root = (
            ET.fromstring(archive.read("word/styles.xml"))
            if "word/styles.xml" in archive.namelist()
            else None
        )

    heading_styles = _heading_style_ids(styles_root)
    paragraphs = document_root.findall(".//w:body//w:p", W)
    if body_scope is None:
        body_scope = _detect_body_scope_from_roots(document_root, styles_root)
    element_order = {id(element): index for index, element in enumerate(document_root.iter())}
    parent_map = {child: parent for parent in document_root.iter() for child in parent}
    tables = [node for node in document_root.iter() if _xml_local_name(node.tag) == "tbl"]
    table_ids = {id(node): index for index, node in enumerate(tables, start=1)}

    def table_location(node: ET.Element) -> tuple[int, int]:
        row = None
        current = node
        while current is not None:
            local = _xml_local_name(current.tag)
            if local == "tr" and row is None:
                row = current
            if local == "tbl":
                rows = [child for child in current if _xml_local_name(child.tag) == "tr"]
                row_number = rows.index(row) + 1 if row in rows else 0
                return table_ids.get(id(current), 0), row_number
            current = parent_map.get(current)
        return 0, 0
    paragraph_orders = [element_order[id(paragraph)] for paragraph in paragraphs]
    paragraph_sections: list[tuple[int, str]] = []
    section_id = 0
    section_hint = "Section 확인 필요"
    for paragraph in paragraphs:
        paragraph_text = _xml_text(paragraph)
        heading = _paragraph_heading(paragraph, paragraph_text, heading_styles)
        if heading:
            section_id += 1
            section_hint = heading
        paragraph_sections.append((section_id, section_hint))

    raw_items: list[dict[str, Any]] = []
    seen: set[int] = set()
    for paragraph_index, paragraph in enumerate(paragraphs):
        if not (
            body_scope.paragraph_start
            <= paragraph_index
            < body_scope.paragraph_end
        ):
            continue
        paragraph_text = _xml_text(paragraph)
        current_section_id, current_section = paragraph_sections[paragraph_index]
        table_id, table_row = table_location(paragraph)
        for descendant in paragraph.iter():
            if _xml_local_name(descendant.tag) in XML_REVISION_TAGS:
                seen.add(id(descendant))
        for node in _outer_revision_nodes(paragraph):
            raw_items.append({
                "order": element_order[id(node)],
                "paragraph_index": paragraph_index,
                "section_id": current_section_id,
                "section_hint": current_section,
                "change_type": _xml_local_name(node.tag),
                "text": _xml_text(node),
                "paragraph_text": paragraph_text,
                "table_id": table_id,
                "table_row": table_row,
            })

    # Include tracked table/cell/property changes that are outside a paragraph.
    for node in document_root.iter():
        if id(node) in seen or _xml_local_name(node.tag) not in XML_REVISION_TAGS:
            continue
        ancestor = parent_map.get(node)
        nested = False
        while ancestor is not None:
            if _xml_local_name(ancestor.tag) in XML_REVISION_TAGS:
                nested = True
                break
            ancestor = parent_map.get(ancestor)
        if nested:
            continue
        order = element_order[id(node)]
        paragraph_index = max(0, bisect.bisect_right(paragraph_orders, order) - 1)
        if not (
            body_scope.paragraph_start
            <= paragraph_index
            < body_scope.paragraph_end
        ):
            continue
        current_section_id, current_section = (
            paragraph_sections[paragraph_index]
            if paragraph_sections
            else (0, "Section 확인 필요")
        )
        table_id, table_row = table_location(node)
        raw_items.append({
            "order": order,
            "paragraph_index": paragraph_index,
            "section_id": current_section_id,
            "section_hint": current_section,
            "change_type": _xml_local_name(node.tag),
            "text": _xml_text(node),
            "paragraph_text": "",
            "table_id": table_id,
            "table_row": table_row,
        })

    raw_items.sort(key=lambda item: int(item["order"]))
    workers = min(4, max(1, os.cpu_count() or 1), max(1, len(raw_items)))
    with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="xml-classify") as executor:
        classified = list(executor.map(_classify_xml_change, raw_items))
    return [
        XmlRevisionRecord(
            revision_index=index,
            paragraph_index=int(item["paragraph_index"]),
            section_id=int(item["section_id"]),
            section_hint=str(item["section_hint"]),
            change_type=str(item["change_type"]),
            summary=str(item["summary"]),
            table_id=int(item.get("table_id", 0)),
            table_row=int(item.get("table_row", 0)),
        )
        for index, item in enumerate(classified, start=1)
    ]


def group_xml_revisions(records: list[XmlRevisionRecord]) -> list[XmlChangeGroup]:
    """Group revisions by reviewable topic, never by an entire Section.

    Paragraph changes merge only when separated by at most one unchanged
    paragraph. Table changes merge only inside the same table and when their
    changed rows are adjacent. This is the central Alpha 0.8 rules engine.
    """
    groups: list[XmlChangeGroup] = []
    for record in records:
        context_type = "table" if record.table_id else "paragraph"
        can_merge = False
        if groups:
            previous = groups[-1]
            if context_type == "table":
                can_merge = (
                    previous.context_type == "table"
                    and previous.table_id == record.table_id
                    and record.table_row - previous.table_row_end <= 1
                )
            else:
                can_merge = (
                    previous.context_type == "paragraph"
                    and record.section_id == previous.section_id
                    and record.paragraph_index - previous.paragraph_end <= 2
                )
        if not can_merge:
            groups.append(XmlChangeGroup(
                section_id=record.section_id,
                section_hint=record.section_hint,
                revision_start=record.revision_index,
                revision_end=record.revision_index,
                paragraph_start=record.paragraph_index,
                paragraph_end=record.paragraph_index,
                summary=record.summary,
                context_type=context_type,
                table_id=record.table_id,
                table_row_start=record.table_row,
                table_row_end=record.table_row,
            ))
            continue
        group = groups[-1]
        group.revision_end = record.revision_index
        group.paragraph_end = record.paragraph_index
        if record.table_row:
            group.table_row_end = max(group.table_row_end, record.table_row)
        if record.summary and record.summary not in group.summary:
            group.summary = clean_word_text(f"{group.summary} / {record.summary}")[:180]
    return groups


def clean_word_text(value: str) -> str:
    value = value.replace("\r", " ").replace("\x07", " ").replace("\v", " ")
    return re.sub(r"\s+", " ", value).strip()


def detect_metadata_from_name(path: Path) -> dict[str, str]:
    stem = path.stem
    stem = re.sub(r"(?i)(?:_|\s)*(?:draft[-_ ]?tracked|tracked|변경추적)$", "", stem)
    return {"document_name": stem.replace("_", " ").strip()}


def _word_range_bounds(rng) -> tuple[int, int]:
    try:
        if bool(rng.Information(WD_WITHIN_TABLE)) and rng.Rows.Count:
            rng = rng.Rows(1).Range.Duplicate
        elif rng.Paragraphs.Count:
            rng = rng.Paragraphs(1).Range.Duplicate
    except Exception:
        if rng.Paragraphs.Count:
            rng = rng.Paragraphs(1).Range.Duplicate
    return int(rng.Start), int(rng.End)


def _revision_range_bounds(revision) -> tuple[int, int]:
    return _word_range_bounds(revision.Range.Duplicate)


def _word_body_bounds(document, scope: XmlBodyScope) -> tuple[int, int]:
    paragraph_count = int(document.Paragraphs.Count)
    content_end = max(0, int(document.Content.End) - 1)
    if paragraph_count <= 0:
        return 0, content_end
    start_index = min(paragraph_count, max(1, scope.paragraph_start + 1))
    start = int(document.Paragraphs(start_index).Range.Start)
    if scope.paragraph_end < paragraph_count:
        end = int(document.Paragraphs(scope.paragraph_end + 1).Range.Start)
    else:
        end = content_end
    return min(start, end), max(start, end)


def _nearby_text(document, position: int, before: int = 5000, after: int = 500) -> str:
    start = max(0, position - before)
    end = min(int(document.Content.End) - 1, position + after)
    return clean_word_text(document.Range(start, end).Text)


def _guess_section(document, position: int, page: int) -> str:
    nearby = _nearby_text(document, position)
    tail = nearby[-1800:]

    if page <= 2 and any(token in tail for token in ("회사명", "허가(신청)", "위해성 관리계획 번호")):
        return "위해성 관리 계획 개요"
    if page <= 4 and "위해성 관리 계획 변경이력" in tail:
        return "위해성 관리 계획 변경이력"
    if page <= 10 and any(token in tail for token in ("버전(작성일)", "Version No.", "대상 의약품")):
        return "표지"
    if re.search(r"(?:^|\s)Reference(?:\s|$)", tail, re.I):
        return "Reference"

    # Prefer real Word outline headings when the source uses heading styles.
    search_start = max(0, position - 30000)
    paragraphs = document.Range(search_start, position).Paragraphs
    for index in range(paragraphs.Count, 0, -1):
        paragraph = paragraphs(index)
        text = clean_word_text(paragraph.Range.Text)
        if not text:
            continue
        try:
            outline_level = int(paragraph.OutlineLevel)
        except Exception:
            outline_level = WD_OUTLINE_LEVEL_BODY_TEXT
        if outline_level < WD_OUTLINE_LEVEL_BODY_TEXT:
            return text[:160]

    # Fallback for documents whose headings are directly formatted.
    candidates = re.findall(
        r"(?:^|\s)(\d+(?:\.\d+){0,3})\s+([^\r\n]{3,120}?)(?=(?:\s\d+(?:\.\d+){0,3}\s)|$)",
        tail,
    )
    if candidates:
        number, title = candidates[-1]
        return clean_word_text(f"{number} {title}")[:160]
    return "Section 확인 필요"


def _summary_for_revision(revision) -> str:
    text = clean_word_text(revision.Range.Text)
    if not text:
        text = "표/문단 구조 변경"
    return text[:100]


def _merge_by_section(items: list[dict]) -> list[ChangeGroup]:
    """Compatibility grouping for Word-only analysis using local proximity."""
    groups: list[ChangeGroup] = []
    for item in sorted(items, key=lambda value: value["start"]):
        can_merge = bool(
            groups
            and groups[-1].section == item["section"]
            and item["start"] - groups[-1].range_end <= 800
            and item["page"] - groups[-1].page_end <= 1
        )
        if not can_merge:
            group = ChangeGroup(
                section=item["section"],
                page_start=item["page"],
                page_end=item["page"],
                range_start=item["start"],
                range_end=item["end"],
                summary=item["summary"],
            )
            groups.append(group)
        else:
            group = groups[-1]
            group.page_start = min(group.page_start, item["page"])
            group.page_end = max(group.page_end, item["page"])
            group.range_start = min(group.range_start, item["start"])
            group.range_end = max(group.range_end, item["end"])
            if item["summary"] not in group.summary:
                group.summary = clean_word_text(group.summary + " / " + item["summary"])[:180]
    return groups


def _coalesce_change_groups(items: list[ChangeGroup]) -> list[ChangeGroup]:
    # Alpha 0.7 coalesced every item sharing a Section, producing page-spanning
    # cells. XML grouping in Alpha 0.8 already represents a reviewable topic.
    return sorted(items, key=lambda group: group.range_start)


class WordEngine:
    def __init__(
        self,
        progress: Callable[[dict[str, Any]], None] | None = None,
        cancel_event: threading.Event | None = None,
        logger: logging.Logger | None = None,
    ) -> None:
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise RuntimeError("pywin32가 설치되지 않았습니다. INSTALL.bat을 먼저 실행하십시오.") from exc
        self.pythoncom = pythoncom
        self.win32com = win32com.client
        self.progress = progress or (lambda _event: None)
        self.cancel_event = cancel_event or threading.Event()
        self.logger = logger
        self.word_pid: int | None = None

    def _check_cancelled(self) -> None:
        if self.cancel_event.is_set():
            raise OperationCancelled("사용자가 작업을 취소했습니다.")

    @staticmethod
    def _exception_hresult(exc: Exception) -> int | None:
        value = getattr(exc, "hresult", None)
        if value is None and getattr(exc, "args", None):
            value = exc.args[0]
        if not isinstance(value, int):
            return None
        if value > 0x7FFFFFFF:
            value -= 0x100000000
        return value

    def _wait_for_word_retry(self, seconds: float, check_cancel: bool) -> None:
        with suppress(Exception):
            self.pythoncom.PumpWaitingMessages()
        if check_cancel:
            if self.cancel_event.wait(seconds):
                raise OperationCancelled("사용자가 작업을 취소했습니다.")
        else:
            time.sleep(seconds)

    def _retry_word_call(
        self,
        operation: str,
        action: Callable[[], Any],
        *,
        attempts: int = WORD_CALL_RETRY_ATTEMPTS,
        check_cancel: bool = True,
        progress_info: tuple[str, float, int | None, int | None] | None = None,
    ) -> Any:
        last_error: Exception | None = None
        for attempt in range(1, max(1, attempts) + 1):
            if check_cancel:
                self._check_cancelled()
            try:
                return action()
            except Exception as exc:
                last_error = exc
                hresult = self._exception_hresult(exc)
                if hresult not in RETRYABLE_WORD_HRESULTS or attempt >= attempts:
                    raise
                delay = min(
                    WORD_CALL_RETRY_MAX_DELAY,
                    0.2 * (1.5 ** (attempt - 1)),
                )
                detail = (
                    f"Word가 다른 작업을 처리 중입니다. {operation} 자동 재시도 "
                    f"{attempt}/{attempts}"
                )
                if self.logger:
                    self.logger.warning("%s · %.1f초 후 재시도 · HRESULT %s", detail, delay, hresult)
                if progress_info is not None and check_cancel:
                    stage, percent, current, total = progress_info
                    self._report(stage, percent, detail, current, total)
                self._wait_for_word_retry(delay, check_cancel)
        if last_error is not None:
            raise last_error
        raise RuntimeError(f"Word 작업을 실행하지 못했습니다: {operation}")

    def _report(
        self,
        stage: str,
        percent: float,
        detail: str = "",
        current: int | None = None,
        total: int | None = None,
    ) -> None:
        self._check_cancelled()
        event = {
            "kind": "progress",
            "stage": stage,
            "percent": max(0.0, min(100.0, float(percent))),
            "detail": detail,
            "current": current,
            "total": total,
            "word_pid": self.word_pid,
        }
        if self.logger:
            suffix = f" ({current}/{total})" if current is not None and total else ""
            self.logger.info("%s %.1f%%%s %s", stage, percent, suffix, detail)
        self.progress(event)

    @staticmethod
    def _word_process_id(word) -> int | None:
        if sys.platform != "win32":
            return None
        try:
            import ctypes

            process_id = ctypes.c_ulong()
            ctypes.windll.user32.GetWindowThreadProcessId(int(word.Hwnd), ctypes.byref(process_id))
            return int(process_id.value) or None
        except Exception:
            return None

    def _start_word(self, stage: str, percent: float):
        self._report(stage, percent, "Microsoft Word 전용 프로세스를 시작하는 중")
        self.pythoncom.CoInitialize()
        word = self.win32com.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        with suppress(Exception):
            word.AutomationSecurity = 3  # Disable macros for automation-opened files.
        with suppress(Exception):
            word.Options.UpdateLinksAtOpen = False
        with suppress(Exception):
            word.ScreenUpdating = False
        self.word_pid = self._word_process_id(word)
        self._report(stage, percent + 2, "Microsoft Word 시작 완료")
        return word

    def _close_document(self, document, label: str = "문서") -> None:
        if document is None:
            return
        try:
            self._retry_word_call(
                f"{label} 닫기",
                lambda: document.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES),
                attempts=8,
                check_cancel=False,
            )
        except Exception as exc:
            if self.logger:
                self.logger.warning("%s 닫기 실패: %s", label, exc)

    def _close_word(self, document, word) -> None:
        closing_pid = self.word_pid
        self._close_document(document, "작업 문서")
        quit_succeeded = word is None
        if word is not None:
            try:
                self._retry_word_call(
                    "Microsoft Word 종료",
                    lambda: word.Quit(),
                    attempts=8,
                    check_cancel=False,
                )
                quit_succeeded = True
            except Exception as exc:
                if self.logger:
                    self.logger.warning("Microsoft Word 정상 종료 실패: %s", exc)
        if not quit_succeeded and closing_pid and sys.platform == "win32":
            # This PID belongs to the dedicated Word instance created above.
            with suppress(Exception):
                subprocess.run(
                    ["taskkill", "/PID", str(closing_pid), "/T", "/F"],
                    check=False,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
        document = None
        word = None
        gc.collect()
        time.sleep(0.2)
        with suppress(Exception):
            self.pythoncom.CoUninitialize()
        if closing_pid:
            self.progress({"kind": "word_closed", "word_pid": closing_pid})
        self.word_pid = None

    def analyze(self, input_path: Path) -> tuple[list[ChangeGroup], dict[str, str]]:
        word = None
        document = None
        # Analysis and generation both use the same deterministic, comment-free
        # copy so Word character positions remain aligned. The user's source is
        # never edited.
        with tempfile.TemporaryDirectory(
            prefix="change_analysis_", ignore_cleanup_errors=True
        ) as temp_dir_name:
            working_path = Path(temp_dir_name) / "source_without_comments.docx"
            try:
                self._report("1/7 메모 제거", 1, "분석용 복사본을 준비하는 중")
                comment_count = create_comment_free_copy(input_path, working_path)
                self._report(
                    "1/7 메모 제거",
                    5,
                    f"메모 {comment_count}개 제거 완료 · 원본 파일은 변경하지 않음",
                )

                body_scope = detect_xml_body_scope(working_path)
                boundary_detail = (
                    f"본문 시작: {body_scope.start_label} · "
                    f"본문 끝: {body_scope.end_label}"
                )
                self._report("2/7 본문 범위 확인", 7, boundary_detail)

                self._report("2/7 XML 변경기록 추출", 9, working_path.name)
                xml_records = extract_xml_revisions(working_path, body_scope)
                if not xml_records:
                    raise ValueError(
                        "목차와 부록을 제외한 본문에서 변경 추적 기록을 찾지 못했습니다."
                    )
                self._report(
                    "2/7 XML 변경기록 추출",
                    13,
                    f"본문 XML 변경기록 {len(xml_records)}개 추출 완료",
                    len(xml_records),
                    len(xml_records),
                )
                self._report(
                    "3/7 Python 병렬 분류",
                    16,
                    f"최대 {min(4, max(1, os.cpu_count() or 1))}개 작업 스레드로 분류 완료",
                )
                xml_groups = group_xml_revisions(xml_records)
                self._report(
                    "4/7 인접 변경 묶기",
                    19,
                    f"본문 변경기록 {len(xml_records)}개 → {len(xml_groups)}개 묶음",
                )

                word = self._start_word("5/7 Word 시작", 21)
                self._report("5/7 문서 열기", 24, working_path.name)
                document = word.Documents.Open(
                    str(working_path),
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    Visible=False,
                    OpenAndRepair=False,
                    NoEncodingDialog=True,
                )
                self._report("5/7 문서 열기", 27, "메모가 제거된 분석용 문서 열기 완료")

                revision_count = int(document.Revisions.Count)
                if revision_count == 0:
                    raise ValueError("이 문서에는 Word 변경 추적 기록이 없습니다.")
                self._report(
                    "6/7 페이지 계산",
                    29,
                    f"Word 변경기록 {revision_count}개 · 페이지를 다시 계산하는 중",
                    0,
                    revision_count,
                )
                document.Repaginate()
                self._report("6/7 페이지 계산", 32, "페이지 계산 완료", 0, revision_count)

                metadata = detect_metadata_from_name(input_path)
                paragraph_count = int(document.Paragraphs.Count)
                max_paragraph_index = max(group.paragraph_end for group in xml_groups)
                if paragraph_count > max_paragraph_index:
                    groups = self._resolve_xml_groups(document, xml_groups)
                else:
                    if self.logger:
                        self.logger.warning(
                            "XML/Word 문단 수 불일치: XML 최대 인덱스=%s Word 문단=%s; Word 호환 분석으로 전환",
                            max_paragraph_index,
                            paragraph_count,
                        )
                    self._report(
                        "6/7 Word 호환 분석",
                        33,
                        "XML/Word 문단 구조 불일치 · 본문 범위 내에서 호환 분석 실행",
                    )
                    body_bounds = _word_body_bounds(document, body_scope)
                    groups = self._analyze_word_fallback(
                        document, revision_count, body_bounds=body_bounds
                    )
                if not groups:
                    raise ValueError(
                        "목차와 부록을 제외한 본문에서 변경 묶음을 만들지 못했습니다."
                    )
                self._report(
                    "7/7 분석 마무리",
                    99,
                    f"본문 변경 묶음 {len(groups)}개 · Word 종료 중",
                )
                return groups, metadata
            finally:
                self._close_word(document, word)

    def _resolve_xml_groups(
        self, document, xml_groups: list[XmlChangeGroup]
    ) -> list[ChangeGroup]:
        groups: list[ChangeGroup] = []
        total = len(xml_groups)
        for index, xml_group in enumerate(xml_groups, start=1):
            self._check_cancelled()
            first_paragraph_number = xml_group.paragraph_start + 1
            last_paragraph_number = xml_group.paragraph_end + 1
            first_range = document.Paragraphs(first_paragraph_number).Range.Duplicate
            last_range = document.Paragraphs(last_paragraph_number).Range.Duplicate
            page_start = int(first_range.Information(WD_ACTIVE_END_ADJUSTED_PAGE_NUMBER))
            page_end = int(last_range.Information(WD_ACTIVE_END_ADJUSTED_PAGE_NUMBER))
            start, _ = _word_range_bounds(first_range)
            _, end = _word_range_bounds(last_range)
            context_type = xml_group.context_type
            omit_before = False
            omit_after = False

            if context_type == "table":
                # Keep changed table rows and at most one adjacent row on each
                # side. Large unchanged tables are intentionally not copied.
                with suppress(Exception):
                    table = first_range.Tables(1)
                    first_row = int(first_range.Cells(1).RowIndex)
                    last_row = int(last_range.Cells(1).RowIndex)
                    row_count = int(table.Rows.Count)
                    display_first = max(1, first_row - 1)
                    display_last = min(row_count, last_row + 1)
                    start = int(table.Rows(display_first).Range.Start)
                    end = int(table.Rows(display_last).Range.End)
                    omit_before = display_first > 1
                    omit_after = display_last < row_count
            else:
                # Include one context paragraph before and after only when the
                # resulting excerpt remains compact (about 800 characters).
                candidate_first = max(1, first_paragraph_number - 1)
                candidate_last = min(int(document.Paragraphs.Count), last_paragraph_number + 1)
                candidate_start = int(document.Paragraphs(candidate_first).Range.Start)
                candidate_end = int(document.Paragraphs(candidate_last).Range.End)
                if candidate_end - candidate_start <= 800:
                    start, end = candidate_start, candidate_end
                omit_before = candidate_first > 1
                omit_after = candidate_last < int(document.Paragraphs.Count)
            word_section = _guess_section(document, int(first_range.Start), page_start)
            section = (
                word_section
                if word_section != "Section 확인 필요"
                else xml_group.section_hint
            )
            groups.append(ChangeGroup(
                section=section,
                page_start=min(page_start, page_end),
                page_end=max(page_start, page_end),
                range_start=min(start, end),
                range_end=max(start, end),
                summary=xml_group.summary,
                context_type=context_type,
                omit_before=omit_before,
                omit_after=omit_after,
            ))
            page_label = str(page_start) if page_start == page_end else f"{page_start}-{page_end}"
            self._report(
                "6/7 묶음별 page·Section 조회",
                32 + (65 * index / max(1, total)),
                f"page {page_label} · {section}",
                index,
                total,
            )
        return _coalesce_change_groups(groups)

    def _analyze_word_fallback(
        self,
        document,
        revision_count: int,
        body_bounds: tuple[int, int] | None = None,
    ) -> list[ChangeGroup]:
        items: list[dict[str, Any]] = []
        section_by_page: dict[int, str] = {}
        report_step = max(1, revision_count // 200)
        for index in range(1, revision_count + 1):
            self._check_cancelled()
            revision = document.Revisions(index)
            revision_range = revision.Range.Duplicate
            revision_type = int(revision.Type)
            start, end = int(revision_range.Start), int(revision_range.End)
            if body_bounds is not None:
                body_start, body_end = body_bounds
                if end <= body_start or start >= body_end:
                    if index == 1 or index == revision_count or index % report_step == 0:
                        self._report(
                            "6/7 Word 호환 분석",
                            33 + (64 * index / revision_count),
                            "목차·부록 등 본문 범위 밖 변경 제외 중",
                            index,
                            revision_count,
                        )
                    continue
            page = int(revision_range.Information(WD_ACTIVE_END_ADJUSTED_PAGE_NUMBER))
            if page not in section_by_page:
                section_by_page[page] = _guess_section(document, start, page)
            section = section_by_page[page]
            text = clean_word_text(revision_range.Text)
            items.append({
                "type": revision_type,
                "page": page,
                "start": start,
                "end": end,
                "section": section,
                "summary": (text or "표/문단 구조 변경")[:100],
            })
            if index == 1 or index == revision_count or index % report_step == 0:
                self._report(
                    "6/7 Word 호환 분석",
                    33 + (64 * index / revision_count),
                    f"page {page} · {section}",
                    index,
                    revision_count,
                )
        groups = _merge_by_section(items)
        for group in groups:
            first = document.Range(group.range_start, group.range_start).Duplicate
            last = document.Range(group.range_end, group.range_end).Duplicate
            group.range_start = _word_range_bounds(first)[0]
            group.range_end = _word_range_bounds(last)[1]
        return groups

    @staticmethod
    def _bookmark_name(index: int, boundary: str) -> str:
        marker = "S" if boundary == "start" else "E"
        return f"BDSG_{marker}_{index:04d}"

    def _prepare_full_variant(
        self,
        word,
        input_path: Path,
        groups: list[ChangeGroup],
        side: str,
        out_path: Path,
        stage: str,
        base_percent: float,
        span_percent: float,
    ):
        self._check_cancelled()
        shutil.copy2(input_path, out_path)
        document = word.Documents.Open(
            str(out_path),
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
            Visible=False,
            OpenAndRepair=False,
            NoEncodingDialog=True,
        )
        try:
            document.TrackRevisions = False
            content_end = int(document.Content.End) - 1
            for group_index, group in enumerate(groups, start=1):
                safe_start = max(0, min(group.range_start, content_end))
                safe_end = max(safe_start, min(group.range_end, content_end))
                document.Bookmarks.Add(
                    self._bookmark_name(group_index, "start"),
                    document.Range(safe_start, safe_start),
                )
                document.Bookmarks.Add(
                    self._bookmark_name(group_index, "end"),
                    document.Range(safe_end, safe_end),
                )

            revision_count = int(document.Revisions.Count)
            report_step = max(1, revision_count // 50) if revision_count else 1
            for index in range(1, revision_count + 1):
                self._check_cancelled()
                revision = document.Revisions(index)
                revision_type = int(revision.Type)
                if side == "before":
                    if revision_type in OLD_SIDE_REVISIONS or revision_type == WD_REVISION_REPLACE:
                        revision.Range.Font.Color = WD_COLOR_RED
                        revision.Range.Font.StrikeThrough = True
                else:
                    if revision_type in NEW_SIDE_REVISIONS or revision_type == WD_REVISION_REPLACE:
                        revision.Range.Font.Color = WD_COLOR_BLUE
                        revision.Range.Font.Underline = WD_UNDERLINE_SINGLE
                if index == 1 or index == revision_count or index % report_step == 0:
                    # Applying formatting does not mutate the Revisions collection.
                    # Accepting/rejecting each item here would: Word may merge adjacent
                    # revisions and remove more than one collection element at once.
                    fraction = index / max(1, revision_count)
                    self._report(
                        stage,
                        base_percent + span_percent * 0.9 * fraction,
                        f"변경 표시 {index}/{revision_count}",
                        index,
                        revision_count,
                    )

            self._check_cancelled()
            if revision_count:
                action = "거부" if side == "before" else "승인"
                self._report(
                    stage,
                    base_percent + span_percent * 0.92,
                    f"변경 기록 {revision_count}개 일괄 {action} 중",
                    revision_count,
                    revision_count,
                )
                if side == "before":
                    document.Revisions.RejectAll()
                else:
                    document.Revisions.AcceptAll()
                remaining = int(document.Revisions.Count)
                if remaining:
                    raise RuntimeError(
                        f"변경 기록 일괄 {action} 후 {remaining}개가 남았습니다."
                    )
                self._report(
                    stage,
                    base_percent + span_percent,
                    f"변경 기록 {revision_count}개 일괄 {action} 완료",
                    revision_count,
                    revision_count,
                )
            missing = [
                self._bookmark_name(index, boundary)
                for index in range(1, len(groups) + 1)
                for boundary in ("start", "end")
                if not bool(document.Bookmarks.Exists(self._bookmark_name(index, boundary)))
            ]
            if missing:
                raise RuntimeError(
                    "변경 처리 중 일부 범위 북마크가 사라졌습니다: " + ", ".join(missing[:5])
                )
            document.Save()
            return document
        except Exception:
            with suppress(Exception):
                document.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            raise

    @staticmethod
    def _set_cell_text(cell, value: str) -> None:
        # Newly created comparison cells are already blank. Assigning an empty
        # string to some Word table ranges can be interpreted as deleting a row
        # end marker and raise "줄 끝에 대한 작업이 잘못되었습니다."
        if value == "":
            return
        target = cell.Range
        target.End -= 1
        target.Text = value

    @staticmethod
    def _copy_range_to_cell(source_range, target_cell) -> None:
        source = source_range.Duplicate
        if source.End > source.Start:
            source.End -= 1
        # Never replace or clear the cell range itself. A Word table cell includes
        # protected paragraph/end-of-cell markers, and assigning an empty string
        # can fail at a row boundary. Newly created cells are already empty, so
        # insert at a collapsed start point while preserving both markers.
        target = target_cell.Range.Duplicate
        target.Collapse(WD_COLLAPSE_START)
        target.FormattedText = source.FormattedText

    @staticmethod
    def _format_excerpt_cell(cell, group: ChangeGroup) -> None:
        """Apply human-readable comparison-table typography and omissions."""
        content = cell.Range.Duplicate
        content.End -= 1
        content.Font.Size = 8
        if group.context_type == "table":
            for index in range(1, int(content.Tables.Count) + 1):
                content.Tables(index).Range.Font.Size = 7

        markers: list[tuple[int, int]] = []
        if group.omit_before:
            label = "<전략>\r" if group.context_type == "paragraph" else "<앞쪽 표 행 생략>\r"
            start = int(content.Start)
            content.InsertBefore(label)
            markers.append((start, start + len(label) - 1))
        if group.omit_after:
            label = "\r<후략>" if group.context_type == "paragraph" else "\r<뒤쪽 표 행 생략>"
            end = int(cell.Range.End) - 1
            tail = cell.Range.Duplicate
            tail.SetRange(end, end)
            tail.InsertBefore(label)
            markers.append((end + 1, end + len(label)))
        for start, end in markers:
            marker = cell.Range.Duplicate
            marker.SetRange(start, end)
            marker.Font.Size = 7
            marker.Font.Color = WD_COLOR_GRAY_50
            marker.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER

    def _prepare_comparison_rows(
        self,
        comparison_table,
        groups: list[ChangeGroup],
    ) -> list[tuple[Any, Any]]:
        # Create every outer table row before inserting rich Word content. Adding
        # another row after a previous cell received a nested/complex table can
        # make Word return a row with fewer logical cells.
        for _group in groups:
            comparison_table.Rows.Add()

        expected_rows = len(groups) + 1
        actual_rows = int(comparison_table.Rows.Count)
        if actual_rows < expected_rows:
            raise RuntimeError(
                f"변경대비표 행 생성이 완료되지 않았습니다: {actual_rows}/{expected_rows}"
            )

        content_cells: list[tuple[Any, Any]] = []
        for index, group in enumerate(groups, start=1):
            row_number = index + 1
            try:
                section_cell = comparison_table.Cell(row_number, 1)
                page_cell = comparison_table.Cell(row_number, 2)
                before_cell = comparison_table.Cell(row_number, 3)
                after_cell = comparison_table.Cell(row_number, 4)
                reason_cell = comparison_table.Cell(row_number, 5)
            except Exception as exc:
                raise RuntimeError(
                    f"변경대비표 {index}번 항목의 5열 셀 구조를 만들지 못했습니다."
                ) from exc
            self._set_cell_text(section_cell, group.section)
            self._set_cell_text(page_cell, group.page_label)
            # The user requested that every reason cell remain blank.
            self._set_cell_text(reason_cell, "")
            content_cells.append((before_cell, after_cell))
        return content_cells

    def _set_landscape_layout(self, document) -> None:
        """Set every output section to landscape without touching source files."""
        section_count = int(document.Sections.Count)
        for index in range(1, section_count + 1):
            self._retry_word_call(
                f"출력 문서 Section {index} 가로 방향 설정",
                lambda index=index: setattr(
                    document.Sections(index).PageSetup,
                    "Orientation",
                    WD_ORIENT_LANDSCAPE,
                ),
                progress_info=("3/5 템플릿 준비", 80, index, section_count),
            )

    def generate(
        self,
        input_path: Path,
        output_path: Path,
        groups: list[ChangeGroup],
        metadata: dict[str, str],
    ) -> None:
        if not TEMPLATE_PATH.exists():
            raise FileNotFoundError(f"템플릿을 찾을 수 없습니다: {TEMPLATE_PATH}")
        # Cleanup errors must never hide the original Word error. Document/Word
        # shutdown below is retried first; this flag is only the final safeguard.
        with tempfile.TemporaryDirectory(
            prefix="change_table_", ignore_cleanup_errors=True
        ) as temp_dir_name:
            temp_dir = Path(temp_dir_name)
            group_count = len(groups)
            working_path = temp_dir / "source_without_comments.docx"
            self._report(
                "1/5 병렬 생성 준비",
                1,
                "메모가 제거된 생성용 복사본을 준비하는 중",
            )
            comment_count = create_comment_free_copy(input_path, working_path)
            self._report(
                "1/5 병렬 생성 준비",
                3,
                f"메모 {comment_count}개 제거 · 본문 변경 묶음 {group_count}개 · 독립 Word 프로세스 최대 2개",
            )
            jobs = [
                ("before", temp_dir / "full_before.docx"),
                ("after", temp_dir / "full_after.docx"),
            ]
            prepared: dict[str, Path] = {}
            prepared_lock = threading.Lock()
            worker_count = 2

            def variant_worker(worker_number: int, side: str, variant_path: Path) -> None:
                child = WordEngine(
                    progress=self.progress,
                    cancel_event=self.cancel_event,
                    logger=self.logger,
                )
                child_word = None
                try:
                    child_word = child._start_word(
                        f"2/5 병렬 Word {worker_number} 시작", 4
                    )
                    self._check_cancelled()
                    side_label = "변경 전" if side == "before" else "변경 후"
                    variant = child._prepare_full_variant(
                        child_word,
                        working_path,
                        groups,
                        side,
                        variant_path,
                        f"2/5 전체 {side_label} 문서 생성 · Word {worker_number}",
                        6,
                        70,
                    )
                    child._close_document(variant, f"전체 {side_label} 문서")
                    variant = None
                    with prepared_lock:
                        prepared[side] = variant_path
                finally:
                    child._close_word(None, child_word)

            with ThreadPoolExecutor(
                max_workers=worker_count, thread_name_prefix="word-variant"
            ) as executor:
                futures = [
                    executor.submit(variant_worker, index + 1, side, path)
                    for index, (side, path) in enumerate(jobs)
                ]
                for future in futures:
                    future.result()

            self._check_cancelled()
            if set(prepared) != {"before", "after"}:
                raise RuntimeError("변경 전·후 전체 문서가 모두 생성되지 않았습니다.")

            word = None
            output_document = None
            before_document = None
            after_document = None
            try:
                word = self._start_word("3/5 최종 Word 시작", 78)
                # The final document can contain many pages inside a table cell.
                # Background pagination and proofing can keep Word's COM server
                # busy long after FormattedText insertion has returned.
                with suppress(Exception):
                    word.Options.Pagination = False
                with suppress(Exception):
                    word.Options.CheckSpellingAsYouType = False
                with suppress(Exception):
                    word.Options.CheckGrammarAsYouType = False
                self._report("3/5 템플릿 준비", 80, "최종 변경대비표 구조를 준비하는 중")
                output_path.parent.mkdir(parents=True, exist_ok=True)
                output_path.write_bytes(TEMPLATE_PATH.read_bytes())
                output_document = self._retry_word_call(
                    "출력 템플릿 열기",
                    lambda: word.Documents.Open(
                        str(output_path),
                        ConfirmConversions=False,
                        ReadOnly=False,
                        AddToRecentFiles=False,
                        Visible=False,
                        OpenAndRepair=False,
                        NoEncodingDialog=True,
                    ),
                )

                if output_document.Tables.Count < 2:
                    raise ValueError("변경대비표 템플릿 구조가 올바르지 않습니다.")
                output_document.TrackRevisions = False
                self._set_landscape_layout(output_document)
                metadata_table = output_document.Tables(1)
                comparison_table = output_document.Tables(2)

                self._set_cell_text(metadata_table.Cell(1, 2), metadata.get("document_name", ""))

                try:
                    header = output_document.Sections(1).Headers(1).Range
                    header.Text = f"변경대비표 ({metadata.get('document_name', '')})"
                    header.Font.Italic = True
                    header.Font.Bold = True
                except Exception:
                    pass

                comparison_table.Rows(1).HeadingFormat = True
                with suppress(Exception):
                    comparison_table.Rows.AllowBreakAcrossPages = True
                before_document = self._retry_word_call(
                    "변경 전 문서 열기",
                    lambda: word.Documents.Open(
                        str(prepared["before"]),
                        ConfirmConversions=False,
                        ReadOnly=True,
                        AddToRecentFiles=False,
                        Visible=False,
                        OpenAndRepair=False,
                        NoEncodingDialog=True,
                    ),
                )
                after_document = self._retry_word_call(
                    "변경 후 문서 열기",
                    lambda: word.Documents.Open(
                        str(prepared["after"]),
                        ConfirmConversions=False,
                        ReadOnly=True,
                        AddToRecentFiles=False,
                        Visible=False,
                        OpenAndRepair=False,
                        NoEncodingDialog=True,
                    ),
                )
                self._report(
                    "4/5 최종 표 조립",
                    81,
                    f"5열 표 행 {group_count}개를 먼저 준비하는 중",
                    0,
                    group_count,
                )
                content_cells = self._prepare_comparison_rows(comparison_table, groups)
                for index, group in enumerate(groups, start=1):
                    self._check_cancelled()
                    before_cell, after_cell = content_cells[index - 1]

                    start_bookmark = self._bookmark_name(index, "start")
                    end_bookmark = self._bookmark_name(index, "end")
                    for label, document in (
                        ("변경 전", before_document),
                        ("변경 후", after_document),
                    ):
                        progress_info = (
                            "4/5 최종 표 조립",
                            81 + (15 * (index - 1) / max(1, group_count)),
                            index,
                            group_count,
                        )
                        start_exists = self._retry_word_call(
                            f"{label} 시작 범위 확인",
                            lambda document=document: bool(
                                document.Bookmarks.Exists(start_bookmark)
                            ),
                            progress_info=progress_info,
                        )
                        if not start_exists:
                            raise RuntimeError(f"{label} 문서에서 시작 범위를 찾을 수 없습니다: {start_bookmark}")
                        end_exists = self._retry_word_call(
                            f"{label} 끝 범위 확인",
                            lambda document=document: bool(
                                document.Bookmarks.Exists(end_bookmark)
                            ),
                            progress_info=progress_info,
                        )
                        if not end_exists:
                            raise RuntimeError(f"{label} 문서에서 끝 범위를 찾을 수 없습니다: {end_bookmark}")
                    before_start = int(self._retry_word_call(
                        "변경 전 시작 위치 조회",
                        lambda: before_document.Bookmarks(start_bookmark).Range.Start,
                        progress_info=progress_info,
                    ))
                    before_end = int(self._retry_word_call(
                        "변경 전 끝 위치 조회",
                        lambda: before_document.Bookmarks(end_bookmark).Range.Start,
                        progress_info=progress_info,
                    ))
                    after_start = int(self._retry_word_call(
                        "변경 후 시작 위치 조회",
                        lambda: after_document.Bookmarks(start_bookmark).Range.Start,
                        progress_info=progress_info,
                    ))
                    after_end = int(self._retry_word_call(
                        "변경 후 끝 위치 조회",
                        lambda: after_document.Bookmarks(end_bookmark).Range.Start,
                        progress_info=progress_info,
                    ))
                    before_range = self._retry_word_call(
                        "변경 전 범위 읽기",
                        lambda: before_document.Range(
                            min(before_start, before_end), max(before_start, before_end)
                        ),
                        progress_info=progress_info,
                    )
                    after_range = self._retry_word_call(
                        "변경 후 범위 읽기",
                        lambda: after_document.Range(
                            min(after_start, after_end), max(after_start, after_end)
                        ),
                        progress_info=progress_info,
                    )
                    self._retry_word_call(
                        "변경 전 내용 삽입",
                        lambda: self._copy_range_to_cell(before_range, before_cell),
                        progress_info=progress_info,
                    )
                    self._retry_word_call(
                        "변경 후 내용 삽입",
                        lambda: self._copy_range_to_cell(after_range, after_cell),
                        progress_info=progress_info,
                    )
                    self._retry_word_call(
                        "변경 전 발췌 서식 적용",
                        lambda: self._format_excerpt_cell(before_cell, group),
                        progress_info=progress_info,
                    )
                    self._retry_word_call(
                        "변경 후 발췌 서식 적용",
                        lambda: self._format_excerpt_cell(after_cell, group),
                        progress_info=progress_info,
                    )
                    self._report(
                        "4/5 최종 표 조립",
                        81 + (15 * index / max(1, group_count)),
                        f"page {group.page_label} · {group.section}",
                        index,
                        group_count,
                    )

                # Release source/cell proxies and give Word a short message-pump
                # window before saving the large table. Updating all fields is
                # intentionally omitted: the template does not require it and it
                # can start an expensive, persistent layout/update operation.
                before_range = None
                after_range = None
                before_cell = None
                after_cell = None
                content_cells.clear()
                content_cells = []
                comparison_table = None
                metadata_table = None
                gc.collect()
                self._report(
                    "5/5 문서 저장",
                    97,
                    "표 레이아웃을 정리한 후 저장하는 중",
                )
                self._wait_for_word_retry(1.0, True)
                self._retry_word_call(
                    "변경대비표 저장",
                    lambda: output_document.Save(),
                    attempts=WORD_SAVE_RETRY_ATTEMPTS,
                    progress_info=("5/5 문서 저장", 98, None, None),
                )
                self._report("5/5 생성 마무리", 99, f"{output_path.name} · Word 종료 중")
            finally:
                self._close_document(before_document, "변경 전 임시 문서")
                self._close_document(after_document, "변경 후 임시 문서")
                before_document = None
                after_document = None
                self._close_word(output_document, word)


class AlphaApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        root.title(APP_TITLE)
        root.geometry("1120x780")
        root.minsize(900, 650)
        root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.input_var = StringVar()
        self.document_name_var = StringVar()
        self.status_var = StringVar(value="tracked DOCX를 선택하십시오.")
        self.stage_var = StringVar(value="대기 중")
        self.detail_var = StringVar(value="진행할 작업이 없습니다.")
        self.elapsed_var = StringVar(value="경과 시간 00:00")
        self.progress_var = DoubleVar(value=0)
        self.groups: list[ChangeGroup] = []
        self.task_queue: queue.Queue[dict[str, Any]] = queue.Queue()
        self.cancel_event: threading.Event | None = None
        self.busy = False
        self.task_kind: str | None = None
        self.task_started_at = 0.0
        self.last_progress_at = 0.0
        self.cancel_requested_at: float | None = None
        self.stall_warned = False
        self.word_pids: set[int] = set()
        self.last_log_path: Path | None = None
        self.pending_output_path: Path | None = None

        frame = ttk.Frame(root, padding=14)
        frame.pack(fill="both", expand=True)

        file_row = ttk.Frame(frame)
        file_row.pack(fill="x")
        ttk.Label(file_row, text="입력 파일").pack(side="left")
        ttk.Entry(file_row, textvariable=self.input_var).pack(side="left", fill="x", expand=True, padx=8)
        self.choose_button = ttk.Button(file_row, text="tracked DOCX 선택", command=self.choose_file)
        self.choose_button.pack(side="left")

        meta = ttk.LabelFrame(frame, text="문서 정보", padding=10)
        meta.pack(fill="x", pady=(12, 8))
        ttk.Label(meta, text="문서명").grid(row=0, column=0, sticky="w", padx=(0, 5))
        ttk.Entry(meta, textvariable=self.document_name_var, width=80).grid(
            row=0, column=1, sticky="ew", padx=(0, 12)
        )
        meta.columnconfigure(1, weight=2)

        action_row = ttk.Frame(frame)
        action_row.pack(fill="x", pady=(0, 8))
        self.analyze_button = ttk.Button(action_row, text="변경 분석", command=self.analyze)
        self.analyze_button.pack(side="left")
        self.merge_button = ttk.Button(action_row, text="선택 항목 병합", command=self.merge_selected)
        self.merge_button.pack(side="left", padx=6)
        self.remove_button = ttk.Button(action_row, text="선택 항목 제외", command=self.remove_selected)
        self.remove_button.pack(side="left")
        self.generate_button = ttk.Button(action_row, text="변경대비표 생성", command=self.generate)
        self.generate_button.pack(side="right")

        progress_frame = ttk.LabelFrame(frame, text="진행 상황", padding=10)
        progress_frame.pack(fill="x", pady=(0, 8))
        progress_top = ttk.Frame(progress_frame)
        progress_top.pack(fill="x")
        ttk.Label(progress_top, textvariable=self.stage_var).pack(side="left")
        ttk.Label(progress_top, textvariable=self.elapsed_var).pack(side="right")
        ttk.Progressbar(
            progress_frame, variable=self.progress_var, maximum=100, mode="determinate"
        ).pack(fill="x", pady=6)
        progress_bottom = ttk.Frame(progress_frame)
        progress_bottom.pack(fill="x")
        ttk.Label(progress_bottom, textvariable=self.detail_var).pack(side="left", fill="x", expand=True)
        self.log_button = ttk.Button(progress_bottom, text="로그 열기", command=self.open_log, state="disabled")
        self.log_button.pack(side="right", padx=(6, 0))
        self.cancel_button = ttk.Button(
            progress_bottom, text="작업 취소", command=self.cancel_task, state="disabled"
        )
        self.cancel_button.pack(side="right")

        columns = ("index", "section", "page", "summary")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("index", text="번호")
        self.tree.heading("section", text="Section")
        self.tree.heading("page", text="page")
        self.tree.heading("summary", text="변경 내용 미리보기")
        self.tree.column("index", width=55, anchor="center", stretch=False)
        self.tree.column("section", width=290)
        self.tree.column("page", width=80, anchor="center", stretch=False)
        self.tree.column("summary", width=590)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self.edit_selected)

        ttk.Label(frame, textvariable=self.status_var).pack(fill="x", pady=(8, 0))

    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        state = "disabled" if busy else "normal"
        for button in (
            self.choose_button,
            self.analyze_button,
            self.merge_button,
            self.remove_button,
            self.generate_button,
        ):
            button.configure(state=state)
        self.cancel_button.configure(state="normal" if busy else "disabled", text="작업 취소")
        self.log_button.configure(state="normal" if self.last_log_path else "disabled")
        self.root.config(cursor="watch" if busy else "")

    def _begin_task(
        self,
        kind: str,
        input_path: Path,
        output_path: Path | None = None,
    ) -> None:
        while True:
            try:
                self.task_queue.get_nowait()
            except queue.Empty:
                break
        logger, log_path = create_operation_logger("분석" if kind == "analyze" else "생성")
        self.last_log_path = log_path
        self.log_button.configure(state="normal")
        self.cancel_event = threading.Event()
        self.task_kind = kind
        self.task_started_at = time.monotonic()
        self.last_progress_at = self.task_started_at
        self.cancel_requested_at = None
        self.stall_warned = False
        self.word_pids.clear()
        self.pending_output_path = output_path
        self.progress_var.set(0)
        self.stage_var.set("작업 준비")
        self.detail_var.set("Microsoft Word 작업을 준비하고 있습니다.")
        self.elapsed_var.set("경과 시간 00:00")
        self.status_var.set("백그라운드 작업을 시작했습니다.")
        self._set_busy(True)

        metadata = self._metadata()
        groups = list(self.groups)

        def worker() -> None:
            try:
                logger.info("작업 시작: %s", kind)
                logger.info("입력 파일: %s", input_path)
                engine = WordEngine(
                    progress=self.task_queue.put,
                    cancel_event=self.cancel_event,
                    logger=logger,
                )
                if kind == "analyze":
                    result_groups, detected = engine.analyze(input_path)
                    self.task_queue.put({
                        "kind": "done",
                        "operation": kind,
                        "groups": result_groups,
                        "metadata": detected,
                    })
                else:
                    if output_path is None:
                        raise ValueError("출력 파일 경로가 없습니다.")
                    engine.generate(input_path, output_path, groups, metadata)
                    self.task_queue.put({
                        "kind": "done",
                        "operation": kind,
                        "output_path": output_path,
                    })
                logger.info("작업 완료")
            except OperationCancelled as exc:
                logger.warning("작업 취소: %s", exc)
                self.task_queue.put({"kind": "cancelled", "message": str(exc)})
            except Exception as exc:
                logger.exception("작업 오류")
                self.task_queue.put({
                    "kind": "error",
                    "message": str(exc),
                    "traceback": traceback.format_exc(),
                })
            finally:
                close_operation_logger(logger)

        threading.Thread(target=worker, name=f"byeondaesaenggi-{kind}", daemon=True).start()
        self.root.after(100, self._poll_task)

    def _poll_task(self) -> None:
        if not self.busy:
            return
        terminal_event: dict[str, Any] | None = None
        while True:
            try:
                event = self.task_queue.get_nowait()
            except queue.Empty:
                break
            if event.get("kind") == "progress":
                self.last_progress_at = time.monotonic()
                self.progress_var.set(max(
                    float(self.progress_var.get()), float(event.get("percent", 0))
                ))
                self.stage_var.set(str(event.get("stage", "처리 중")))
                detail = str(event.get("detail", ""))
                current, total = event.get("current"), event.get("total")
                if current is not None and total:
                    detail = f"변경 기록 {current}/{total}" + (f" · {detail}" if detail else "")
                self.detail_var.set(detail)
                if event.get("word_pid"):
                    self.word_pids.add(int(event["word_pid"]))
            elif event.get("kind") == "word_closed":
                if event.get("word_pid"):
                    self.word_pids.discard(int(event["word_pid"]))
            else:
                terminal_event = event

        now = time.monotonic()
        self.elapsed_var.set(f"경과 시간 {format_elapsed(now - self.task_started_at)}")
        silent_for = now - self.last_progress_at
        if silent_for >= STALL_WARNING_SECONDS and not self.stall_warned:
            self.stall_warned = True
            self.status_var.set(
                "2분 이상 진행 정보가 없습니다. Word가 응답을 기다리는 중일 수 있습니다."
            )
            self.detail_var.set(
                f"정지 가능성 감지 · 마지막 진행 후 {format_elapsed(silent_for)} · 필요하면 작업 취소"
            )
        if self.cancel_requested_at is not None:
            if now - self.cancel_requested_at >= FORCE_STOP_SECONDS and self.word_pids:
                self.cancel_button.configure(text="Word 강제 종료")
            else:
                self.cancel_button.configure(text="취소 요청됨")

        if terminal_event is not None:
            self._finish_task(terminal_event)
            return
        self.root.after(150, self._poll_task)

    def _show_completion(self, message: str) -> None:
        # Attach the dialog to the main window and briefly lift it so completion
        # is visible even when a long Word task finished behind another window.
        with suppress(Exception):
            self.root.bell()
            self.root.lift()
            self.root.attributes("-topmost", True)
        try:
            messagebox.showinfo(APP_TITLE, message, parent=self.root)
        finally:
            with suppress(Exception):
                self.root.attributes("-topmost", False)

    def _finish_task(self, event: dict[str, Any]) -> None:
        operation = self.task_kind
        self._set_busy(False)
        self.word_pids.clear()
        self.cancel_requested_at = None
        kind = event.get("kind")
        if kind == "done" and operation == "analyze":
            self.groups = event["groups"]
            self._refresh_tree()
            self.progress_var.set(100)
            self.stage_var.set("분석 완료")
            self.detail_var.set(f"변경 묶음 {len(self.groups)}개")
            self.status_var.set(
                f"{len(self.groups)}개 변경 묶음을 탐지했습니다. Section과 page를 확인하십시오."
            )
            self._show_completion(
                "변경 분석이 완료되었습니다.\n\n"
                f"변경 묶음: {len(self.groups)}개\n\n"
                "Section과 page를 확인한 후 '변경대비표 생성'을 진행하십시오."
            )
        elif kind == "done" and operation == "generate":
            output_path = Path(event["output_path"])
            self.progress_var.set(100)
            self.stage_var.set("생성 완료")
            self.detail_var.set(output_path.name)
            self.status_var.set(f"생성 완료: {output_path}")
            self._show_completion(f"변경대비표 생성이 완료되었습니다.\n\n{output_path}")
            with suppress(OSError):
                os.startfile(output_path)
        elif kind == "cancelled" or (self.cancel_event and self.cancel_event.is_set()):
            self.stage_var.set("작업 취소")
            self.detail_var.set("작업이 취소되었습니다. 로그에서 마지막 처리 단계를 확인할 수 있습니다.")
            self.status_var.set("작업을 취소했습니다.")
        else:
            message = str(event.get("message") or "알 수 없는 오류가 발생했습니다.")
            self.stage_var.set("작업 오류")
            self.detail_var.set("로그에서 오류 단계와 상세 내용을 확인하십시오.")
            self.status_var.set("작업 중 오류가 발생했습니다.")
            messagebox.showerror(
                APP_TITLE,
                f"{message}\n\n로그 파일:\n{self.last_log_path}",
            )
        self.task_kind = None
        self.pending_output_path = None

    def cancel_task(self) -> None:
        if not self.busy or self.cancel_event is None:
            return
        if not self.cancel_event.is_set():
            self.cancel_event.set()
            self.cancel_requested_at = time.monotonic()
            self.status_var.set("취소를 요청했습니다. 현재 Word 작업이 끝나는 즉시 중단합니다.")
            self.cancel_button.configure(text="취소 요청됨")
            return
        if (
            self.cancel_requested_at is not None
            and time.monotonic() - self.cancel_requested_at >= FORCE_STOP_SECONDS
            and self.word_pids
        ):
            if messagebox.askyesno(
                APP_TITLE,
                "변대생기가 시작한 Word 프로세스를 강제로 종료하시겠습니까?\n\n"
                "생성기 작업 중인 임시 문서는 저장되지 않습니다.",
            ):
                self._force_stop_word()
        else:
            self.status_var.set("Word 응답을 기다리고 있습니다. 잠시 후 강제 종료가 활성화됩니다.")

    def _force_stop_word(self) -> None:
        if not self.word_pids:
            return
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        results = []
        for process_id in sorted(self.word_pids):
            results.append(subprocess.run(
                ["taskkill", "/PID", str(process_id), "/T", "/F"],
                capture_output=True,
                text=True,
                creationflags=flags,
                check=False,
            ))
        if any(result.returncode == 0 for result in results):
            self.status_var.set("변대생기 전용 Word 프로세스를 종료했습니다.")
        else:
            self.status_var.set("Word 프로세스를 종료하지 못했습니다. 작업 관리자를 확인하십시오.")

    def open_log(self) -> None:
        if self.last_log_path and self.last_log_path.exists():
            with suppress(OSError):
                os.startfile(self.last_log_path)

    def on_close(self) -> None:
        if self.busy:
            if not messagebox.askyesno(
                APP_TITLE,
                "작업이 진행 중입니다. 변대생기 전용 Word 작업을 종료하고 창을 닫으시겠습니까?",
            ):
                return
            if self.cancel_event:
                self.cancel_event.set()
            self._force_stop_word()
        self.root.destroy()

    def choose_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="변경 추적 DOCX 선택",
            filetypes=[("Word 문서", "*.docx")],
        )
        if not filename:
            return
        path = Path(filename)
        self.input_var.set(str(path))
        metadata = detect_metadata_from_name(path)
        self.document_name_var.set(metadata["document_name"])
        self.status_var.set("파일을 선택했습니다. 변경 분석을 실행하십시오.")

    def _metadata(self) -> dict[str, str]:
        return {
            "document_name": self.document_name_var.get().strip(),
        }

    def _refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        for index, group in enumerate(self.groups, start=1):
            self.tree.insert("", "end", iid=str(index - 1), values=(index, group.section, group.page_label, group.summary))

    def analyze(self) -> None:
        input_path = Path(self.input_var.get())
        if not input_path.exists():
            messagebox.showwarning(APP_TITLE, "tracked DOCX 파일을 먼저 선택하십시오.")
            return
        self._begin_task("analyze", input_path)

    def edit_selected(self, _event=None) -> None:
        selected = self.tree.selection()
        if len(selected) != 1:
            return
        index = int(selected[0])
        group = self.groups[index]
        section = simpledialog.askstring("Section 수정", "Section", initialvalue=group.section, parent=self.root)
        if section is None:
            return
        page = simpledialog.askstring("page 수정", "page 또는 page 범위", initialvalue=group.page_label, parent=self.root)
        if page is None:
            return
        match = re.fullmatch(r"\s*(\d+)\s*(?:-\s*(\d+)\s*)?", page)
        if not match:
            messagebox.showwarning(APP_TITLE, "page는 50 또는 50-51 형식으로 입력하십시오.")
            return
        group.section = section.strip()
        group.page_start = int(match.group(1))
        group.page_end = int(match.group(2) or match.group(1))
        self._refresh_tree()

    def merge_selected(self) -> None:
        indices = sorted(int(item) for item in self.tree.selection())
        if len(indices) < 2 or indices != list(range(indices[0], indices[-1] + 1)):
            messagebox.showwarning(APP_TITLE, "서로 인접한 항목을 두 개 이상 선택하십시오.")
            return
        selected = [self.groups[index] for index in indices]
        merged = ChangeGroup(
            section=selected[0].section,
            page_start=min(group.page_start for group in selected),
            page_end=max(group.page_end for group in selected),
            range_start=min(group.range_start for group in selected),
            range_end=max(group.range_end for group in selected),
            summary=clean_word_text(" / ".join(group.summary for group in selected))[:180],
        )
        self.groups[indices[0] : indices[-1] + 1] = [merged]
        self._refresh_tree()

    def remove_selected(self) -> None:
        indices = sorted((int(item) for item in self.tree.selection()), reverse=True)
        for index in indices:
            del self.groups[index]
        self._refresh_tree()

    def generate(self) -> None:
        input_path = Path(self.input_var.get())
        if not input_path.exists() or not self.groups:
            messagebox.showwarning(APP_TITLE, "먼저 tracked DOCX를 선택하고 변경 분석을 실행하십시오.")
            return
        default_name = input_path.stem.replace("draft-tracked", "변경대비표_alpha") + ".docx"
        output = filedialog.asksaveasfilename(
            title="변경대비표 저장",
            defaultextension=".docx",
            initialfile=default_name,
            filetypes=[("Word 문서", "*.docx")],
        )
        if not output:
            return
        output_path = Path(output)
        if output_path.resolve() == input_path.resolve():
            messagebox.showwarning(APP_TITLE, "입력 파일과 다른 이름으로 저장하십시오.")
            return
        self._begin_task("generate", input_path, output_path)


def self_test() -> int:
    sample = detect_metadata_from_name(Path("시험계획서_V3.2_draft-tracked.docx"))
    assert sample["document_name"] == "시험계획서 V3.2"
    assert ChangeGroup("A", 3, 5, 1, 2, "x").page_label == "3-5"
    xml_groups = group_xml_revisions([
        XmlRevisionRecord(1, 10, 1, "A", "ins", "x"),
        XmlRevisionRecord(2, 11, 1, "A", "del", "y"),
        XmlRevisionRecord(3, 20, 2, "B", "ins", "z"),
    ])
    assert len(xml_groups) == 2
    assert xml_groups[0].revision_start == 1 and xml_groups[0].revision_end == 2
    assert format_elapsed(0) == "00:00"
    assert format_elapsed(125.9) == "02:05"
    assert STALL_WARNING_SECONDS == 120
    assert WordEngine._bookmark_name(3, "start") == "BDSG_S_0003"
    assert WordEngine._bookmark_name(3, "end") == "BDSG_E_0003"
    print("PASS: 변분기 Alpha 0.2 engine self-test")
    return 0


def main() -> int:
    if "--self-test" in sys.argv:
        return self_test()
    print("변분기는 RUN.bat 또는 analyzer.py로 실행하십시오.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
