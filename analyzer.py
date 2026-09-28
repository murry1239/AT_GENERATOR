from __future__ import annotations

import difflib
import json
import logging
import os
import queue
import re
import shutil
import sys
import tempfile
import threading
import time
import traceback
import zipfile
from concurrent.futures import ThreadPoolExecutor
from contextlib import suppress
from dataclasses import asdict
from pathlib import Path
from tkinter import DoubleVar, StringVar, Tk, filedialog, messagebox
from tkinter import ttk

from app import (
    APP_DIR,
    ChangeGroup,
    OperationCancelled,
    WordEngine,
    close_operation_logger,
    create_comment_free_copy,
    create_operation_logger,
    detect_metadata_from_name,
)


VERSION = "변분기 Alpha 0.2"
PACKAGE_EXTENSION = ".bdsg"
PACKAGE_FORMAT = "byeondaesaenggi-analysis-v1"


def _safe_extract(archive: zipfile.ZipFile, destination: Path) -> None:
    root = destination.resolve()
    for member in archive.infolist():
        target = (destination / member.filename).resolve()
        if root != target and root not in target.parents:
            raise ValueError("분석 패키지 내부 경로가 올바르지 않습니다.")
    archive.extractall(destination)


def _range_text(document, group_index: int, engine: WordEngine) -> str:
    start_name = engine._bookmark_name(group_index, "start")
    end_name = engine._bookmark_name(group_index, "end")
    if not document.Bookmarks.Exists(start_name) or not document.Bookmarks.Exists(end_name):
        raise RuntimeError(f"분석 범위 북마크가 없습니다: {group_index}")
    start = int(document.Bookmarks(start_name).Range.Start)
    end = int(document.Bookmarks(end_name).Range.Start)
    value = str(document.Range(min(start, end), max(start, end)).Text or "")
    # Word table cells end with CR+BEL. Preserve rows and cells in plain text so
    # the preview remains usable without nesting a table in another Word table.
    value = value.replace("\r\x07", "\n").replace("\x07", "\t")
    value = value.replace("\r", "\n").replace("\v", "\n")
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in value.splitlines()]
    return "\n".join(line for line in lines if line).strip()


def _with_omission_markers(text: str, group: ChangeGroup) -> str:
    before = "<앞쪽 표 행 생략>" if group.context_type == "table" else "<전략>"
    after = "<뒤쪽 표 행 생략>" if group.context_type == "table" else "<후략>"
    parts = []
    if group.omit_before:
        parts.append(before)
    if text:
        parts.append(text)
    if group.omit_after:
        parts.append(after)
    return "\n".join(parts)


def _bookmark_range(document, group_index: int, engine: WordEngine):
    start_name = engine._bookmark_name(group_index, "start")
    end_name = engine._bookmark_name(group_index, "end")
    if not document.Bookmarks.Exists(start_name) or not document.Bookmarks.Exists(end_name):
        raise RuntimeError(f"분석 범위 북마크가 없습니다: {group_index}")
    start = int(document.Bookmarks(start_name).Range.Start)
    end = int(document.Bookmarks(end_name).Range.Start)
    return document.Range(min(start, end), max(start, end))


def _capture_range_png(document, group_index: int, engine: WordEngine, output: Path) -> None:
    """Capture a Word range as a PNG without inserting it into another table."""
    from PIL import Image, ImageGrab

    source_range = _bookmark_range(document, group_index, engine)
    last_error: Exception | None = None
    for attempt in range(1, 7):
        try:
            engine._retry_word_call(
                f"표 이미지 클립보드 복사 {group_index}",
                lambda: source_range.CopyAsPicture(),
                attempts=4,
            )
            engine._wait_for_word_retry(0.25 + attempt * 0.10, True)
            grabbed = ImageGrab.grabclipboard()
            if isinstance(grabbed, Image.Image):
                image = grabbed.convert("RGB")
                # Limit pathological clipboard sizes while retaining legibility.
                if image.width > 1800:
                    height = max(1, round(image.height * 1800 / image.width))
                    image = image.resize((1800, height), Image.Resampling.LANCZOS)
                output.parent.mkdir(parents=True, exist_ok=True)
                image.save(output, "PNG", optimize=True)
                return
            last_error = RuntimeError("클립보드에서 그림을 읽지 못했습니다.")
        except Exception as exc:
            last_error = exc
        engine._wait_for_word_retry(min(1.5, attempt * 0.25), True)
    raise RuntimeError(f"Word 범위를 그림으로 캡처하지 못했습니다: {last_error}")


def _write_preview_xlsx(path: Path, items: list[dict], document_name: str) -> None:
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as ExcelImage
    from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    # Keep one row per reviewable change range. The page value is repeated when
    # a page has several ranges, which avoids stacking multiple images in one cell.
    page_items = list(items)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "페이지별 변경 미리보기"
    sheet.sheet_view.showGridLines = False
    sheet.merge_cells("A1:F1")
    sheet["A1"] = f"변경 내용 미리보기 · {document_name}"
    sheet["A1"].font = Font(name="맑은 고딕", size=14, bold=True)
    sheet["A1"].alignment = Alignment(horizontal="center")
    headers = ["No.", "Section", "page", "변경 유형", "변경 전", "변경 후"]
    for column, value in enumerate(headers, start=1):
        cell = sheet.cell(3, column, value)
        cell.font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row_number, item in enumerate(page_items, start=4):
        before_image = item.get("before_image")
        after_image = item.get("after_image")
        values = [
            item["index"], item["section"], item["page"],
            "표" if item.get("context_type") == "table" else "문단",
            "" if before_image else item["before"],
            "" if after_image else item["after"],
        ]
        for column, value in enumerate(values, start=1):
            cell = sheet.cell(row_number, column, value)
            cell.font = Font(name="맑은 고딕", size=9)
            if column == 5:
                cell.fill = PatternFill("solid", fgColor="FCE8E6")
            elif column == 6:
                cell.fill = PatternFill("solid", fgColor="E8F1FB")
            cell.alignment = Alignment(
                horizontal="center" if column <= 4 else "left",
                vertical="top",
                wrap_text=True,
            )
        image_heights = []
        for column_index, column_letter, image_path in (
            (4, "E", before_image), (5, "F", after_image)
        ):
            if not image_path or not Path(image_path).exists():
                continue
            picture = ExcelImage(str(image_path))
            scale = min(1.0, 520 / max(1, picture.width), 640 / max(1, picture.height))
            picture.width = max(1, round(picture.width * scale))
            picture.height = max(1, round(picture.height * scale))
            picture.anchor = TwoCellAnchor(
                editAs="twoCell",
                _from=AnchorMarker(col=column_index, row=row_number - 1),
                to=AnchorMarker(col=column_index + 1, row=row_number),
            )
            sheet.add_image(picture)
            image_heights.append(picture.height)
        text_lines = max(item["before"].count("\n") + 1, item["after"].count("\n") + 1)
        text_height = min(180, max(36, text_lines * 14))
        picture_height = max(image_heights, default=0) * 0.75 + (12 if image_heights else 0)
        sheet.row_dimensions[row_number].height = min(405, max(text_height, picture_height))
        # Searchable source text and summary are retained in hidden columns.
        sheet.cell(row_number, 7, item["before"])
        sheet.cell(row_number, 8, item["after"])
        sheet.cell(row_number, 9, item.get("summary", ""))
    widths = [7, 30, 11, 11, 72, 72, 2, 2, 2]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet["G3"] = "변경 전 원문"
    sheet["H3"] = "변경 후 원문"
    sheet["I3"] = "변경 요약"
    for column_letter in ("G", "H", "I"):
        sheet.column_dimensions[column_letter].hidden = True
    sheet.freeze_panes = "E4"
    sheet.auto_filter.ref = f"A3:F{max(3, len(page_items) + 3)}"
    sheet.print_options.horizontalCentered = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    detail = workbook.create_sheet("변경항목별 상세")
    detail.append(["번호", "Section", "page", "형식", "미리보기 방식", "변경 전", "변경 후", "비고"])
    for item in items:
        detail.append([
            item["index"], item["section"], item["page"],
            "표" if item.get("context_type") == "table" else "문단",
            item.get("image_status", "텍스트"),
            item["before"], item["after"], item.get("image_error", ""),
        ])
    for cell in detail[1]:
        cell.font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in detail.iter_rows(min_row=2):
        for column, cell in enumerate(row, start=1):
            cell.font = Font(name="맑은 고딕", size=9)
            if column == 6:
                cell.fill = PatternFill("solid", fgColor="FCE8E6")
            elif column == 7:
                cell.fill = PatternFill("solid", fgColor="E8F1FB")
            cell.alignment = Alignment(
                horizontal="center" if column <= 5 else "left",
                vertical="top", wrap_text=True,
            )
    for index, width in enumerate([7, 32, 11, 10, 14, 60, 60, 36], start=1):
        detail.column_dimensions[get_column_letter(index)].width = width
    detail.freeze_panes = "F2"
    detail.auto_filter.ref = f"A1:H{max(1, len(items) + 1)}"
    detail.sheet_view.showGridLines = False
    workbook.save(path)


def create_analysis_package(
    source_path: Path,
    package_path: Path,
    progress,
    cancel_event: threading.Event,
    logger: logging.Logger,
) -> int:
    def analysis_progress(event: dict) -> None:
        mapped = dict(event)
        if mapped.get("kind") == "progress":
            mapped["percent"] = float(mapped.get("percent", 0)) * 0.40
            mapped["stage"] = "변경 분석 · " + str(mapped.get("stage", ""))
        progress(mapped)

    engine = WordEngine(progress=analysis_progress, cancel_event=cancel_event, logger=logger)
    groups, metadata = engine.analyze(source_path)
    document_name = metadata.get("document_name") or detect_metadata_from_name(source_path)["document_name"]

    with tempfile.TemporaryDirectory(prefix="bdsg_package_", ignore_cleanup_errors=True) as temp_name:
        temp = Path(temp_name)
        working = temp / "source_without_comments.docx"
        before_path = temp / "변경전_전체.docx"
        after_path = temp / "변경후_전체.docx"
        preview_path = temp / "변경내용_미리보기.xlsx"
        manifest_path = temp / "analysis.json"
        images_dir = temp / "images"
        create_comment_free_copy(source_path, working)
        prepared: dict[str, Path] = {}
        lock = threading.Lock()

        def prepare(side: str, output: Path, number: int) -> None:
            child = WordEngine(progress=progress, cancel_event=cancel_event, logger=logger)
            word = None
            document = None
            try:
                word = child._start_word(f"복원본 {number}/2 준비", 45 + number * 3)
                document = child._prepare_full_variant(
                    word,
                    working,
                    groups,
                    side,
                    output,
                    "변경 전·후 복원본 생성",
                    52,
                    24,
                )
                child._close_document(document, output.name)
                document = None
                with lock:
                    prepared[side] = output
            finally:
                child._close_word(document, word)

        with ThreadPoolExecutor(max_workers=2, thread_name_prefix="analysis-variant") as executor:
            futures = [
                executor.submit(prepare, "before", before_path, 1),
                executor.submit(prepare, "after", after_path, 2),
            ]
            for future in futures:
                future.result()
        if set(prepared) != {"before", "after"}:
            raise RuntimeError("변경 전·후 복원본을 모두 생성하지 못했습니다.")

        reader = WordEngine(progress=progress, cancel_event=cancel_event, logger=logger)
        word = before_doc = after_doc = None
        items: list[dict] = []
        captured_images: list[Path] = []
        try:
            word = reader._start_word("미리보기 내용 추출", 78)
            before_doc = word.Documents.Open(str(before_path), ReadOnly=True, Visible=False, AddToRecentFiles=False)
            after_doc = word.Documents.Open(str(after_path), ReadOnly=True, Visible=False, AddToRecentFiles=False)
            for index, group in enumerate(groups, start=1):
                if cancel_event.is_set():
                    raise OperationCancelled("사용자가 작업을 취소했습니다.")
                before = _with_omission_markers(_range_text(before_doc, index, reader), group)
                after = _with_omission_markers(_range_text(after_doc, index, reader), group)
                item = {
                    "index": index,
                    "section": group.section,
                    "page": group.page_label,
                    "summary": group.summary,
                    "context_type": group.context_type,
                    "before": before,
                    "after": after,
                    "image_status": "텍스트",
                }
                reader._report(
                    "변경 영역 이미지 캡처",
                    80 + 10 * (index - 1) / max(1, len(groups)),
                    f"{index}/{len(groups)} 변경 전 · {group.context_type} · page {group.page_label}",
                    index,
                    len(groups),
                )
                before_image = images_dir / f"{index:04d}_before.png"
                after_image = images_dir / f"{index:04d}_after.png"
                try:
                    _capture_range_png(before_doc, index, reader, before_image)
                    reader._report(
                        "변경 영역 이미지 캡처",
                        80 + 10 * (index - 0.5) / max(1, len(groups)),
                        f"{index}/{len(groups)} 변경 후 · {group.context_type} · page {group.page_label}",
                        index,
                        len(groups),
                    )
                    _capture_range_png(after_doc, index, reader, after_image)
                    item["before_image"] = str(before_image)
                    item["after_image"] = str(after_image)
                    item["image_status"] = "이미지"
                    captured_images.extend((before_image, after_image))
                except Exception as exc:
                    item["image_status"] = "텍스트 대체"
                    item["image_error"] = str(exc)
                    logger.warning(
                        "변경 영역 이미지 캡처 실패 · 항목 %s · 형식 %s · page %s · 텍스트로 대체: %s",
                        index,
                        group.context_type,
                        group.page_label,
                        exc,
                    )
                items.append(item)
                reader._report(
                    "페이지별 미리보기 추출",
                    80 + 10 * index / max(1, len(groups)),
                    f"page {group.page_label} · {group.section}",
                    index,
                    len(groups),
                )
        finally:
            reader._close_document(before_doc, "변경 전 복원본")
            reader._close_document(after_doc, "변경 후 복원본")
            reader._close_word(None, word)

        progress({"kind": "progress", "stage": "미리보기 Excel 생성", "percent": 92, "detail": preview_path.name})
        _write_preview_xlsx(preview_path, items, document_name)
        manifest_items = []
        for item in items:
            stored = dict(item)
            for key in ("before_image", "after_image"):
                value = stored.get(key)
                if value:
                    stored[key] = "images/" + Path(value).name
            manifest_items.append(stored)
        manifest = {
            "format": PACKAGE_FORMAT,
            "version": VERSION,
            "document_name": document_name,
            "source_name": source_path.name,
            "item_count": len(items),
            "items": manifest_items,
            "groups": [asdict(group) for group in groups],
        }
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        package_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(package_path, "w", zipfile.ZIP_DEFLATED) as archive:
            for file_path in (manifest_path, preview_path, before_path, after_path):
                archive.write(file_path, file_path.name)
            for image_path in captured_images:
                archive.write(image_path, "images/" + image_path.name)
        preview_output = package_path.with_name(package_path.stem + "_변경내용_미리보기.xlsx")
        before_output = package_path.with_name(package_path.stem + "_변경전_전체.docx")
        after_output = package_path.with_name(package_path.stem + "_변경후_전체.docx")
        shutil.copy2(preview_path, preview_output)
        shutil.copy2(before_path, before_output)
        shutil.copy2(after_path, after_output)
        progress({"kind": "progress", "stage": "분석 패키지 저장", "percent": 99, "detail": package_path.name})
    return len(groups)


def read_package_manifest(package_path: Path) -> dict:
    with zipfile.ZipFile(package_path) as archive:
        if archive.testzip() is not None or "analysis.json" not in archive.namelist():
            raise ValueError("올바른 변대생기 분석 패키지가 아닙니다.")
        manifest = json.loads(archive.read("analysis.json").decode("utf-8"))
    if manifest.get("format") != PACKAGE_FORMAT:
        raise ValueError("지원하지 않는 분석 패키지 형식입니다.")
    return manifest


def _tokens(text: str) -> list[str]:
    return re.findall(r"\s+|[가-힣A-Za-z0-9_]+|[^\w\s]", text, re.UNICODE)


def _write_diff_cell(cell, before: str, after: str, side: str) -> None:
    from docx.shared import Pt, RGBColor

    before_tokens, after_tokens = _tokens(before), _tokens(after)
    matcher = difflib.SequenceMatcher(a=before_tokens, b=after_tokens, autojunk=False)
    for tag, a1, a2, b1, b2 in matcher.get_opcodes():
        tokens = before_tokens[a1:a2] if side == "before" else after_tokens[b1:b2]
        if not tokens:
            continue
        run = cell.paragraphs[0].add_run("".join(tokens))
        run.font.name = "맑은 고딕"
        run.font.size = Pt(7 if "표 행 생략" in run.text else 8)
        if side == "before" and tag in {"delete", "replace"}:
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.strike = True
        elif side == "after" and tag in {"insert", "replace"}:
            run.font.color.rgb = RGBColor(0, 0, 255)
            run.font.underline = True


def generate_comparison_from_package(package_path: Path, output_path: Path, progress) -> int:
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor

    manifest = read_package_manifest(package_path)
    items = list(manifest.get("items") or [])
    if not items:
        raise ValueError("분석 패키지에 변경 항목이 없습니다.")
    document = Document()
    section = document.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = section.bottom_margin = Cm(1.2)
    section.left_margin = section.right_margin = Cm(1.2)
    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("변경대비표")
    run.bold = True
    run.font.name = "맑은 고딕"
    run.font.size = Pt(16)
    info = document.add_paragraph(f"문서명: {manifest.get('document_name', '')}")
    info.runs[0].font.name = "맑은 고딕"
    info.runs[0].font.size = Pt(9)

    table = document.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    table.autofit = False
    widths = [Cm(2.0), Cm(1.3), Cm(10.4), Cm(10.4), Cm(2.0)]
    headers = ["Section", "page", "변경 전", "변경 후", "변경 사유"]
    header_row = table.rows[0]
    header_row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    for index, (cell, label, width) in enumerate(zip(header_row.cells, headers, widths)):
        cell.width = width
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        header_run = cell.paragraphs[0].add_run(label)
        header_run.bold = True
        header_run.font.name = "맑은 고딕"
        header_run.font.size = Pt(9)
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), "D9EAF7")
        cell._tc.get_or_add_tcPr().append(shading)

    for index, item in enumerate(items, start=1):
        row = table.add_row()
        for cell, width in zip(row.cells, widths):
            cell.width = width
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            cell.paragraphs[0].paragraph_format.space_after = Pt(0)
        for column, text in ((0, item["section"]), (1, item["page"])):
            row.cells[column].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            value_run = row.cells[column].paragraphs[0].add_run(str(text))
            value_run.font.name = "맑은 고딕"
            value_run.font.size = Pt(8)
        _write_diff_cell(row.cells[2], item["before"], item["after"], "before")
        _write_diff_cell(row.cells[3], item["before"], item["after"], "after")
        progress({
            "kind": "progress",
            "stage": "변경대비표 작성",
            "percent": 5 + 90 * index / len(items),
            "detail": f"page {item['page']} · {item['section']}",
            "current": index,
            "total": len(items),
        })
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    progress({"kind": "progress", "stage": "변경대비표 저장", "percent": 99, "detail": output_path.name})
    return len(items)


class SimpleTool:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.mode = "analyze"
        self.events: queue.Queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.busy = False
        self.input_var = StringVar()
        self.output_var = StringVar()
        self.status_var = StringVar(value="입력 파일을 선택하십시오.")
        self.stage_var = StringVar(value="대기 중")
        self.detail_var = StringVar(value="진행할 작업이 없습니다.")
        self.elapsed_var = StringVar(value="경과 시간 00:00")
        self.progress_var = DoubleVar(value=0)
        self.last_log_path: Path | None = None
        self.task_started_at = 0.0
        self.last_progress_at = 0.0
        title = "변분기 - Alpha 0.2"
        root.title(title)
        root.geometry("860x440")
        root.minsize(760, 400)
        frame = ttk.Frame(root, padding=18)
        frame.pack(fill="both", expand=True)
        description = (
            "1단계: tracked DOCX를 분석하여 미리보기 Excel과 변경 전·후 Word가 포함된 분석 패키지를 만듭니다."
            if self.mode == "analyze"
            else "2단계: 1단계에서 만든 분석 패키지로 최종 Word 변경대비표를 만듭니다."
        )
        ttk.Label(frame, text=description, wraplength=760).pack(fill="x", pady=(0, 18))
        self._file_row(frame, "입력", self.input_var, self.choose_input)
        self._file_row(frame, "출력", self.output_var, self.choose_output)
        self.run_button = ttk.Button(
            frame,
            text="변경 분석 및 미리보기 생성",
            command=self.run,
        )
        self.run_button.pack(pady=18)
        progress_frame = ttk.LabelFrame(frame, text="진행 상황", padding=10)
        progress_frame.pack(fill="x")
        progress_top = ttk.Frame(progress_frame)
        progress_top.pack(fill="x")
        ttk.Label(progress_top, textvariable=self.stage_var).pack(side="left")
        ttk.Label(progress_top, textvariable=self.elapsed_var).pack(side="right")
        ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100).pack(fill="x", pady=7)
        progress_bottom = ttk.Frame(progress_frame)
        progress_bottom.pack(fill="x")
        ttk.Label(progress_bottom, textvariable=self.detail_var, wraplength=650).pack(side="left", fill="x", expand=True)
        self.log_button = ttk.Button(progress_bottom, text="로그 열기", command=self.open_log, state="disabled")
        self.log_button.pack(side="right", padx=(8, 0))
        ttk.Label(frame, textvariable=self.status_var, wraplength=800).pack(fill="x", pady=(12, 0))

    @staticmethod
    def _file_row(parent, label, variable, command) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=5)
        ttk.Label(row, text=label, width=7).pack(side="left")
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="찾아보기", command=command).pack(side="right")

    def choose_input(self) -> None:
        if self.mode == "analyze":
            name = filedialog.askopenfilename(filetypes=[("Word 문서", "*.docx")])
            if name:
                self.input_var.set(name)
                self.output_var.set(str(Path(name).with_suffix(PACKAGE_EXTENSION)))
        else:
            name = filedialog.askopenfilename(filetypes=[("변대생기 분석 패키지", "*.bdsg")])
            if name:
                self.input_var.set(name)
                self.output_var.set(str(Path(name).with_name(Path(name).stem + "_변경대비표.docx")))

    def choose_output(self) -> None:
        if self.mode == "analyze":
            name = filedialog.asksaveasfilename(defaultextension=".bdsg", filetypes=[("변대생기 분석 패키지", "*.bdsg")])
        else:
            name = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word 문서", "*.docx")])
        if name:
            self.output_var.set(name)

    def run(self) -> None:
        source, output = Path(self.input_var.get()), Path(self.output_var.get())
        if not source.exists() or not output.name:
            messagebox.showwarning(self.root.title(), "입력 및 출력 파일을 확인하십시오.")
            return
        self.busy = True
        self.run_button.configure(state="disabled")
        self.progress_var.set(0)
        self.stage_var.set("작업 준비")
        self.detail_var.set(source.name)
        self.task_started_at = time.monotonic()
        self.last_progress_at = self.task_started_at
        self.elapsed_var.set("경과 시간 00:00")
        self.status_var.set("작업을 시작했습니다.")
        logger, log_path = create_operation_logger("변분기_분석")
        self.last_log_path = log_path
        self.log_button.configure(state="normal")

        def record_progress(event: dict) -> None:
            if event.get("kind") == "progress":
                logger.info(
                    "%s %.1f%% (%s/%s) %s",
                    event.get("stage", "처리 중"),
                    float(event.get("percent", 0)),
                    event.get("current", "-"),
                    event.get("total", "-"),
                    event.get("detail", ""),
                )
            self.events.put(event)

        def worker() -> None:
            try:
                if self.mode == "analyze":
                    logger.info("작업 시작 · 입력: %s · 출력: %s", source, output)
                    count = create_analysis_package(source, output, record_progress, self.cancel_event, logger)
                    logger.info("작업 완료 · 변경 항목 %s개", count)
                else:
                    count = generate_comparison_from_package(source, output, self.events.put)
                self.events.put({"kind": "done", "count": count, "output": output})
            except Exception as exc:
                logger.exception("작업 오류")
                self.events.put({"kind": "error", "message": str(exc), "traceback": traceback.format_exc(), "log": log_path})
            finally:
                close_operation_logger(logger)

        threading.Thread(target=worker, daemon=True).start()
        self.root.after(120, self.poll)

    def poll(self) -> None:
        terminal = None
        while True:
            try:
                event = self.events.get_nowait()
            except queue.Empty:
                break
            if event.get("kind") == "progress":
                self.last_progress_at = time.monotonic()
                self.progress_var.set(max(self.progress_var.get(), float(event.get("percent", 0))))
                detail = str(event.get("detail", ""))
                current, total = event.get("current"), event.get("total")
                if current is not None and total:
                    detail = f"{current}/{total} · {detail}"
                self.stage_var.set(str(event.get("stage", "처리 중")))
                self.detail_var.set(detail)
            else:
                terminal = event
        if terminal:
            self.busy = False
            self.run_button.configure(state="normal")
            if terminal["kind"] == "done":
                self.progress_var.set(100)
                self.stage_var.set("분석 완료")
                self.detail_var.set(f"변경 항목 {terminal['count']}개 · 분석 패키지 저장 완료")
                output = Path(terminal["output"])
                preview_output = output.with_name(output.stem + "_변경내용_미리보기.xlsx")
                self.status_var.set(f"완료: {output}")
                messagebox.showinfo(
                    self.root.title(),
                    "작업이 완료되었습니다.\n\n"
                    f"변경 항목: {terminal['count']}개\n"
                    f"미리보기: {preview_output}\n"
                    f"분석 패키지: {output}",
                )
                with suppress(OSError):
                    os.startfile(preview_output if preview_output.exists() else output)
            else:
                self.stage_var.set("작업 오류")
                self.detail_var.set("로그 열기에서 마지막 완료 단계와 오류 상세를 확인하십시오.")
                self.status_var.set("작업 중 오류가 발생했습니다.")
                messagebox.showerror(self.root.title(), f"{terminal['message']}\n\n로그: {terminal['log']}")
            return
        elapsed = max(0, int(time.monotonic() - self.task_started_at))
        minutes, seconds = divmod(elapsed, 60)
        self.elapsed_var.set(f"경과 시간 {minutes:02d}:{seconds:02d}")
        silent = time.monotonic() - self.last_progress_at
        if silent >= 120:
            self.status_var.set(
                f"마지막 진행 기록 후 {int(silent)}초가 지났습니다. 로그 열기에서 현재 단계를 확인하십시오."
            )
        self.root.after(120, self.poll)

    def open_log(self) -> None:
        if self.last_log_path and self.last_log_path.exists():
            with suppress(OSError):
                os.startfile(self.last_log_path)


def main() -> int:
    if sys.platform != "win32":
        print("변분기는 Microsoft Word가 설치된 Windows에서 실행하십시오.")
        return 2
    root = Tk()
    SimpleTool(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
