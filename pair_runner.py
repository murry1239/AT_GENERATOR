"""Orchestrate DOCX pair analysis and retain independent review outputs."""
from __future__ import annotations

import json
import os
import shutil
import tempfile
import time
import zipfile
from contextlib import contextmanager
from pathlib import Path

from app import WordEngine, OperationCancelled
from pair_analysis import read_snapshot, compare_snapshots, tracked_advice, bookmark_copy, validate_paths, digest


def create_package(before_path, after_path, tracked_path, package_path, progress, cancel_event, logger,
                   *, preview_writer, capture, bookmark_range, version, package_format):
    package_path = Path(package_path).resolve()
    validate_paths(before_path, after_path, tracked_path, package_path)
    timings = {}
    warnings = []

    def report(stage, percent, detail="", **kw):
        if cancel_event.is_set():
            raise OperationCancelled("사용자가 작업을 취소했습니다.")
        logger.info("%s %.1f%% %s", stage, percent, detail)
        progress(dict(kind="progress", stage=stage, percent=percent, detail=detail, **kw))

    @contextmanager
    def timed(stage):
        started = time.monotonic()
        logger.info("단계 시작: %s", stage)
        try:
            yield
        except OperationCancelled:
            logger.info("단계 취소: %s", stage)
            raise
        except Exception:
            logger.exception("단계 실패: %s", stage)
            raise
        finally:
            timings[stage] = round(time.monotonic()-started, 3)
            logger.info("단계 소요시간: %s %.3f초", stage, timings[stage])

    with timed("DOCX 읽기"):
        report("DOCX 읽기", 2, "변경 전")
        before = read_snapshot(before_path)
        report("DOCX 읽기", 8, "변경 후")
        after = read_snapshot(after_path)
    with timed("전후 변경 분석"):
        report("전후 변경 분석", 15)
        pairs, warnings = compare_snapshots(before, after)
        report("전후 변경 분석", 25, f"검토 블록 {len(pairs)}개")
    with timed("선택 이력 확인"):
        report("선택 이력 확인", 28)
        if tracked_path:
            try:
                advice = tracked_advice(tracked_path, before, after)
            except Exception as exc:
                logger.exception("선택 이력 읽기 실패: 전후 파일로 계속 진행")
                advice = {"status": "unreadable_advisory_disabled", "revision_count": 0, "records": [], "error_type": type(exc).__name__}
            if advice["status"] != "text_match_advisory_only":
                warnings.append("선택한 이력 파일의 일치 여부를 확인하지 못해 보조 활용을 중단했습니다. 변경 전·후 파일로 분석했습니다.")
        else:
            advice = tracked_advice(None, before, after)
    for warning in warnings:
        logger.warning(warning)
    package_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="bdsg_pair_", ignore_cleanup_errors=True) as temporary:
        temp = Path(temporary)
        working_before, working_after = temp / "before_work.docx", temp / "after_work.docx"
        items = []
        reader = word = before_doc = after_doc = None
        try:
            if pairs:
                with timed("이미지용 복사본 준비"):
                    report("이미지용 복사본 준비", 32, "원본 변경 및 승인·거부 없이 북마크 추가")
                    bookmark_copy(before, pairs, "before", working_before)
                    bookmark_copy(after, pairs, "after", working_after)
                with timed("Word 열기 및 페이지 계산"):
                    reader = WordEngine(progress=progress, cancel_event=cancel_event, logger=logger)
                    word = reader._start_word("Word 열기", 36)
                    before_doc = reader._retry_word_call("변경 전 열기", lambda: word.Documents.Open(str(working_before), ReadOnly=True, Visible=False, AddToRecentFiles=False))
                    after_doc = reader._retry_word_call("변경 후 열기", lambda: word.Documents.Open(str(working_after), ReadOnly=True, Visible=False, AddToRecentFiles=False))
                    for doc in (before_doc, after_doc):
                        # Render clean content; comment balloons must not narrow pages.
                        doc.ShowRevisions = False
                        reader._retry_word_call("페이지 계산", lambda d=doc: d.Repaginate())
                with timed("원문 및 이미지 추출"):
                    for index, pair in enumerate(pairs, 1):
                        old, new = pair["before"], pair["after"]
                        item = {"index": index, "change_id": f"change-{index:06d}",
                                "section": (new or old).section,
                                "context_type": "table" if any(b is not None and b.kind == "table" for b in (old, new)) else "paragraph",
                                "change_type": pair["change_type"], "needs_review": pair["needs_review"],
                                "match_method": pair["match_method"], "summary": "서식 영향 검토" if pair["format_review"] else "",
                                "before": old.text if old else "", "after": new.text if new else "",
                                "before_block": old.index if old else None, "after_block": new.index if new else None}
                        errors = []
                        for side, doc, block in (("before", before_doc, old), ("after", after_doc, new)):
                            report("변경 영역 이미지 캡처", 42+43*((index-1)+(0.5 if side == "after" else 0))/len(pairs),
                                   f"{side} · {item['context_type']} · {item['section']}", current=index, total=len(pairs))
                            item[side+"_page"] = "없음" if block is None else "확인 실패"
                            if block is None:
                                item[side+"_status"] = "해당 없음"
                                continue
                            try:
                                rng = bookmark_range(doc, index, reader)
                                # Physical page number avoids ambiguity from restarted numbering.
                                start = int(doc.Range(rng.Start, rng.Start).Information(3))
                                endpos = max(int(rng.Start), int(rng.End)-1)
                                end = int(doc.Range(endpos, endpos).Information(3))
                                item[side+"_page"] = str(start) if start == end else f"{start}-{end}"
                                image_path = temp / "images" / f"{index:04d}_{side}.png"
                                if end != start:
                                    raise ValueError("여러 페이지에 걸친 블록: 잘린 이미지 방지를 위해 원문으로 대체")
                                capture(doc, index, reader, image_path)
                                item[side+"_image"] = str(image_path)
                                item[side+"_status"] = "이미지"
                            except OperationCancelled:
                                raise
                            except Exception as exc:
                                item[side+"_status"] = "텍스트 대체"
                                item["needs_review"] = True
                                errors.append(f"{side}: {type(exc).__name__}: {str(exc)[:160]}")
                                logger.exception("캡처 실패: 항목 %s %s page %s", index, side, item[side+"_page"])
                        item["page"] = f"전 {item['before_page']} / 후 {item['after_page']}"
                        item["image_status"] = "텍스트 대체" if errors else "이미지"
                        item["image_error"] = "\n".join(errors)
                        items.append(item)
        finally:
            if reader:
                reader._close_document(before_doc, "변경 전")
                reader._close_document(after_doc, "변경 후")
                reader._close_word(None, word)
        fallback_count = sum(i["image_status"] == "텍스트 대체" for i in items)
        if fallback_count:
            warnings.append(f"이미지 실패 또는 여러 페이지 블록 {fallback_count}개는 텍스트로 대체했습니다. 원본 Word를 함께 확인하십시오.")
        with timed("Excel 생성"):
            report("미리보기 Excel 생성", 90)
            preview = temp / "변경내용_미리보기.xlsx"
            preview_writer(preview, items, Path(after_path).stem, warnings=warnings)
        stored_items = []
        for item in items:
            stored = dict(item)
            for key in ("before_image", "after_image"):
                if stored.get(key):
                    stored[key] = "images/" + Path(stored[key]).name
            stored_items.append(stored)
        with timed("패키지 준비"):
            # Retain exact input bytes, not Word-saved or bookmarked variants.
            for snap, name in ((before, "변경전_전체.docx"), (after, "변경후_전체.docx")):
                (temp / name).write_bytes(snap.content)
        manifest = {"format": package_format, "schema_version": 2, "version": version,
                    "input_mode": "pair_with_tracked" if tracked_path else "pair", "document_name": Path(after_path).stem,
                    "sources": {side: {"name": snap.path.name, "sha256": digest(snap.content)} for side, snap in (("before", before), ("after", after))},
                    "scope_before": before.scope, "scope_after": after.scope, "warnings": warnings,
                    "comparison_scope": "본문 블록(목차 이전/부록 이후 자동 제외). 머리말·꼬리말·각주·미주는 차이 경고만 제공.",
                    "tracked_advice": advice, "item_count": len(items), "items": stored_items, "timings_seconds": timings}
        if tracked_path:
            manifest["sources"]["tracked"] = {"name": Path(tracked_path).name, "sha256": digest(Path(tracked_path).read_bytes())}
        (temp / "analysis.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        report("분석 패키지 저장", 96)
        # Stage on the destination volume so a failed ZIP write never replaces
        # an existing complete package with a partial ZIP.
        handle, staged_name = tempfile.mkstemp(dir=package_path.parent, suffix=".tmp")
        os.close(handle)
        try:
            with zipfile.ZipFile(staged_name, "w", zipfile.ZIP_DEFLATED) as archive:
                for name in ("analysis.json", "변경내용_미리보기.xlsx", "변경전_전체.docx", "변경후_전체.docx"):
                    archive.write(temp/name, name)
                for item in items:
                    for key in ("before_image", "after_image"):
                        if item.get(key):
                            p = Path(item[key])
                            archive.write(p, "images/"+p.name)
            report("분석 패키지 저장", 98)
            os.replace(staged_name, package_path)
        finally:
            Path(staged_name).unlink(missing_ok=True)
        for name in ("변경내용_미리보기.xlsx", "변경전_전체.docx", "변경후_전체.docx"):
            shutil.copy2(temp/name, package_path.with_name(package_path.stem+"_"+name))
        report("분석 완료", 100, f"검토 블록 {len(items)}개 · 경고 {len(warnings)}개")
        return {"count": len(items), "warnings": warnings, "timings_seconds": timings}
