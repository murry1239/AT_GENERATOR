"""Bounded Word capture: native EMF first, bitmap clipboard as a fallback."""
from io import BytesIO
from pathlib import Path

from app import OperationCancelled


def save_image(image, output):
    from PIL import Image
    image = image.convert("RGB")
    image.thumbnail((1800, 6000), Image.Resampling.LANCZOS)
    Path(output).parent.mkdir(parents=True, exist_ok=True)
    image.save(output, "PNG")


def save_emf(data, output):
    # Pillow uses Windows' native metafile renderer (no external converter).
    from PIL import Image
    with Image.open(BytesIO(bytes(data))) as image:
        if image.format != "WMF":
            raise ValueError("Word가 EMF 이미지 데이터를 반환하지 않았습니다.")
        width, height = image.size
        if min(width, height) <= 0:
            raise ValueError("빈 EMF 범위입니다.")
        dpi = image.info.get("dpi", 96)
        dx, dy = dpi if isinstance(dpi, tuple) else (dpi, dpi)
        scale = min(2.0, 1800 / width, 6000 / height)
        image.load(dpi=(dx * scale, dy * scale))
        save_image(image, output)


def clipboard_capture(source_range, engine, output):
    from PIL import Image, ImageGrab
    import win32clipboard
    # Never accept stale clipboard contents from the preceding item.
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
    finally:
        win32clipboard.CloseClipboard()
    engine._retry_word_call("그림 클립보드 복사", source_range.CopyAsPicture, attempts=2)
    engine._wait_for_word_retry(0.25, True)
    grabbed = ImageGrab.grabclipboard()
    if not isinstance(grabbed, Image.Image):
        raise RuntimeError("클립보드에 읽을 수 있는 비트맵이 없습니다.")
    save_image(grabbed, output)


def capture_range(source_range, engine, output):
    """Return the successful method for diagnostics; never swallow cancellation."""
    engine._check_cancelled()
    try:
        data = engine._retry_word_call("범위 EMF 직접 추출", lambda: source_range.EnhMetaFileBits, attempts=2)
        engine._check_cancelled()
        save_emf(data, output)
        return "direct_emf"
    except OperationCancelled:
        raise
    except Exception as exc:
        direct_error = f"{type(exc).__name__}: {exc}"
    engine._check_cancelled()
    try:
        clipboard_capture(source_range, engine, output)
        return "clipboard_bitmap"
    except OperationCancelled:
        raise
    except Exception as exc:
        raise RuntimeError(f"그림 추출 실패 · EMF: {direct_error}; 클립보드: {type(exc).__name__}: {exc}") from exc
