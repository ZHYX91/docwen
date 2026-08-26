"""Shared typed OCR API over RapidOCR.

All plugins needing image OCR import from here; no plugin vendors its own.
"""

from __future__ import annotations

import logging
import os
import sys
import threading
from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path
from typing import Any

from docwen_core.formats.categories import CATEGORY_IMAGE, get_category
from docwen_core.text import OCR_LANGUAGE_AUTO, OCR_LANGUAGE_CHINESE, OCR_LANGUAGE_MODELS, resolve_ocr_language

logger = logging.getLogger(__name__)


class OcrStatus(StrEnum):
    """Machine-readable status for one best-effort OCR attempt."""

    SUCCESS = "success"
    NO_TEXT = "no_text"
    INPUT_MISSING = "input_missing"
    UNAVAILABLE = "unavailable"
    MODEL_MISSING = "model_missing"
    INITIALIZATION_FAILED = "initialization_failed"
    RECOGNITION_FAILED = "recognition_failed"


@dataclass(frozen=True, slots=True)
class OcrOutcome:
    """Typed OCR result used when callers need failure observability."""

    status: OcrStatus
    text: str = ""
    message: str = ""

    @property
    def recognized_text(self) -> str:
        """Return text only when the outcome explicitly reports success."""
        return self.text if self.status is OcrStatus.SUCCESS else ""


_OCR_OPERATIONAL_FAILURE_DETAILS: dict[OcrStatus, str] = {
    OcrStatus.INPUT_MISSING: "OCR input file is missing or is not a regular file",
    OcrStatus.UNAVAILABLE: "OCR engine is unavailable",
    OcrStatus.MODEL_MISSING: "OCR model files are missing",
    OcrStatus.INITIALIZATION_FAILED: "OCR engine initialization failed",
    OcrStatus.RECOGNITION_FAILED: "OCR recognition failed",
}

_OCR_RESULT_QUALITY_DETAILS: dict[OcrStatus, str] = {
    OcrStatus.SUCCESS: (
        "OCR text is machine-generated and may contain recognition errors or omissions; verify it against the source"
    ),
    OcrStatus.NO_TEXT: "OCR detected no text; text may have been missed, so verify it against the source",
}


def format_ocr_best_effort_warning(status: object, *, context: str = "") -> str | None:
    """Return the canonical safe warning for a fallible best-effort OCR outcome."""
    try:
        normalized = status if isinstance(status, OcrStatus) else OcrStatus(str(status))
    except ValueError:
        return None
    suffix = f"; {context}" if context else ""
    quality_detail = _OCR_RESULT_QUALITY_DETAILS.get(normalized)
    if quality_detail is not None:
        return f"OCR best-effort result: status={normalized.value}; {quality_detail}{suffix}."
    detail = _OCR_OPERATIONAL_FAILURE_DETAILS.get(normalized)
    if detail is None:
        return None
    return f"OCR best-effort fallback: status={normalized.value}; {detail}{suffix}."


class _OcrModelFilesMissing(FileNotFoundError):
    """Selected model files failed the pre-initialization ownership check."""


class _OcrEngineSlot:
    """A cached OCR engine and the lock protecting its request-time state."""

    __slots__ = ("engine", "failure_message", "failure_status", "invocation_lock")

    def __init__(
        self,
        engine: Any,
        failure_status: OcrStatus | None = None,
        failure_message: str = "",
    ) -> None:
        self.engine = engine
        self.failure_status = failure_status
        self.failure_message = failure_message
        self.invocation_lock = threading.Lock()


_ocr_instances: dict[tuple[str, str], _OcrEngineSlot] = {}
_ocr_lock = threading.Lock()


def _configure_onnxruntime_logging() -> None:
    """Keep native inference diagnostics off protocol stderr unless they are errors."""
    try:
        import onnxruntime
    except ImportError:
        return
    onnxruntime.set_default_logger_severity(3)


def _normalize_ocr_source_format(source_format: str) -> str:
    """Validate one content-derived concrete image format for OCR."""
    normalized = str(source_format).strip().lower()
    if not normalized or normalized == CATEGORY_IMAGE or get_category(normalized) != CATEGORY_IMAGE:
        raise ValueError("source_format must be a concrete image format")
    return normalized


def _unsupported_source_outcome(source_format: str) -> OcrOutcome | None:
    """Return the typed pre-conversion requirement for unsupported inputs."""
    if source_format not in {"heic", "heif"}:
        return None
    return OcrOutcome(
        OcrStatus.UNAVAILABLE,
        message=(f"OCR does not accept {source_format.upper()} directly; pre-convert the admitted image before OCR"),
    )


def _get_ocr_slot(
    ocr_language: str | None = None,
    *,
    current_locale: str = "zh_CN",
    model_dir: str | Path | None = None,
) -> _OcrEngineSlot:
    """Return the cached engine slot for one effective language/model key."""
    target_language = resolve_ocr_language(ocr_language or OCR_LANGUAGE_AUTO, current_locale)
    if target_language not in OCR_LANGUAGE_MODELS:
        logger.warning("Unknown OCR language %r; falling back to %s.", target_language, OCR_LANGUAGE_CHINESE)
        target_language = OCR_LANGUAGE_CHINESE

    resolved_model_dir = _resolve_model_dir(model_dir)
    cache_key = (target_language, str(resolved_model_dir))
    with _ocr_lock:
        if cache_key in _ocr_instances:
            return _ocr_instances[cache_key]

        try:
            from rapidocr_onnxruntime import RapidOCR

            _configure_onnxruntime_logging()

            det_model_path, rec_model_path, cls_model_path = _model_paths_for_language(
                target_language,
                resolved_model_dir,
            )
            missing_models = [path for path in (det_model_path, rec_model_path, cls_model_path) if not path.exists()]
            if missing_models:
                missing = ", ".join(str(path) for path in missing_models)
                raise _OcrModelFilesMissing(f"OCR models missing for language {target_language}: {missing}")

            engine: Any = RapidOCR(
                det_model_path=str(det_model_path),
                rec_model_path=str(rec_model_path),
                cls_model_path=str(cls_model_path),
            )
            logger.info("RapidOCR initialised for %s with models at %s.", target_language, resolved_model_dir)
        except ImportError:
            failure_status = OcrStatus.UNAVAILABLE
            failure_message = "rapidocr_onnxruntime not installed"
            logger.warning("%s.", failure_message)
            engine = False
        except _OcrModelFilesMissing as exc:
            failure_status = OcrStatus.MODEL_MISSING
            failure_message = str(exc)
            logger.warning("%s", failure_message)
            engine = False
        except Exception as exc:
            failure_status = OcrStatus.INITIALIZATION_FAILED
            failure_message = str(exc)
            logger.warning("RapidOCR init failed: %s", failure_message)
            engine = False
        else:
            failure_status = None
            failure_message = ""
        slot = _OcrEngineSlot(engine, failure_status, failure_message)
        _ocr_instances[cache_key] = slot
        return slot


def _get_ocr(
    ocr_language: str | None = None,
    *,
    current_locale: str = "zh_CN",
    model_dir: str | Path | None = None,
) -> Any:
    """Return a cached RapidOCR instance; ``False`` if unavailable."""
    return _get_ocr_slot(
        ocr_language,
        current_locale=current_locale,
        model_dir=model_dir,
    ).engine


def _resolve_model_dir(model_dir: str | Path | None = None) -> Path:
    """Resolve the ``models/rapidocr`` directory without importing runtime."""
    candidates: list[Path] = []
    if model_dir is not None:
        candidates.append(Path(model_dir))

    for env_name in ("DOCWEN_RAPIDOCR_MODEL_DIR", "DOCWEN_OCR_MODEL_DIR", "DOCWEN_RESOURCE_ROOT"):
        env_value = os.environ.get(env_name)
        if env_value:
            candidates.append(Path(env_value))

    meipass = getattr(sys, "_MEIPASS", None)
    if isinstance(meipass, str) and meipass:
        internal = Path(meipass)
        candidates.extend([internal, internal.parent])

    candidates.append(Path.cwd())
    candidates.extend(Path(__file__).resolve().parents)

    executable = getattr(sys, "executable", "")
    if executable:
        candidates.append(Path(executable).resolve().parent)

    normalized: list[Path] = []
    for candidate in candidates:
        normalized.extend(_model_dir_candidates(candidate))

    for candidate in normalized:
        if candidate.exists() and any(candidate.glob("*.onnx")):
            return candidate

    for candidate in normalized:
        if candidate.exists() and candidate.name == "rapidocr" and candidate.parent.name == "models":
            return candidate

    for candidate in normalized:
        if candidate.name == "rapidocr" and candidate.parent.name == "models":
            return candidate

    return normalized[0] if normalized else Path("models") / "rapidocr"


def _model_dir_candidates(candidate: Path) -> list[Path]:
    candidate = candidate.expanduser()
    return [
        candidate,
        candidate / "rapidocr",
        candidate / "models" / "rapidocr",
    ]


def _model_paths_for_language(language: str, model_dir: Path) -> tuple[Path, Path, Path]:
    models = OCR_LANGUAGE_MODELS.get(language, OCR_LANGUAGE_MODELS[OCR_LANGUAGE_CHINESE])
    return (
        model_dir / models["det"],
        model_dir / models["rec"],
        model_dir / models["cls"],
    )


def ocr_available(
    ocr_language: str | None = None,
    *,
    current_locale: str = "zh_CN",
    model_dir: str | Path | None = None,
) -> bool:
    """Return whether the engine for the requested language/model key is ready."""
    try:
        return (
            _get_ocr(
                ocr_language,
                current_locale=current_locale,
                model_dir=model_dir,
            )
            is not False
        )
    except FileNotFoundError:
        return False


def run_ocr_outcome(
    image_path: str | Path,
    ocr_language: str | None = None,
    *,
    source_format: str,
    current_locale: str = "zh_CN",
    model_dir: str | Path | None = None,
) -> OcrOutcome:
    """Run OCR once and return a typed status without folding failures into text.

    Native HEIC/HEIF input returns :attr:`OcrStatus.UNAVAILABLE` so a caller
    can pre-convert it explicitly.  A misleading HEIC filename has no effect
    when ``source_format`` identifies supported image content.
    """
    normalized_source_format = _normalize_ocr_source_format(source_format)
    image_file = Path(image_path)
    try:
        input_is_file = image_file.is_file()
    except OSError as exc:
        return OcrOutcome(OcrStatus.RECOGNITION_FAILED, message=str(exc))
    if not input_is_file:
        return OcrOutcome(
            OcrStatus.INPUT_MISSING,
            message=f"OCR input file is missing or is not a regular file: {image_file}",
        )
    if unsupported := _unsupported_source_outcome(normalized_source_format):
        return unsupported
    try:
        slot = _get_ocr_slot(
            ocr_language,
            current_locale=current_locale,
            model_dir=model_dir,
        )
    except _OcrModelFilesMissing as exc:
        message = str(exc)
        logger.warning("%s", message)
        return OcrOutcome(OcrStatus.MODEL_MISSING, message=message)
    except Exception as exc:
        message = str(exc)
        logger.warning("OCR initialization failed: %s", message)
        return OcrOutcome(OcrStatus.INITIALIZATION_FAILED, message=message)

    engine = slot.engine
    if engine is False:
        return OcrOutcome(
            slot.failure_status or OcrStatus.UNAVAILABLE,
            message=slot.failure_message,
        )
    try:
        with slot.invocation_lock:
            result, _elapsed = engine(str(image_file))
        if result is None:
            return OcrOutcome(OcrStatus.NO_TEXT)
        lines = [text for item in result if (text := _extract_result_text(item))]
        if not lines:
            return OcrOutcome(OcrStatus.NO_TEXT)
        return OcrOutcome(OcrStatus.SUCCESS, text="\n".join(lines))
    except Exception as exc:
        logger.warning("OCR failed for %s: %s", image_file, exc)
        return OcrOutcome(OcrStatus.RECOGNITION_FAILED, message=str(exc))


def _extract_result_text(item: Any) -> str:
    """Extract trusted text from known RapidOCR result item shapes."""
    if not isinstance(item, (list, tuple)):
        return ""
    if len(item) >= 3 and isinstance(item[1], str):
        text = item[1]
        confidence = item[2]
    elif len(item) >= 2 and isinstance(item[0], str):
        text = item[0]
        confidence = item[1]
    else:
        return ""

    if isinstance(confidence, (int, float)):
        confidence_value = float(confidence)
    elif isinstance(confidence, str):
        try:
            confidence_value = float(confidence)
        except ValueError:
            return ""
    else:
        return ""
    # The OCR admission contract accepts only scores above 0.5. Keep this as
    # a positive allow condition so non-finite NaN values are rejected too.
    # RapidOCR filters low scores by default, but retaining the explicit owner
    # boundary prevents custom engine noise from leaking downstream.
    if confidence_value > 0.5:
        return text.strip()
    return ""


def reset_ocr() -> None:
    """Drain active cached-engine calls, then invalidate every published slot."""
    with _ocr_lock:
        slots = tuple(_ocr_instances.values())
        acquired_locks: list[threading.Lock] = []
        try:
            for slot in slots:
                slot.invocation_lock.acquire()
                acquired_locks.append(slot.invocation_lock)
            _ocr_instances.clear()
        finally:
            for invocation_lock in reversed(acquired_locks):
                invocation_lock.release()
