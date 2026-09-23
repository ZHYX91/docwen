"""CPU-only PicoDet table localisation, adapted from RapidLayout v1.0.2.

The preprocessing and distribution decoding follow RapidAI/RapidLayout and
PaddleOCR (Apache-2.0). Models are supplied by the caller; no downloads occur.
"""

from __future__ import annotations

import threading
from pathlib import Path

import cv2
import numpy as np
import onnxruntime

_HEIGHT, _WIDTH = 800, 608
_LOCK = threading.Lock()
_SESSIONS: dict[str, onnxruntime.InferenceSession] = {}


def add_borderless_candidates(
    rectangles: list[tuple[int, int, int, int]], boxes: np.ndarray, image_shape: tuple[int, int]
) -> list[tuple[int, int, int, int]]:
    """Supplement detector misses with three aligned multi-column OCR rows.

    This only proposes crops to SLANet; prose and failed crops remain plain OCR.
    Single-column paragraphs cannot become candidates by this rule.
    """
    low, high = boxes.min(axis=1), boxes.max(axis=1)
    centers = (low + high) / 2
    text_height = max(1.0, float(np.median(high[:, 1] - low[:, 1])))
    rows: list[list[int]] = []
    for index in sorted(range(len(boxes)), key=lambda i: (centers[i, 1], low[i, 0])):
        if not rows or abs(centers[index, 1] - np.mean(centers[rows[-1], 1])) > text_height * 0.5:
            rows.append([])
        rows[-1].append(index)
    groups: list[list[list[int]]] = []
    for row in rows:
        row.sort(key=lambda i: low[i, 0])
        if len(row) < 2:
            groups.append([])
            continue
        previous = groups[-1][-1] if groups and groups[-1] else None
        aligned = (
            previous is not None
            and len(previous) == len(row)
            and all(
                abs(low[first, 0] - low[second, 0]) <= text_height for first, second in zip(previous, row, strict=True)
            )
            and centers[row[0], 1] - centers[previous[0], 1] <= text_height * 4
        )
        if not aligned:
            groups.append([])
        groups[-1].append(row)
    output = list(rectangles)
    for group in groups:
        if len(group) < 3:
            continue
        indexes = [i for row in group for i in row]
        left, top = low[indexes].min(axis=0) - text_height * 0.6
        right, bottom = high[indexes].max(axis=0) + text_height * 0.6
        if any(
            max(0, min(right, x2) - max(left, x1)) * max(0, min(bottom, y2) - max(top, y1))
            > 0.5 * (right - left) * (bottom - top)
            for x1, y1, x2, y2 in output
        ):
            continue
        height, width = image_shape
        output.append(
            (max(0, int(left)), max(0, int(top)), min(width, int(np.ceil(right))), min(height, int(np.ceil(bottom))))
        )
    return output


def refine_ruled_table_bounds(
    image: np.ndarray, rectangles: list[tuple[int, int, int, int]]
) -> list[tuple[int, int, int, int]]:
    """Tighten loose detector boxes only when a large ruled grid proves its extent.

    Word editing guides and blank margins can otherwise be decoded as extra rows.
    This does not trim by text, so genuinely empty edge cells remain present.
    """
    height, width = image.shape[:2]
    result = []
    for left, top, right, bottom in rectangles:
        x, y = max(0, left - 8), max(0, top - 8)
        crop = image[y : min(height, bottom + 8), x : min(width, right + 8)]
        h, w = crop.shape[:2]
        if min(h, w) < 20:
            result.append((left, top, right, bottom))
            continue
        gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
        ink = cv2.threshold(gray, 170, 255, cv2.THRESH_BINARY_INV)[1]
        horizontal = cv2.morphologyEx(ink, cv2.MORPH_OPEN, np.ones((1, max(15, w // 3)), np.uint8))
        vertical = cv2.morphologyEx(ink, cv2.MORPH_OPEN, np.ones((max(10, h // 3), 1), np.uint8))
        rows = np.flatnonzero((horizontal > 0).sum(axis=1) > w * 0.65)
        cols = np.flatnonzero((vertical > 0).sum(axis=0) > h * 0.45)
        row_groups = np.split(rows, np.flatnonzero(np.diff(rows) > 2) + 1)
        col_groups = np.split(cols, np.flatnonzero(np.diff(cols) > 2) + 1)
        if (
            len(row_groups) >= 3
            and len(col_groups) >= 3
            and rows[-1] - rows[0] > h * 0.4
            and cols[-1] - cols[0] > w * 0.65
        ):
            result.append(
                (
                    max(0, x + int(cols[0]) - 4),
                    max(0, y + int(rows[0]) - 4),
                    min(width, x + int(cols[-1]) + 5),
                    min(height, y + int(rows[-1]) + 5),
                )
            )
        else:
            result.append((left, top, right, bottom))
    return result


def _session(model: Path) -> onnxruntime.InferenceSession:
    identity = f"{model.resolve()}:{model.stat().st_mtime_ns}"
    with _LOCK:
        if identity not in _SESSIONS:
            options = onnxruntime.SessionOptions()
            options.log_severity_level = 4
            options.enable_cpu_mem_arena = False
            _SESSIONS[identity] = onnxruntime.InferenceSession(
                str(model), sess_options=options, providers=["CPUExecutionProvider"]
            )
        return _SESSIONS[identity]


def detect_table_regions(image: np.ndarray, model: Path) -> list[tuple[int, int, int, int]]:
    """Return non-overlapping table rectangles in source-image coordinates."""
    height, width = image.shape[:2]
    resized = cv2.resize(image, (_WIDTH, _HEIGHT)).astype(np.float32) / 255.0
    normalized = (resized - np.asarray([0.485, 0.456, 0.406], dtype=np.float32)) / np.asarray(
        [0.229, 0.224, 0.225], dtype=np.float32
    )
    tensor = normalized.transpose(2, 0, 1)[None, ...]
    session = _session(model)
    outputs = session.run(None, {session.get_inputs()[0].name: tensor})
    boxes = _decode([np.asarray(value) for value in outputs])
    return [
        (
            max(0, int(x1 * width / _WIDTH)),
            max(0, int(y1 * height / _HEIGHT)),
            min(width, int(np.ceil(x2 * width / _WIDTH))),
            min(height, int(np.ceil(y2 * height / _HEIGHT))),
        )
        for x1, y1, x2, y2 in boxes
    ]


def _decode(outputs: list[np.ndarray]) -> np.ndarray:
    if len(outputs) != 8:
        raise ValueError("PicoDet table detector returned an incomplete output")
    candidates: list[np.ndarray] = []
    confidences: list[np.ndarray] = []
    for index, stride in enumerate((8, 16, 32, 64)):
        scores = np.asarray(outputs[index])[0].reshape(-1)
        logits = np.asarray(outputs[index + 4])[0]
        bins = logits.shape[-1] // 4
        distribution = logits.reshape(-1, 4, bins)
        distribution = np.exp(distribution - distribution.max(axis=2, keepdims=True))
        distribution /= distribution.sum(axis=2, keepdims=True)
        distances = (distribution * np.arange(bins)).sum(axis=2) * stride
        yy, xx = np.meshgrid(
            np.arange(int(np.ceil(_HEIGHT / stride))), np.arange(int(np.ceil(_WIDTH / stride))), indexing="ij"
        )
        centers = (np.stack((xx, yy, xx, yy), axis=-1).reshape(-1, 4) + 0.5) * stride
        if len(scores) != len(centers):
            raise ValueError("PicoDet table detector geometry does not match its output")
        keep = np.flatnonzero(scores > 0.4)
        candidates.append(centers[keep] + distances[keep] * [-1, -1, 1, 1])
        confidences.append(scores[keep])
    boxes = np.concatenate(candidates)
    scores = np.concatenate(confidences)
    order = np.argsort(scores)[::-1][:1000]
    picked: list[int] = []
    while len(order) and len(picked) < 100:
        first = int(order[0])
        picked.append(first)
        rest = order[1:]
        top_left = np.maximum(boxes[first, :2], boxes[rest, :2])
        bottom_right = np.minimum(boxes[first, 2:], boxes[rest, 2:])
        intersection = np.maximum(0, bottom_right - top_left).prod(axis=1)
        area = np.maximum(0, boxes[:, 2:] - boxes[:, :2]).prod(axis=1)
        overlap = intersection / np.maximum(area[first] + area[rest] - intersection, 1e-6)
        # Nested regions must not consume the same OCR region twice either.
        containment = intersection / np.maximum(np.minimum(area[first], area[rest]), 1e-6)
        order = rest[(overlap <= 0.5) & (containment <= 0.8)]
    return boxes[picked]
