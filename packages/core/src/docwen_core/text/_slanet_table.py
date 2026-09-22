"""Minimal SLANet table-structure runtime adapted from RapidTable/PaddleOCR.

This module intentionally carries only the CPU ONNX inference path needed by
DocWen. It avoids RapidTable's downloader, logging/configuration stack and its
second OCR engine: DocWen supplies already-admitted OCR geometry and text.

Portions are adapted from RapidAI/RapidTable v2.0.3 and PaddleOCR code under
the Apache License 2.0. See LICENSE_THIRD_PARTY.txt.
"""

from __future__ import annotations

import threading
from pathlib import Path
from typing import Any

import cv2
import numpy as np
import onnxruntime

_MODEL_SIZE = 488
_MEAN = np.array([0.485, 0.456, 0.406], dtype=np.float32)
_STD = np.array([0.229, 0.224, 0.225], dtype=np.float32)


class _SLANetEngine:
    def __init__(self, model_path: Path) -> None:
        options = onnxruntime.SessionOptions()
        options.log_severity_level = 4
        options.enable_cpu_mem_arena = False
        options.graph_optimization_level = onnxruntime.GraphOptimizationLevel.ORT_ENABLE_ALL
        self._session = onnxruntime.InferenceSession(
            str(model_path),
            sess_options=options,
            providers=["CPUExecutionProvider"],
        )
        self._input_names = [value.name for value in self._session.get_inputs()]
        self._output_names = [value.name for value in self._session.get_outputs()]
        metadata = self._session.get_modelmeta().custom_metadata_map
        raw_characters = metadata.get("character", "")
        if not raw_characters:
            raise ValueError("RapidTable SLANet model does not expose its character metadata")
        self._characters = _prepare_characters(raw_characters.splitlines())

    def infer(self, image: np.ndarray) -> tuple[list[str], np.ndarray]:
        tensor, shape = _preprocess(image)
        outputs = self._session.run(
            self._output_names,
            dict(zip(self._input_names, [tensor], strict=True)),
        )
        if len(outputs) < 2:
            raise ValueError("RapidTable SLANet model returned an incomplete output")
        bbox_preds = np.asarray(outputs[0])
        structure_probs = np.asarray(outputs[1])
        structures, cell_bboxes = _decode(bbox_preds, structure_probs, shape, self._characters)
        return structures, _rescale_and_filter_bboxes(image, cell_bboxes)


_ENGINE_LOCK = threading.Lock()
_ENGINE_CACHE: dict[str, _SLANetEngine] = {}


def _engine(model_path: Path) -> _SLANetEngine:
    key = str(model_path.resolve())
    with _ENGINE_LOCK:
        cached = _ENGINE_CACHE.get(key)
        if cached is None:
            cached = _SLANetEngine(model_path)
            _ENGINE_CACHE[key] = cached
        return cached


def reset_table_engine_cache() -> None:
    with _ENGINE_LOCK:
        _ENGINE_CACHE.clear()


def infer_table_html(
    image_path: str | Path,
    model_path: str | Path,
    *,
    boxes: np.ndarray,
    texts: tuple[str, ...],
    scores: tuple[float, ...],
) -> str:
    """Infer one table and merge existing OCR text into predicted cells."""

    if len(texts) != len(scores) or len(texts) != len(boxes):
        raise ValueError("OCR table inputs must have equal cardinality")

    image = _load_image(Path(image_path))
    structures, cell_bboxes = _engine(Path(model_path)).infer(image)
    if not structures or cell_bboxes.size == 0:
        return ""

    dt_boxes, rec_res = _format_ocr(boxes, texts, scores, image.shape[:2])
    dt_boxes, rec_res = _filter_caption_ocr(cell_bboxes, dt_boxes, rec_res)
    if len(dt_boxes) == 0:
        return ""

    matched = _match_result(dt_boxes, cell_bboxes)
    return _fill_structure(structures, matched, rec_res)


def _load_image(path: Path) -> np.ndarray:
    payload = np.fromfile(path, dtype=np.uint8)
    image = cv2.imdecode(payload, cv2.IMREAD_COLOR)
    if image is None:
        raise ValueError(f"Unable to decode image for table recognition: {path.name}")
    return image


def _preprocess(image: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
    height, width = image.shape[:2]
    if height <= 0 or width <= 0:
        raise ValueError("Table recognition image has invalid dimensions")

    ratio = _MODEL_SIZE / float(max(height, width))
    resize_h = max(1, int(height * ratio))
    resize_w = max(1, int(width * ratio))
    resized = cv2.resize(image, (resize_w, resize_h))
    normalized = (resized.astype(np.float32) / 255.0 - _MEAN) / _STD
    padded = np.zeros((_MODEL_SIZE, _MODEL_SIZE, 3), dtype=np.float32)
    padded[:resize_h, :resize_w, :] = normalized
    tensor = padded.transpose((2, 0, 1))[None, ...]
    shape = np.asarray([[height, width, ratio, ratio, _MODEL_SIZE, _MODEL_SIZE]], dtype=np.float32)
    return tensor, shape


def _prepare_characters(characters: list[str]) -> list[str]:
    values = list(characters)
    if "<td></td>" not in values:
        values.append("<td></td>")
    if "<td>" in values:
        values.remove("<td>")
    return ["sos", *values, "eos"]


def _decode(
    bbox_preds: np.ndarray,
    structure_probs: np.ndarray,
    shape_list: np.ndarray,
    characters: list[str],
) -> tuple[list[str], np.ndarray]:
    char_to_index = {char: index for index, char in enumerate(characters)}
    if "sos" not in char_to_index or "eos" not in char_to_index:
        raise ValueError("RapidTable character metadata is incomplete")
    ignored = {char_to_index["sos"], char_to_index["eos"]}
    end_idx = char_to_index["eos"]
    td_tokens = {"<td>", "<td", "<td></td>"}

    structure_idx = structure_probs.argmax(axis=2)
    structures: list[str] = []
    boxes_out: list[np.ndarray] = []
    for index, raw_index in enumerate(structure_idx[0]):
        char_index = int(raw_index)
        if index > 0 and char_index == end_idx:
            break
        if char_index in ignored:
            continue
        if char_index < 0 or char_index >= len(characters):
            continue
        token = characters[char_index]
        if token in td_tokens:
            bbox = np.asarray(bbox_preds[0, index], dtype=np.float32).copy()
            height, width = shape_list[0][:2]
            bbox[0::2] *= width
            bbox[1::2] *= height
            boxes_out.append(bbox)
        structures.append(token)

    wrapped = ["<html>", "<body>", "<table>", *structures, "</table>", "</body>", "</html>"]
    return wrapped, np.asarray(boxes_out, dtype=np.float32)


def _rescale_and_filter_bboxes(image: np.ndarray, cell_bboxes: np.ndarray) -> np.ndarray:
    if cell_bboxes.size == 0:
        return np.empty((0, 8), dtype=np.float32)
    result = cell_bboxes.copy()
    height, width = image.shape[:2]
    ratio = min(_MODEL_SIZE / height, _MODEL_SIZE / width)
    result[:, 0::2] *= _MODEL_SIZE / (width * ratio)
    result[:, 1::2] *= _MODEL_SIZE / (height * ratio)
    return result[~np.all(result == 0, axis=1)]


def _format_ocr(
    boxes: np.ndarray,
    texts: tuple[str, ...],
    scores: tuple[float, ...],
    image_shape: tuple[int, int],
) -> tuple[np.ndarray, list[tuple[str, float]]]:
    height, width = image_shape
    bboxes = np.asarray(boxes, dtype=np.float32)
    min_coords = bboxes[..., :2].min(axis=1)
    max_coords = bboxes[..., :2].max(axis=1)
    min_coords = np.maximum(min_coords, 0)
    max_coords = np.minimum(max_coords, [width, height])
    return np.hstack([min_coords, max_coords]), list(zip(texts, scores, strict=True))


def _filter_caption_ocr(
    cell_bboxes: np.ndarray,
    dt_boxes: np.ndarray,
    rec_res: list[tuple[str, float]],
) -> tuple[np.ndarray, list[tuple[str, float]]]:
    if cell_bboxes.size == 0:
        return np.empty((0, 4), dtype=np.float32), []
    top = float(cell_bboxes[:, 1::2].min())
    keep = [index for index, box in enumerate(dt_boxes) if float(np.max(box[1::2])) >= top]
    return dt_boxes[keep], [rec_res[index] for index in keep]


def _rect_from_cell(box: np.ndarray) -> tuple[float, float, float, float]:
    if len(box) == 8:
        return (
            float(np.min(box[0::2])),
            float(np.min(box[1::2])),
            float(np.max(box[0::2])),
            float(np.max(box[1::2])),
        )
    return tuple(float(value) for value in box[:4])  # type: ignore[return-value]


def _distance(first: np.ndarray | tuple[float, ...], second: tuple[float, float, float, float]) -> float:
    x1, y1, x2, y2 = (float(value) for value in first[:4])
    x3, y3, x4, y4 = second
    total = abs(x3 - x1) + abs(y3 - y1) + abs(x4 - x2) + abs(y4 - y2)
    return total + min(abs(x3 - x1) + abs(y3 - y1), abs(x4 - x2) + abs(y4 - y2))


def _iou(first: np.ndarray | tuple[float, ...], second: tuple[float, float, float, float]) -> float:
    x1, y1, x2, y2 = (float(value) for value in first[:4])
    x3, y3, x4, y4 = second
    area1 = max(0.0, x2 - x1) * max(0.0, y2 - y1)
    area2 = max(0.0, x4 - x3) * max(0.0, y4 - y3)
    left = max(x1, x3)
    right = min(x2, x4)
    top = max(y1, y3)
    bottom = min(y2, y4)
    if left >= right or top >= bottom:
        return 0.0
    intersection = (right - left) * (bottom - top)
    denominator = area1 + area2 - intersection
    return intersection / denominator if denominator > 0 else 0.0


def _match_result(dt_boxes: np.ndarray, cell_bboxes: np.ndarray) -> dict[int, list[int]]:
    matched: dict[int, list[int]] = {}
    for ocr_index, ocr_box in enumerate(dt_boxes):
        ranked: list[tuple[float, float, int]] = []
        for cell_index, cell_box in enumerate(cell_bboxes):
            rect = _rect_from_cell(cell_box)
            ranked.append((1.0 - _iou(ocr_box, rect), _distance(ocr_box, rect), cell_index))
        if not ranked:
            continue
        iou_distance, _l1, cell_index = min(ranked)
        if iou_distance >= 1.0 - 1e-8:
            continue
        matched.setdefault(cell_index, []).append(ocr_index)
    return matched


def _fill_structure(
    structures: list[str],
    matched: dict[int, list[int]],
    ocr_contents: list[tuple[str, float]],
) -> str:
    output: list[str] = []
    cell_index = 0
    for token in structures:
        if "</td>" not in token:
            output.append(token)
            continue

        empty_cell = token == "<td></td>"
        if empty_cell:
            output.append("<td>")

        indexes = matched.get(cell_index, [])
        for position, ocr_index in enumerate(indexes):
            content = ocr_contents[ocr_index][0].strip()
            if not content:
                continue
            if position != len(indexes) - 1:
                content += " "
            output.append(content)

        output.append("</td>" if empty_cell else token)
        cell_index += 1

    return "".join(token for token in output if token not in {"<thead>", "</thead>", "<tbody>", "</tbody>"})


__all__ = ["infer_table_html", "reset_table_engine_cache"]
