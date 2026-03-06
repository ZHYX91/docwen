from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from docwen.errors import InvalidInputError


@dataclass(slots=True)
class ConversionRequest:
    file_path: str | None = None
    action_type: str | None = None
    source_format: str | None = None
    target_format: str | None = None
    file_list: list[str] | None = None
    options: dict[str, Any] = field(default_factory=dict)
    actual_format: str | None = None
    category: str | None = None

    def validate(self) -> None:
        if not self.file_path and self.file_list:
            self.file_path = self.file_list[0]

        if self.action_type is not None:
            action = self.action_type.strip()
            if not action:
                raise InvalidInputError("缺少 action_type")
            self.action_type = action

        if self.source_format is not None:
            value = self.source_format.strip()
            self.source_format = value or None

        if self.target_format is not None:
            value = self.target_format.strip()
            self.target_format = value.lower() if value else None

        if not self.file_path:
            raise InvalidInputError("缺少输入文件路径")
        if not Path(self.file_path).exists():
            raise InvalidInputError("文件不存在", details=self.file_path)

        is_named_action = self.action_type is not None
        if is_named_action and (self.source_format or self.target_format):
            raise InvalidInputError("命名动作不应同时指定 source/target")

        aggregate_actions = {"merge_pdfs", "merge_images_to_tiff", "merge_tables"}
        if self.action_type in aggregate_actions:
            if not self.file_list or len(self.file_list) < 2:
                raise InvalidInputError("聚合操作需要至少两个输入文件")
            if self.file_path not in self.file_list:
                raise InvalidInputError("聚合操作需要选中的基准文件包含在 file_list 中")
            if self.action_type == "merge_tables":
                base_table = self.options.get("base_table")
                if base_table and str(base_table) not in self.file_list:
                    raise InvalidInputError("汇总表格需要基准表格包含在 file_list 中")

        if self.action_type is None:
            target = self.target_format or self.options.get("target_format")
            target = str(target).strip().lower() if target else None
            if not target:
                raise InvalidInputError("缺少目标格式", details="target_format")
            self.target_format = target


@dataclass(slots=True)
class BatchRequest:
    requests: list[ConversionRequest]
    continue_on_error: bool = True
    max_workers: int = 4
