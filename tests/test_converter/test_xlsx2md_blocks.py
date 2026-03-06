"""converter 单元测试。"""

from __future__ import annotations

import pandas as pd
import pytest

from docwen.converter.xlsx2md import core as xlsx2md_core

pytestmark = pytest.mark.unit


def test_find_blocks_does_not_merge_diagonal_cells() -> None:
    df = pd.DataFrame(
        [
            ["X", ""],
            ["", "Y"],
        ]
    )

    blocks = xlsx2md_core._find_blocks(df)

    assert len(blocks) == 2


def test_find_blocks_treats_whitespace_only_cells_as_empty() -> None:
    df = pd.DataFrame(
        [
            [" ", "A"],
        ]
    )

    blocks = xlsx2md_core._find_blocks(df)

    assert len(blocks) == 1
    assert blocks[0].iat[0, 0] == "A"
