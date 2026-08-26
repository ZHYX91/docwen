"""Windows long-path and concurrency gate for the real EasyOFD fixture."""

from __future__ import annotations

import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from zipfile import ZipFile

import pytest
from scripts.release.verify_packaged_cli import _build_long_path_work_dir, _write_physical_page_ofd

from docwen_core.detection import detect_content_format

pytestmark = [
    pytest.mark.integration,
    pytest.mark.pr_gate,
    pytest.mark.release_gate,
    pytest.mark.windows_only,
    pytest.mark.skipif(sys.platform != "win32", reason="Windows MAX_PATH regression"),
]


def test_physical_page_ofd_fixture_survives_deep_output_path_concurrently(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    caller = tmp_path / "caller"
    caller.mkdir()
    monkeypatch.chdir(caller)
    output_root = _build_long_path_work_dir(tmp_path / "output")
    output_root.mkdir(parents=True)
    outputs = [output_root / f"machine-physical-pages-{index}.ofd" for index in range(2)]
    legacy_deepest = (
        output_root / ".machine-physical-pages-0-generation" / "test" / "Doc_0" / "Pages" / "Page_1" / "Content.xml"
    )
    assert len(str(outputs[0])) < 260
    assert len(str(legacy_deepest)) >= 260

    with ThreadPoolExecutor(max_workers=2) as executor:
        futures = [executor.submit(_write_physical_page_ofd, output) for output in outputs]
        for future in futures:
            future.result()

    assert Path.cwd() == caller
    assert not (caller / "test").exists()
    assert not (caller / "test.ofd").exists()
    assert not (output_root / "test").exists()
    assert not (output_root / "test.ofd").exists()
    assert not any(path.name.endswith("-generation") for path in output_root.iterdir())
    for output in outputs:
        detection = detect_content_format(output)
        assert detection.format == "ofd"
        assert detection.structure_status.value == "valid"
        with ZipFile(output) as archive:
            members = set(archive.namelist())
        assert {
            "OFD.xml",
            "Doc_0/Document.xml",
            "Doc_0/Pages/Page_0/Content.xml",
            "Doc_0/Pages/Page_1/Content.xml",
            "Doc_0/Res/Image_0.jpg",
            "Doc_0/Res/Image_1.jpg",
        } <= members
