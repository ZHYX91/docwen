from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

import pytest
from scripts.release import build_v4_package_input as producer
from scripts.release import v4_package_input_machine as machine
from tests.test_repo import test_build_v4_package_input as base

pytestmark = pytest.mark.contract


def _quarantined(fixture: base.SyntheticFixture) -> list[Path]:
    return list(fixture.output.parent.glob(f".{fixture.output.name}.rejected-*"))


def _loaded_output(fixture: base.SyntheticFixture) -> tuple[producer.HarnessInput, producer.HarnessOutput]:
    harness = producer._load_harness(fixture.docwen)
    output = base._harness_runner()(
        fixture.checkpoint_path.parent / "DocWenCLI.exe",
        fixture.docwen,
        fixture.work / "transcript-run",
        harness,
    )
    return harness, output


@pytest.mark.parametrize(
    "mutate",
    [
        lambda value: value["events"].pop(0),
        lambda value: value["events"].__setitem__(0, value["events"][1]),
        lambda value: value["process"].__setitem__("sessionCount", 2),
        lambda value: next(event for event in value["events"] if str(event["operation"]).startswith("forward-plan:"))[
            "message"
        ]["params"]["inputs"][0].__setitem__("role", "source"),
        lambda value: next(event for event in value["events"] if str(event["operation"]).startswith("forward-plan:"))[
            "message"
        ]["params"].__setitem__("options", {"add_numbering": True}),
        lambda value: next(event for event in value["events"] if str(event["operation"]).startswith("reverse-plan:"))[
            "message"
        ]["params"]["inputs"][0].__setitem__("sha256", "0" * 64),
        lambda value: next(event for event in value["events"] if event["direction"] == "response").__setitem__(
            "rawFrameBase64", "AAAA"
        ),
        lambda value: next(event for event in value["events"] if event["direction"] == "request").__setitem__(
            "rawFrameBase64", "AAAA"
        ),
    ],
)
def test_sealed_single_session_transcript_mutations_fail_closed(tmp_path: Path, mutate: Any) -> None:
    fixture = base._fixture(tmp_path)
    harness, output = _loaded_output(fixture)
    value = json.loads(output.transcript)
    mutate(value)
    with pytest.raises(producer.V4PackageInputError):
        producer._validate_session_transcript(
            base._json_bytes(value),
            harness=harness,
            outputs=output.cases,
        )


def test_success_seals_exact_two_transcript_in_packaged_record(tmp_path: Path) -> None:
    fixture = base._fixture(tmp_path)
    producer.build_package_input(**base._arguments(fixture))  # type: ignore[arg-type]
    plan = json.loads((fixture.output / "evidence-plan.json").read_text(encoding="utf-8"))
    packaged = next(item for item in plan["records"] if item["layer"] == "packaged")
    envelope = json.loads((fixture.output / "evidence-input" / packaged["artifact"]).read_text(encoding="utf-8"))
    pointer = envelope["observation"]["payload"]["invocation"]["stdout"]
    relative = Path(pointer["relativePath"]).relative_to("evidence")
    raw = (fixture.output / "evidence-input" / relative).read_bytes()
    assert len(raw) == pointer["bytes"]
    assert hashlib.sha256(raw).hexdigest() == pointer["sha256"]
    assert json.loads(raw)["schema"] == machine.TRANSCRIPT_SCHEMA


def test_loaded_harness_pointer_drift_is_rejected(tmp_path: Path) -> None:
    fixture = base._fixture(tmp_path)
    harness = producer._load_harness(fixture.docwen)
    harness.cases[0].plan_path.write_bytes(b"{}")
    with pytest.raises(producer.V4PackageInputError, match="harness_pointer_changed"):
        producer._revalidate_harness(harness, docwen=fixture.docwen)


def test_binary_drift_during_version_probe_is_quarantined(tmp_path: Path) -> None:
    fixture = base._fixture(tmp_path)

    def drift(executable: Path) -> str:
        executable.write_bytes(b"mutated-after-version-read")
        return "DocWen 0.9.0 (CLI protocol 3)"

    with pytest.raises(producer.V4PackageInputError, match="changed_during_version_probe"):
        producer.build_package_input(
            **base._arguments(fixture, version_reader=drift)  # type: ignore[arg-type]
        )
    assert not fixture.output.exists()
    assert len(_quarantined(fixture)) == 1


def test_package_and_harness_drift_after_machine_run_are_quarantined(tmp_path: Path) -> None:
    package_fixture = base._fixture(tmp_path / "package")
    complete = base._harness_runner()

    def package_drift(
        executable: Path,
        docwen_clone: Path,
        run_root: Path,
        harness: producer.HarnessInput,
    ) -> producer.HarnessOutput:
        output = complete(executable, docwen_clone, run_root, harness)
        executable.write_bytes(b"drifted-cli")
        return output

    with pytest.raises(producer.V4PackageInputError, match="package_changed_during_harness"):
        producer.build_package_input(
            **base._arguments(package_fixture, harness_runner=package_drift)  # type: ignore[arg-type]
        )
    assert len(_quarantined(package_fixture)) == 1

    harness_fixture = base._fixture(tmp_path / "harness")

    def harness_drift(
        executable: Path,
        docwen_clone: Path,
        run_root: Path,
        harness: producer.HarnessInput,
    ) -> producer.HarnessOutput:
        output = complete(executable, docwen_clone, run_root, harness)
        harness.cases[0].plan_path.write_bytes(b"{}")
        return output

    with pytest.raises(producer.V4PackageInputError, match="harness_pointer_changed"):
        producer.build_package_input(
            **base._arguments(harness_fixture, harness_runner=harness_drift)  # type: ignore[arg-type]
        )
    assert len(_quarantined(harness_fixture)) == 1


@pytest.mark.parametrize(("drift_call", "error"), [(2, "before_publish"), (3, "after_publish")])
def test_source_checkpoint_is_revalidated_around_atomic_publish(
    tmp_path: Path,
    drift_call: int,
    error: str,
) -> None:
    fixture = base._fixture(tmp_path)
    stable = base._checkpoint_loader(fixture)
    calls = 0

    def loader(path: Path, digest: str, docwen: Path) -> tuple[dict[str, Any], dict[str, object]]:
        nonlocal calls
        calls += 1
        value, identity = stable(path, digest, docwen)
        if calls == drift_call:
            value = {**value, "mutation": True}
        return value, identity

    with pytest.raises(producer.V4PackageInputError, match=f"source_checkpoint_changed_{error}"):
        producer.build_package_input(
            **base._arguments(fixture, checkpoint_loader=loader)  # type: ignore[arg-type]
        )
    assert calls == drift_call
    assert not fixture.output.exists()
    assert len(_quarantined(fixture)) == 1


def test_clone_identity_is_checked_at_every_phase(tmp_path: Path) -> None:
    fixture = base._fixture(tmp_path)
    labels: list[str] = []

    def verifier(_repo: Path, _commit: str, _tree: str, label: str) -> None:
        labels.append(label)

    producer.build_package_input(
        **base._arguments(fixture, clone_verifier=verifier)  # type: ignore[arg-type]
    )
    phases = ["post_clone", "post_build", "post_harness", "post_evidence", "pre_publish", "post_publish"]
    assert labels == [f"docwen_{phase}" for phase in phases]
