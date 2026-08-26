"""Round-trip orchestration helper for numbering integration tests.

Drives the full conversion chain through the production composition
root so that cross-plugin round-trips stay within package boundaries:

    Markdown --[docwen_plugin_markdown]--> DOCX --[docwen_plugin_document]--> Markdown

The helper never imports either plugin's internals directly. It
obtains a real ``RuntimePort`` from ``docwen_bundle.runtime_factory``
(the same factory CLI/GUI use), builds ``ConversionRequest`` objects
from the shared ``docwen_core`` models, and reads the final artifact
back from disk.

Tests consume the small public surface below:

- ``md_to_docx(...)`` — run the forward leg, return the produced
  ``.docx`` path.
- ``docx_to_md(...)`` — run the reverse leg, return the produced
  Markdown text.
- ``round_trip_md(...)`` — run both legs, return ``(docx_path, md_text)``.

The session-scoped ``round_trip_runtime`` fixture that these functions
take as their first argument is defined in ``conftest.py`` (pytest only
collects fixtures from conftest modules).

Numbering options are passed straight through to the forward
(markdown -> docx) leg as the converter reads them verbatim from the
request options dict (``remove_numbering`` / ``add_numbering`` /
``numbering_scheme`` / ``heading_numbering_render_mode``).
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy

# ── Low-level leg runners ─────────────────────────────────────────────────


def _run(
    runtime: Any,
    request_id: str,
    input_path: Path,
    *,
    source_format: str,
    target_format: str,
    output_dir: Path,
    action_name: str = "",
    options: dict[str, Any] | None = None,
) -> Any:
    """Execute one conversion leg and return the ConversionResult.

    Goes through the adapter's synchronous ``execute`` — the public
    ``RuntimePort`` surface — rather than reaching into task-manager
    internals. This is the same path the application layer uses.
    """
    size = input_path.stat().st_size if input_path.exists() else 0
    request = ConversionRequest(
        request_id=request_id,
        input_refs=[
            FileRef(
                path=str(input_path),
                format=source_format,
                category="document",
                size_bytes=size,
            )
        ],
        target_format=target_format,
        action_name=action_name,
        output_policy=OutputPolicy(output_dir=str(output_dir)),
        options=options or {},
    )
    return runtime.execute(request)


def _primary_path(result: Any) -> Path:
    """Return the on-disk path of the primary artifact from a result."""
    assert result.success, f"Conversion failed: {result.error.message if result.error else 'unknown'}"
    primaries = [a for a in result.artifacts if a.kind == "primary"]
    assert primaries, f"No primary artifact in result: {result.artifacts}"
    return Path(primaries[0].staging_path)


# ── Public round-trip API ─────────────────────────────────────────────────


def md_to_docx(
    runtime: Any,
    md_path: Path,
    output_dir: Path,
    *,
    request_id: str = "rt-md-to-docx",
    options: dict[str, Any] | None = None,
) -> Path:
    """Forward leg: Markdown -> DOCX. Returns the produced ``.docx`` path."""
    result = _run(
        runtime,
        request_id,
        md_path,
        source_format="markdown",
        target_format="docx",
        output_dir=output_dir,
        options=options,
    )
    return _primary_path(result)


def docx_to_md(
    runtime: Any,
    docx_path: Path,
    output_dir: Path,
    *,
    request_id: str = "rt-docx-to-md",
    options: dict[str, Any] | None = None,
    preserve_numbering: bool = True,
) -> str:
    """Reverse leg: DOCX -> Markdown. Returns the produced Markdown text.

    The DOCX->MD converter defaults ``remove_numbering`` to ``True``
    internally, which strips every heading prefix matched by the shared
    clean rules. For a *round-trip* that masks the very thing under
    test — whether forward-leg numbering survives — so this helper
    flips the default: ``preserve_numbering=True`` (the default) sends
    ``remove_numbering=False`` so prefixes are kept, unless the caller
    passes an explicit ``remove_numbering`` in *options* or sets
    ``preserve_numbering=False``.
    """
    opts = dict(options or {})
    if "remove_numbering" not in opts:
        opts["remove_numbering"] = not preserve_numbering
    result = _run(
        runtime,
        request_id,
        docx_path,
        source_format="docx",
        target_format="md",
        output_dir=output_dir,
        options=opts,
    )
    path = _primary_path(result)
    return path.read_text(encoding="utf-8")


def round_trip_md(
    runtime: Any,
    md_path: Path,
    work_dir: Path,
    *,
    forward_options: dict[str, Any] | None = None,
    reverse_options: dict[str, Any] | None = None,
    preserve_numbering: bool = True,
) -> tuple[Path, str]:
    """Run Markdown -> DOCX -> Markdown. Returns ``(docx_path, md_text)``.

    Both legs write into ``work_dir`` (created if missing). The forward
    leg's options carry the numbering knobs. The reverse leg preserves
    heading numbering by default (``preserve_numbering=True``) so the
    round-trip reports what actually survived; pass
    ``preserve_numbering=False`` (or an explicit ``remove_numbering``
    in *reverse_options*) to observe the default-stripping behaviour.
    """
    work_dir.mkdir(parents=True, exist_ok=True)
    docx_path = md_to_docx(
        runtime,
        md_path,
        work_dir,
        request_id="rt-forward",
        options=forward_options,
    )
    md_text = docx_to_md(
        runtime,
        docx_path,
        work_dir,
        request_id="rt-reverse",
        options=reverse_options,
        preserve_numbering=preserve_numbering,
    )
    return docx_path, md_text
