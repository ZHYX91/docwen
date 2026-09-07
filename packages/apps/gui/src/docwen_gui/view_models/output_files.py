"""Project user-facing output files from the runtime artifact contract."""

from docwen_core.models.result import ConversionResult
from docwen_runtime.path_io import filesystem_path


def result_output_paths(result: ConversionResult, *, existing_only: bool = False) -> tuple[str, ...]:
    """Keep primary outputs first; exclude extracted images and internal manifests."""
    artifacts = [a for a in result.artifacts if a.kind not in {"manifest", "log"}]
    if existing_only:
        artifacts = [a for a in artifacts if _is_file(a.staging_path)]
    primary = next((a for a in artifacts if a.is_primary), None)
    image_output = primary is not None and primary.media_type.startswith("image/")
    documents = [
        a
        for a in artifacts
        if a.is_primary
        or a.kind == "primary"
        or a.media_type == "text/markdown"
        or (image_output and a.media_type.startswith("image/"))
    ]
    # A failed operation can retain a diagnostic report without a primary output.
    candidates = documents or [a for a in artifacts if a.kind in {"auxiliary", "intermediate"}]
    candidates.sort(key=lambda a: not a.is_primary)
    paths: list[str] = []
    for artifact in candidates:
        path = artifact.staging_path
        if not path or path in paths:
            continue
        paths.append(path)
    return tuple(paths)


def _is_file(path: str) -> bool:
    try:
        return bool(path) and filesystem_path(path).is_file()
    except (OSError, ValueError):
        return False
