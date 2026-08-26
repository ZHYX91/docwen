"""Route manifest for docwen_plugin_print.

Declared routes (explicit source formats so runtime RouteRegistry.find()
can match without category fallback):

    Document family → PDF:
        docx/doc/odt/rtf/wps/document → pdf   (Office bridge)
    Spreadsheet family → PDF:
        xlsx/xls/ods/et/csv/spreadsheet → pdf (Office bridge)
"""

from __future__ import annotations

from docwen_core.models.manifest import HonestyRoute, PluginManifest, RouteCapabilityRule, RouteSpec

PLUGIN_ID = "docwen_plugin_print"
PLUGIN_VERSION = "0.1.0"

_PRINT_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {},
    "required": [],
}

# ── Source format families ────────────────────────────────────────────

_DOCUMENT_SOURCES: tuple[str, ...] = ("docx", "doc", "odt", "rtf", "wps", "document")
_SPREADSHEET_SOURCES: tuple[str, ...] = ("xlsx", "xls", "ods", "et", "csv", "spreadsheet")


def _make_route(
    source: str,
    target: str,
    *,
    action_name: str = "",
) -> RouteSpec:
    """Create a RouteSpec with a standardised label."""
    label = f"{source.upper()} → {target.upper()}"
    if action_name:
        label += f" ({action_name})"
    return RouteSpec(
        source_format=source,
        target_format=target,
        action_name=action_name,
        label=label,
        options_schema=_PRINT_OPTIONS_SCHEMA,
    )


# ── Build all routes ──────────────────────────────────────────────────


def _build_all_routes() -> list[RouteSpec]:
    routes: list[RouteSpec] = []

    # Document family → PDF (Office bridge)
    for src in _DOCUMENT_SOURCES:
        routes.append(_make_route(src, "pdf"))

    # Spreadsheet family → PDF (Office bridge)
    for src in _SPREADSHEET_SOURCES:
        routes.append(_make_route(src, "pdf"))

    return routes


ALL_ROUTES: list[RouteSpec] = _build_all_routes()


def build_manifest() -> PluginManifest:
    return PluginManifest(
        plugin_id=PLUGIN_ID,
        name="DocWen Print Plugin",
        version=PLUGIN_VERSION,
        description=(
            "Generates fixed-layout output (PDF) from structured formats. "
            "Document→PDF and Spreadsheet→PDF are backed by the Office bridge. "
            "OFD export is an unavailable capability and is not executable."
        ),
        author="DocWen",
        routes=ALL_ROUTES,
        capability_rules=[
            RouteCapabilityRule(
                source_formats=_DOCUMENT_SOURCES,
                required_capabilities=("external_office.word",),
                limitations=("PDF fidelity depends on the selected external Word-compatible backend",),
            ),
            RouteCapabilityRule(
                source_formats=_SPREADSHEET_SOURCES,
                required_capabilities=("external_office.spreadsheet",),
                limitations=("PDF fidelity depends on the selected external spreadsheet backend",),
            ),
        ],
        extra={
            "unavailable_target_formats": ["ofd"],
            "third_party_deps": [],
            "unavailable_routes": [
                HonestyRoute(
                    source="document",
                    targets=["ofd"],
                    description="OFD export is unavailable in DocWen 0.9",
                ),
            ],
        },
    )
