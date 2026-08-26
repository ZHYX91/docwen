"""Fake ViewModel and Controller stubs for GUI unit tests."""

from __future__ import annotations

from tests.support.config import FakeConfigView


def optimization_capability_projection() -> dict[str, object]:
    """Canonical test catalog used by GUI ViewModel fakes."""

    route_specs = [
        *(
            ("markdown", "markdown", target, "conversion", None, ["template_name"])
            for target in ("docx", "doc", "odt", "rtf", "wps", "pdf", "xlsx", "xls", "ods", "csv")
        ),
        (
            "markdown",
            "markdown",
            "markdown",
            "action",
            "validate",
            [
                "enable_symbol_pairing",
                "enable_symbol_correction",
                "enable_typos_rule",
                "enable_sensitive_word",
            ],
        ),
        *(
            ("docx", "document", target, "conversion", None, ["locale", "yaml_key_labels"])
            for target in ("md", "doc", "odt", "rtf", "wps", "pdf")
        ),
        *(
            ("document", "document", target, "conversion", None, ["locale", "yaml_key_labels"])
            for target in ("md", "docx", "doc", "odt", "rtf", "wps", "pdf")
        ),
        ("docx", "document", "md", "action", "gongwen", ["locale"]),
        ("docx", "document", "docx", "action", "validate", ["enable_symbol_pairing"]),
        ("document", "document", "docx", "action", "validate", ["enable_symbol_pairing"]),
        *(
            ("xlsx", "spreadsheet", target, "conversion", None, [])
            for target in ("xls", "ods", "csv", "tsv", "et", "pdf")
        ),
        *(("xls", "spreadsheet", target, "conversion", None, []) for target in ("xlsx", "ods", "csv", "et", "pdf")),
        *(("ods", "spreadsheet", target, "conversion", None, []) for target in ("xlsx", "xls", "csv", "et", "pdf")),
        *(("et", "spreadsheet", target, "conversion", None, []) for target in ("xlsx", "xls", "ods", "csv", "pdf")),
        *(("csv", "spreadsheet", target, "conversion", None, []) for target in ("xlsx", "xls", "ods", "pdf")),
        ("tsv", "spreadsheet", "xlsx", "conversion", None, []),
        ("xlsx", "spreadsheet", "xlsx", "action", "merge_tables", ["merge_mode"]),
        *(
            ("image", "image", target, "conversion", None, ["compress_mode", "size_limit"])
            for target in ("md", "png", "jpg", "bmp", "gif", "tif", "webp", "pdf")
        ),
        ("image", "image", "md", "action", "invoice_cn", ["locale", "yaml_key_labels"]),
        ("image", "image", "tif", "action", "merge_images_to_tiff", ["tiff_mode"]),
        *(
            ("pdf", "layout", target, "conversion", None, ["render_dpi"])
            for target in ("md", "docx", "doc", "odt", "rtf", "png", "jpg", "tif")
        ),
        *(
            ("ofd", "layout", target, "conversion", None, ["render_dpi"])
            for target in ("pdf", "docx", "doc", "odt", "rtf", "png", "jpg", "tif")
        ),
        ("pdf", "layout", "md", "action", "invoice_cn", ["locale", "yaml_key_labels"]),
        ("pdf", "layout", "pdf", "action", "merge_pdfs", []),
        ("pdf", "layout", "pdf", "action", "split_pdf", ["pages"]),
    ]
    source_groups: dict[tuple[str, str], list[dict[str, object]]] = {}
    route_by_key: dict[tuple[str, str, str], dict[str, object]] = {}
    for source, category, target, operation, action, options in route_specs:
        route_id = f"{source}:{target}:{action or 'convert'}"
        route = {
            "id": route_id,
            "operation": operation,
            "action": action,
            "source": source,
            "target": target,
            "available": True,
            "state": "available",
            "options": options,
        }
        source_groups.setdefault((source, category), []).append(route)
        route_by_key[(source, action or "", target)] = route
    sources = [
        {"id": source, "category": category, "available": True, "routes": routes}
        for (source, category), routes in source_groups.items()
    ]

    def binding(scope: str, source: str, action: str, target: str) -> dict[str, object]:
        route = route_by_key[(source, action, target)]
        category = next(category for candidate, category in source_groups if candidate == source)
        return {
            "scope": scope,
            "route_id": route["id"],
            "source": source,
            "source_category": category,
            "target": target,
            "available": True,
            "state": "available",
        }

    resources = [
        {
            "id": "gongwen",
            "name": "Gongwen",
            "action_name": "gongwen",
            "scopes": ["document_to_md"],
            "available": True,
            "state": "available",
            "bindings": [binding("document_to_md", "docx", "gongwen", "md")],
        },
        {
            "id": "invoice_cn",
            "name": "Invoice CN",
            "action_name": "invoice_cn",
            "scopes": ["layout_to_md", "image_to_md"],
            "available": True,
            "state": "available",
            "bindings": [
                binding("layout_to_md", "pdf", "invoice_cn", "md"),
                binding("image_to_md", "image", "invoice_cn", "md"),
            ],
        },
    ]
    return {
        "resource": "formats",
        "contract": {"id": "docwen.runtime-capabilities", "version": 1},
        "runtime": {"state": "available", "platform": "test"},
        "security": {"dependency_egress_guard": {}},
        "gates": [],
        "sources": sources,
        "counts": {
            "sources": len(sources),
            "routes": len(route_specs),
            "available_routes": len(route_specs),
            "unavailable_routes": 0,
            "actions": sum(operation == "action" for _s, _c, _t, operation, _a, _o in route_specs),
        },
        "optimizations": {
            "resource": "optimizations",
            "contract": {"id": "docwen.optimizations", "version": 1},
            "runtime": {"state": "available", "platform": "test"},
            "resources": resources,
            "counts": {
                "resources": 2,
                "available_resources": 2,
                "unavailable_resources": 0,
                "bindings": 3,
                "available_bindings": 3,
                "unavailable_bindings": 0,
            },
        },
    }


class FakeController:
    """Minimal controller stub for ``MainWindowViewModel``."""

    def __init__(self, config_values: dict[str, object] | None = None) -> None:
        self.config_port: FakeConfigView = FakeConfigView(config_values or {})
        self.has_runtime = True

    def describe_runtime_capabilities(self) -> dict[str, object]:
        return optimization_capability_projection()

    def stop(self) -> None:
        pass


class FakeMainWindowViewModel:
    """Minimal ``MainWindowViewModel`` stub for ``ActionAreaViewModel``."""

    def __init__(self, config_values: dict[str, object] | None = None) -> None:
        self.controller: FakeController = FakeController(config_values)
