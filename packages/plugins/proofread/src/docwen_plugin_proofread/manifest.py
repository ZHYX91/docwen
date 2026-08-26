"""Proofread plugin manifest — route declarations and option schemas.

Route coverage:
    FEAT-CONV-009  ACT-VALIDATE-DOCX     docx→docx           (validate)
    FEAT-CONV-009  ACT-VALIDATE-DOCUMENT document→docx       (validate, application pre-conversion)
    FEAT-CONV-010  ACT-VALIDATE-MD       markdown→markdown   (validate)

The plugin parses DOCX and Markdown directly.  The category-level document
route exposes the Application-owned DOC/WPS/RTF/ODT → DOCX pre-conversion
chain without duplicating Office bridge code inside this plugin.
"""

from __future__ import annotations

from docwen_core.models.manifest import PluginManifest, RouteCapabilityRule, RouteSpec
from docwen_plugin_proofread.rules import PROOFREAD_OPTIONS_SCHEMA

PLUGIN_ID = "docwen_plugin_proofread"
PLUGIN_VERSION = "0.1.0"

# ── Route definitions ──────────────────────────────────────────────────

ROUTE_VALIDATE_DOCX = RouteSpec(
    source_format="docx",
    target_format="docx",
    action_name="validate",
    label="DOCX Proofread — validate text and annotate issues as comments",
    options_schema=PROOFREAD_OPTIONS_SCHEMA,
)

ROUTE_VALIDATE_DOCUMENT = RouteSpec(
    source_format="document",
    target_format="docx",
    action_name="validate",
    label="Legacy document proofread — pre-convert to DOCX and annotate issues as comments",
    options_schema=PROOFREAD_OPTIONS_SCHEMA,
)

ROUTE_VALIDATE_MD = RouteSpec(
    source_format="markdown",
    target_format="markdown",
    action_name="validate",
    label="Markdown Proofread — validate text and produce a JSON report",
    options_schema=PROOFREAD_OPTIONS_SCHEMA,
)

ALL_ROUTES: list[RouteSpec] = [
    ROUTE_VALIDATE_DOCX,
    ROUTE_VALIDATE_DOCUMENT,
    ROUTE_VALIDATE_MD,
]


def build_manifest() -> PluginManifest:
    """Build and return the Proofread plugin manifest."""
    return PluginManifest(
        plugin_id=PLUGIN_ID,
        name="DocWen Proofread Plugin",
        version=PLUGIN_VERSION,
        description=(
            "Rule-based text proofreading for DOCX and Markdown files. "
            "DOC/WPS/RTF/ODT inputs are pre-converted to DOCX by the "
            "Application layer before validation. "
            "Checks include symbol pairing (brackets/quotes), symbol correction "
            "(fullwidth digits), common Chinese typos, and configurable sensitive "
            "word detection."
        ),
        author="DocWen",
        routes=ALL_ROUTES,
        capability_rules=[
            RouteCapabilityRule(
                source_formats=("docx",),
                action_names=("validate",),
                required_capabilities=("python.docx",),
            ),
            RouteCapabilityRule(
                source_formats=("document",),
                action_names=("validate",),
                required_capabilities=("python.docx", "external_office.word"),
                limitations=("requires Application-owned pre-conversion to DOCX",),
            ),
        ],
        extra={
            "third_party_deps": ["python-docx"],
            "checks": [
                "symbol_pairing",
                "symbol_correction",
                "typos_rule",
                "sensitive_word",
            ],
        },
    )
