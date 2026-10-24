"""Human-readable projection of one normalized, route-scoped request."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from docwen_core.models.request import POSTPROCESS_PROOFREAD_OPTION
from docwen_gui.i18n import t

if TYPE_CHECKING:
    from docwen_core.models.request import ConversionRequest


def request_summary(request: ConversionRequest, *, template_label: str = "") -> str:
    """Show only options actually present; never infer another set of defaults."""
    options = request.options
    parts = [request.target_format.upper()]
    if template := options.get("template_name"):
        parts.append(t("main_window.request_template", value=template_label or template))
    enabled, disabled = t("settings.changes.enabled"), t("settings.changes.disabled")
    if "remove_numbering" in options:
        parts.append(
            t("main_window.request_remove_numbering", value=enabled if options["remove_numbering"] else disabled)
        )
    if "add_numbering" in options:
        parts.append(t("main_window.request_add_numbering", value=enabled if options["add_numbering"] else disabled))
    if options.get("add_numbering") and (scheme := options.get("numbering_scheme")):
        parts.append(str(scheme))
        mode = options.get("heading_numbering_render_mode")
        if mode == "word_native":
            parts.append(t("settings.text.render_mode_word_native"))
        elif mode == "text":
            parts.append(t("settings.text.render_mode_text"))
    checks = options.get(POSTPROCESS_PROOFREAD_OPTION, options if request.action_name == "validate" else {})
    if isinstance(checks, dict):
        labels_by_key = {
            "enable_symbol_pairing": t("conversion_panel.document.symbol_pairing"),
            "enable_symbol_correction": t("conversion_panel.document.symbol_correction"),
            "enable_typos_rule": t("conversion_panel.document.typos_rule"),
            "enable_sensitive_word": t("conversion_panel.document.sensitive_word"),
        }
        labels = [label for key, label in labels_by_key.items() if checks.get(key)]
        if labels:
            parts.append(t("main_window.request_proofread", value=", ".join(labels)))
    policy = request.output_policy
    destination = policy.output_path or policy.output_dir
    if not destination:
        directories = {
            policy.for_input(ref.path).output_dir or str(Path(ref.path).parent)
            for ref in request.input_refs
            if ref.input_role == "source"
        }
        destination = next(iter(directories)) if len(directories) == 1 else t("info_area.output_each_source")
    parts.append(t("main_window.request_output", value=destination))
    if policy.per_input_output_dirs:
        parts.extend(sorted(set(policy.per_input_output_dirs.values())))
    if policy.date_subfolder:
        parts.append(t("info_area.output_date_subfolder"))
    return t("main_window.request_summary", value=" · ".join(parts))
