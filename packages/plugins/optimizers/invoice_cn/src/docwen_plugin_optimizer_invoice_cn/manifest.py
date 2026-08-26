"""Invoice CN optimizer plugin manifest — route declarations and option schemas."""

from __future__ import annotations

from docwen_core.models.manifest import OptimizationResourceSpec, PluginManifest, RouteCapabilityRule, RouteSpec

PLUGIN_ID = "docwen_plugin_optimizer_invoice_cn"
PLUGIN_VERSION = "0.1.0"

# ── Option schemas (JSON Schema dicts) ────────────────────────────────

INVOICE_CN_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "to_md_enable_ocr": {
            "type": "boolean",
            "default": False,
            "description": "Enable OCR for image-based invoice conversion (automatically applied for image inputs)",
        },
        "ocr_language": {
            "type": "string",
            "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
            "default": "auto",
            "description": "OCR language model for scanned/image invoice inputs",
        },
        "locale": {
            "type": "string",
            "description": "App-resolved locale used for the invoice title label and OCR context.",
            "x-docwen-status": "implemented",
        },
        "yaml_key_labels": {
            "type": "object",
            "additionalProperties": {"type": "string"},
            "default": {},
            "description": "App-resolved YAML label override for the localized invoice title.",
            "x-docwen-status": "implemented",
        },
    },
    "required": [],
}

# ── Invoice routes (action-based) ─────────────────────────────────────

ROUTE_PDF_INVOICE_CN = RouteSpec(
    source_format="pdf",
    target_format="md",
    action_name="invoice_cn",
    label="PDF invoice → Structured Markdown",
    options_schema=INVOICE_CN_OPTIONS_SCHEMA,
)

ROUTE_OFD_INVOICE_CN = RouteSpec(
    source_format="ofd",
    target_format="md",
    action_name="invoice_cn",
    label="OFD invoice → Structured Markdown",
    options_schema=INVOICE_CN_OPTIONS_SCHEMA,
)

ROUTE_IMAGE_INVOICE_CN = RouteSpec(
    source_format="image",
    target_format="md",
    action_name="invoice_cn",
    label="Image invoice (OCR) → Structured Markdown",
    options_schema=INVOICE_CN_OPTIONS_SCHEMA,
)

# ── All routes ────────────────────────────────────────────────────────

ALL_ROUTES: list[RouteSpec] = [
    ROUTE_PDF_INVOICE_CN,
    ROUTE_OFD_INVOICE_CN,
    ROUTE_IMAGE_INVOICE_CN,
]


def build_manifest() -> PluginManifest:
    """Build and return the Invoice plugin manifest."""
    return PluginManifest(
        plugin_id=PLUGIN_ID,
        name="DocWen Invoice Plugin",
        version=PLUGIN_VERSION,
        description=(
            "Converts Chinese invoices (PDF/OFD/image) to structured Markdown "
            "with YAML frontmatter (20 metadata fields) and a detail-line table. "
            "Image invoices are processed via OCR."
        ),
        author="DocWen",
        routes=ALL_ROUTES,
        capability_rules=[
            RouteCapabilityRule(
                source_formats=("pdf",),
                required_capabilities=("python.fitz",),
                optional_capabilities=("python.rapidocr",),
                limitations=("scanned PDF pages require the optional RapidOCR capability",),
            ),
            RouteCapabilityRule(
                source_formats=("ofd",),
                required_capabilities=("python.easyofd",),
                optional_capabilities=("python.rapidocr",),
                limitations=("scanned OFD pages require the optional RapidOCR capability",),
            ),
            RouteCapabilityRule(
                source_formats=("image",),
                required_capabilities=("python.pillow", "python.rapidocr"),
            ),
        ],
        optimization_resources=[
            OptimizationResourceSpec(
                id="invoice_cn",
                name="Chinese invoice optimization",
                action_name="invoice_cn",
            )
        ],
        extra={
            "third_party_deps": ["PyMuPDF", "rapidocr_onnxruntime"],
            "invoice_cn_yaml_schema": [
                "发票种类",
                "发票代码",
                "发票号码",
                "开票日期",
                "校验码",
                "购买方名称",
                "购买方纳税人识别号",
                "购买方地址电话",
                "购买方开户行及账号",
                "销售方名称",
                "销售方纳税人识别号",
                "销售方地址电话",
                "销售方开户行及账号",
                "金额",
                "税额",
                "价税合计",
                "备注",
                "收款人",
                "复核",
                "开票人",
            ],
        },
    )
