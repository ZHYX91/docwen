"""Typed data models for gongwen recognition and rendering."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True, slots=True)
class AttachmentDocument:
    """One structured Markdown attachment derived from an official document."""

    ordinal: int
    title: str
    markdown: str
    paragraph_indices: tuple[int, ...]
    children: tuple[AttachmentDocument, ...] = ()


@dataclass
class ParagraphFeature:
    """Extracted features of a single DOCX paragraph."""

    index: int
    text: str
    font_name: str = ""
    font_size_pt: float | None = None
    style_name: str = ""
    outline_level: int | None = None
    alignment: str = ""  # LEFT, CENTER, RIGHT, JUSTIFY
    is_first_in_section: bool = False
    has_image: bool = False
    is_in_textbox: bool = False
    table_cell_context: str = ""  # "header" | "body" | ""
    # Rich content fields (Task 2)
    has_formula: bool = False
    formula_type: str = ""  # "inline" | "block" | ""
    formula_latex: str = ""  # extracted LaTeX, if any
    has_page_break: bool = False
    has_section_break: bool = False
    heading_level: int = 0  # 1-5 if heading detected, 0 otherwise
    heading_numbering_text: str = ""  # the numbering prefix (e.g. "一、")
    heading_body_boundary: int | None = None  # character boundary in cleaned text
    heading_body_boundary_source: str = ""  # "run_format" | "punctuation_fallback" | ""
    extracted_images: list[str] = field(default_factory=list)  # staging paths
    image_ocr_texts: dict[str, str] = field(default_factory=dict)  # path → OCR text
    raw_text: str = ""  # original text before heading cleaning
    source: str = "body"  # body, textbox, table, header, footer
    # Position of the owning top-level element in ``document.xml``.  Unlike
    # ``index`` this may be sparse because empty paragraphs are not emitted,
    # and several injected table/textbox features may share one source index.
    source_index: int | None = None
    table_index: int | None = None
    table_row_index: int | None = None
    table_cell_index: int | None = None
    table_markdown: str = ""
    is_table_anchor: bool = False
    table_fidelity_risks: tuple[str, ...] = ()
    table_output_mode: str = ""  # rendered | structural_metadata | unrepresentable


@dataclass
class ScoringRule:
    """A single scoring rule: condition method name + score."""

    condition: str  # method name on ElementScorer
    score: int  # points added if condition passes


@dataclass(order=True)
class RecognitionCandidate:
    """A paragraph→element identification with score."""

    element_type: str = field(compare=False)
    score: int
    para_index: int = field(compare=False)
    trace: list[str] = field(default_factory=list, compare=False)
    confidence: str = ""  # high, medium, low


@dataclass
class RecognitionResult:
    """Full result of the recognition stage."""

    candidates: dict[int, RecognitionCandidate]  # para_index → best candidate
    yaml_info: dict[str, str | list[str]]
    skip_indices: list[int]  # structural paragraphs to skip in body
    review_signals: list[str]
    missing_required: list[str]
    validation_finding_count: int


@dataclass
class GongwenMetadata:
    """The 18 gongwen YAML fields as typed attributes."""

    aliases: list[str] = field(default_factory=list)
    title: str = ""
    subtitle: str = ""
    copy_id: str = ""
    security_classification: str = ""
    urgency: str = ""
    doc_number: str = ""
    issuing_authority_mark: str = ""
    signer: list[str] = field(default_factory=list)
    issuing_authority_signature: str = ""
    issue_date: str = ""
    printing_date: str = ""
    recipient: str = ""
    notes: str = ""
    printing_authority: str = ""
    copy_to: list[str] = field(default_factory=list)
    attachment: list[str] = field(default_factory=list)
    disclosure: str = ""

    @classmethod
    def default(cls) -> GongwenMetadata:
        return cls()

    def to_dict(self) -> dict[str, str | list[str]]:
        return {
            "aliases": self.aliases,
            "标题": self.title,
            "副标题": self.subtitle,
            "份号": self.copy_id,
            "密级和保密期限": self.security_classification,
            "紧急程度": self.urgency,
            "发文字号": self.doc_number,
            "发文机关标志": self.issuing_authority_mark,
            "签发人": self.signer,
            "发文机关署名": self.issuing_authority_signature,
            "成文日期": self.issue_date,
            "印发日期": self.printing_date,
            "主送机关": self.recipient,
            "附注": self.notes,
            "印发机关": self.printing_authority,
            "抄送机关": self.copy_to,
            "附件说明": self.attachment,
            "公开方式": self.disclosure,
        }

    @classmethod
    def from_dict(cls, d: dict) -> GongwenMetadata:
        """Create GongwenMetadata from a dict (reverse of to_dict)."""
        return cls(
            aliases=d.get("aliases", []),
            title=d.get("标题", ""),
            subtitle=d.get("副标题", ""),
            copy_id=d.get("份号", ""),
            security_classification=d.get("密级和保密期限", ""),
            urgency=d.get("紧急程度", ""),
            doc_number=d.get("发文字号", ""),
            issuing_authority_mark=d.get("发文机关标志", ""),
            signer=d.get("签发人", []),
            issuing_authority_signature=d.get("发文机关署名", ""),
            issue_date=d.get("成文日期", ""),
            printing_date=d.get("印发日期", ""),
            recipient=d.get("主送机关", ""),
            notes=d.get("附注", ""),
            printing_authority=d.get("印发机关", ""),
            copy_to=d.get("抄送机关", []),
            attachment=d.get("附件说明", []),
            disclosure=d.get("公开方式", ""),
        )
