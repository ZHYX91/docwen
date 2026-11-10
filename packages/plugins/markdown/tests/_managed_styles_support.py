"""Imported numeric-style fixtures independent of shipped template identities."""

from docx.oxml.ns import qn

from docwen_core.docx_styles import CUSTOM_DOCUMENT_STYLES

_REF_TAGS = ("pStyle", "rStyle", "tblStyle", "basedOn", "next", "link", "numStyleLink", "styleLink")
_REF_QNAMES = tuple(qn(f"w:{name}") for name in _REF_TAGS)


def _use_numeric_custom_style_ids(document) -> None:
    """Exercise imported numeric IDs independently of current distribution IDs."""
    ids = {definition.style_id: str(9000 + index) for index, definition in enumerate(CUSTOM_DOCUMENT_STYLES)}
    for style in document.styles:
        if style.style_id in ids:
            style._element.set(qn("w:styleId"), ids[style.style_id])
    for root in (document.styles.element, document.element):
        for tag in _REF_QNAMES:
            for reference in root.iter(tag):
                value = reference.get(qn("w:val"))
                if value in ids:
                    reference.set(qn("w:val"), ids[value])
