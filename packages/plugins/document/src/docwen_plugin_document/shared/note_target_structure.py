"""Prove note target payloads without treating formatting metadata as body text."""

from typing import Any

from docwen_core.docx_parsing.xml_ns import NS_W
from docwen_plugin_document.shared.markdown_runs import _is_transparent_proof_marker
from docwen_plugin_document.shared.word_metadata import is_valid_word_metadata

_Q = f"{{{NS_W}}}"
NOTE_WRAPPERS = frozenset(
    _Q + name for name in ("ins", "moveTo", "smartTag", "sdt", "sdtContent", "customXml", "hyperlink", "fldSimple")
)
_METADATA_PARENTS = {
    _Q + "rPr": _Q + "r",
    _Q + "sdtPr": _Q + "sdt",
    _Q + "sdtEndPr": _Q + "sdt",
    _Q + "smartTagPr": _Q + "smartTag",
    _Q + "customXmlPr": _Q + "customXml",
}


def _inside_metadata(element: Any) -> bool:
    return any(ancestor.tag in _METADATA_PARENTS for ancestor in element.iterancestors())


def _valid_wrapper_shell(element: Any) -> bool:
    if any(value and value.strip(" \t\r\n") for value in (element.text, element.tail)):
        return False
    parent = element.getparent()
    if parent is None:
        return False
    if element.tag == _Q + "sdtContent":
        return parent.tag == _Q + "sdt" and not element.attrib
    if parent.tag not in (NOTE_WRAPPERS - {_Q + "sdt"}) | {_Q + "p"}:
        return False
    if element.tag == _Q + "sdt":
        tags = [child.tag for child in element]
        allowed = [_Q + "sdtPr", _Q + "sdtEndPr", _Q + "sdtContent"]
        return (
            not element.attrib
            and tags.count(_Q + "sdtContent") == 1
            and all(tags.count(tag) <= 1 for tag in allowed)
            and tags == [tag for tag in allowed if tag in tags]
        )
    metadata_name = {_Q + "r": "rPr", _Q + "smartTag": "smartTagPr", _Q + "customXml": "customXmlPr"}.get(element.tag)
    if metadata_name is not None:
        metadata = element.findall(_Q + metadata_name)
        return len(metadata) <= 1 and (not metadata or element[0] is metadata[0])
    return True


def is_pure_note_target(contents: list[Any], reference: Any) -> bool:
    """Accept rendered wrappers and independently validate their non-body metadata."""
    shells = set(contents)
    for ancestor in reference.iterancestors():
        if ancestor.tag == _Q + "p":
            break
        shells.add(ancestor)
    wrappers = [node for node in shells if node.tag in NOTE_WRAPPERS | {_Q + "r"} and not _inside_metadata(node)]
    if any(not _valid_wrapper_shell(node) for node in wrappers):
        return False
    metadata = {node for node in contents if node.tag in _METADATA_PARENTS and not _inside_metadata(node)}
    for wrapper in wrappers:
        metadata.update(child for child in wrapper if child.tag in _METADATA_PARENTS)
    metadata_nodes: set[Any] = set()
    for node in metadata:
        parent = node.getparent()
        if parent is None or parent.tag != _METADATA_PARENTS[node.tag] or not is_valid_word_metadata(node):
            return False
        metadata_nodes.update(node.iter())
    containers = NOTE_WRAPPERS | {_Q + "r", _Q + "bookmarkStart", _Q + "bookmarkEnd"}
    return all(
        node is reference or node in metadata_nodes or node.tag in containers or _is_transparent_proof_marker(node)
        for node in contents
    )
