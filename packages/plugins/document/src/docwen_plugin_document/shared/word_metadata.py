"""Offline validation of the Word metadata used in native note proofs."""

from functools import cache
from importlib.resources import files
from typing import Any

from lxml import etree

_Q = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_UNSIGNED_ATTRIBUTES: dict[str, tuple[str, ...]] = {
    _Q + name: (_Q + "val",) for name in ("sz", "szCs", "kern", "fitText", "tabIndex")
}
_UNSIGNED_ATTRIBUTES[_Q + "bdr"] = (_Q + "sz", _Q + "space")


@cache
def _metadata_schema() -> etree.XMLSchema:
    root = files("docwen_plugin_document").joinpath("resources", "word_metadata")
    schemas = {name: root.joinpath(name).read_bytes() for name in ("wml.xsd", "shared-commonSimpleTypes.xsd")}

    class LocalSchemaResolver(etree.Resolver):
        def resolve(self, url: str, public_id: Any, context: Any) -> Any:
            # Only the pinned local import is available; no document or network paths.
            if url not in schemas:
                raise ValueError(f"Unknown note metadata schema import: {url}")
            return self.resolve_string(schemas[url], context, base_url=url)

    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    parser.resolvers.add(LocalSchemaResolver())
    return etree.XMLSchema(etree.fromstring(schemas["wml.xsd"], parser))


def is_valid_word_metadata(element: Any) -> bool:
    if element.tail and element.tail.strip(" \t\r\n"):
        return False
    for node in element.iter():
        if node.tag == _Q + "rPr":
            tags = [child.tag for child in node]
            if len(tags) != len(set(tags)):
                return False
        # Some libxml2 versions accept signed lexical forms for unsignedLong.
        # Keep the same explicit no-sign contract as Core on every platform.
        for attribute in _UNSIGNED_ATTRIBUTES.get(node.tag, ()):
            value = node.get(attribute)
            if value is not None and value.lstrip(" \t\r\n").startswith(("+", "-")):
                return False
    return _metadata_schema().validate(element)


def run_metadata_is_valid(run: Any) -> bool:
    properties = run.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
    return len(properties) <= 1 and (
        not properties or (run[0] is properties[0] and is_valid_word_metadata(properties[0]))
    )
