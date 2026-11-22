"""Read YAML properties, including complete image values stored by note editors."""

from __future__ import annotations

import re
from typing import Any

import yaml
from yaml.nodes import MappingNode
from yaml.tokens import ScalarToken

_IMAGE_VALUE = re.compile(
    r"(!\[\[[^\]\r\n]+\]\]|!\[(?!\[)[^\]\r\n]*\](?:\([^\r\n]*\))?)"
    r"(?=(?:[ \t]+\#[^\r\n]*)?[ \t]*(?:[\r\n]|\Z))"
)


class _PropertiesLoader(yaml.SafeLoader):
    def __init__(self, payload: str) -> None:
        self._source = payload
        super().__init__(payload)

    def _skip_value_separation(self) -> None:
        # A tab may separate a value from its indicator, but must not become
        # part of a block's indentation on the following line.
        while self.peek() in (" ", "\t"):
            self.forward()

    def fetch_value(self) -> None:
        super().fetch_value()
        self._skip_value_separation()

    def fetch_block_entry(self) -> None:
        super().fetch_block_entry()
        self._skip_value_separation()

    def fetch_tag(self) -> None:
        # The scanner reaches this hook only for an actual unquoted token.
        # Image examples inside quoted text and block scalars stay untouched.
        start = self.index
        value_prefix = self._source[start - self.column : start].rstrip(" \t")
        if not self.flow_level and (value_prefix.endswith(":") or value_prefix.lstrip() == "-"):
            match = _IMAGE_VALUE.match(self._source, start)
            if match is not None:
                self.save_possible_simple_key()
                self.allow_simple_key = False
                mark = self.get_mark()
                self.forward(len(match[1]))
                self.tokens.append(ScalarToken(match[1], False, mark, self.get_mark()))
                return
            if self._source.startswith("![", start):
                raise yaml.YAMLError("Incomplete Markdown image property")
        super().fetch_tag()


def is_properties_mapping(payload: str) -> bool:
    """Check structure without constructing custom tags or complex keys."""
    try:
        return isinstance(yaml.compose(payload, Loader=_PropertiesLoader), MappingNode)
    except (yaml.YAMLError, RecursionError):
        return False


def property_values(payload: str) -> dict[str, Any]:
    """Project safely readable values from a structurally accepted properties map."""
    root = yaml.compose(payload, Loader=_PropertiesLoader)
    if not isinstance(root, MappingNode):
        return {}
    loader = yaml.SafeLoader("")
    try:
        try:
            # Preserve the ordinary YAML constructor's alias and merge semantics.
            values = loader.construct_document(root)
            if isinstance(values, dict):
                return values
        except (yaml.YAMLError, ValueError, OverflowError):
            pass
        # An opaque tag or complex key cannot populate a field. Read each
        # sibling independently so a bad value cannot discard a valid title.
        fields: dict[str, Any] = {}
        for key_node, value_node in root.value:
            field = yaml.SafeLoader("")
            try:
                key = field.construct_document(key_node)
                if isinstance(key, str):
                    fields[key] = field.construct_document(value_node)
            except (yaml.YAMLError, ValueError, OverflowError):
                continue
            finally:
                field.dispose()
        return fields
    finally:
        loader.dispose()
