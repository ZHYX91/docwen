"""PluginManifest — metadata a plugin declares at registration time."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

_RESOURCE_ID_PATTERN = re.compile(r"^[a-z0-9][a-z0-9_-]*$")


@dataclass(frozen=True, slots=True)
class OptimizationResourceSpec:
    """Public optimization resource declared by one plugin manifest.

    ``id`` is the stable consumer-facing identifier. ``action_name`` is the
    internal runtime action it binds to; consumers must never infer one from
    the other. Applicable scopes are derived from the final canonical action
    routes so manifests do not maintain a second route table.
    """

    id: str
    name: str
    action_name: str

    def __post_init__(self) -> None:
        values: dict[str, object] = {"id": self.id, "name": self.name, "action_name": self.action_name}
        for field_name, value in values.items():
            if not isinstance(value, str) or not value or value != value.strip():
                raise ValueError(f"Optimization resource {field_name} must be a non-empty trimmed string.")
        if not _RESOURCE_ID_PATTERN.fullmatch(self.id):
            raise ValueError(f"Invalid optimization resource id: {self.id}")
        if not _RESOURCE_ID_PATTERN.fullmatch(self.action_name):
            raise ValueError(f"Invalid optimization action name: {self.action_name}")

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "name": self.name,
            "action_name": self.action_name,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> OptimizationResourceSpec:
        expected_keys = {"id", "name", "action_name"}
        unexpected = sorted(set(data) - expected_keys)
        if unexpected:
            raise ValueError(f"Optimization resource contains unsupported field(s): {', '.join(unexpected)}")
        return cls(
            id=data["id"],
            name=data["name"],
            action_name=data["action_name"],
        )


@dataclass(frozen=True, slots=True)
class RouteCapabilityRule:
    """Typed capability requirements applied to matching manifest routes.

    Empty selector tuples mean "any".  Multiple matching rules compose: the
    runtime intersects their platform sets and unions their required/optional
    capability gates and limitations.  The rules describe executable route
    constraints; current-machine availability is evaluated by the runtime.
    """

    source_formats: tuple[str, ...] = ()
    target_formats: tuple[str, ...] = ()
    action_names: tuple[str, ...] = ()
    operations: tuple[str, ...] = ()
    platforms: tuple[str, ...] = ()
    required_capabilities: tuple[str, ...] = ()
    optional_capabilities: tuple[str, ...] = ()
    limitations: tuple[str, ...] = ()

    def matches(self, route: RouteSpec) -> bool:
        """Return whether this rule applies to *route*."""

        return (
            (not self.source_formats or route.source_format in self.source_formats)
            and (not self.target_formats or route.target_format in self.target_formats)
            and (not self.action_names or route.action_name in self.action_names)
            and (not self.operations or ("action" if route.action_name else "conversion") in self.operations)
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "source_formats": list(self.source_formats),
            "target_formats": list(self.target_formats),
            "action_names": list(self.action_names),
            "operations": list(self.operations),
            "platforms": list(self.platforms),
            "required_capabilities": list(self.required_capabilities),
            "optional_capabilities": list(self.optional_capabilities),
            "limitations": list(self.limitations),
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> RouteCapabilityRule:
        return cls(
            source_formats=tuple(data.get("source_formats", ())),
            target_formats=tuple(data.get("target_formats", ())),
            action_names=tuple(data.get("action_names", ())),
            operations=tuple(data.get("operations", ())),
            platforms=tuple(data.get("platforms", ())),
            required_capabilities=tuple(data.get("required_capabilities", ())),
            optional_capabilities=tuple(data.get("optional_capabilities", ())),
            limitations=tuple(data.get("limitations", ())),
        )


@dataclass(slots=True)
class RouteSpec:
    """Describes one conversion route a plugin can handle."""

    source_format: str
    """Source format (e.g. ``"markdown"``, ``"docx"``, ``"pdf"``)."""

    target_format: str
    """Target format (e.g. ``"docx"``, ``"md"``, ``"pdf"``)."""

    action_name: str = ""
    """Optional action name for non-conversion operations (``"validate"``, ``"merge"``, etc.)."""

    label: str = ""
    """Human-readable label for UI presentation."""

    options_schema: dict[str, Any] = field(default_factory=dict)
    """JSON Schema describing accepted options for this route."""

    def to_dict(self) -> dict[str, Any]:
        return {
            "source_format": self.source_format,
            "target_format": self.target_format,
            "action_name": self.action_name,
            "label": self.label,
            "options_schema": dict(self.options_schema),
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> RouteSpec:
        return cls(
            source_format=data["source_format"],
            target_format=data["target_format"],
            action_name=data.get("action_name", ""),
            label=data.get("label", ""),
            options_schema=dict(data.get("options_schema", {})),
        )


@dataclass(slots=True)
class HonestyRoute:
    """A route-level honesty declaration — machine-verifiable contract data.

    Stored in a plugin manifest's ``extra`` under ``"not_implemented_routes"``,
    ``"unavailable_routes"`` or ``"office_bridge_routes"``.  Unlike the free-form prose strings used
    previously, this carries the ``(source, targets)`` pair as structured
    fields so consistency checks (manifest == code) can verify it directly
    instead of parsing human-readable text.  ``description`` preserves the
    human-readable explanation (e.g. ``"optional backend required"``).
    """

    source: str
    """Source format (e.g. ``"xps"``, ``"layout"``, ``"document"``)."""

    targets: list[str]
    """Target formats this declaration covers (e.g. ``["docx", "doc", "odt"]``)."""

    description: str = ""
    """Human-readable explanation of the declaration."""

    def to_dict(self) -> dict[str, Any]:
        return {
            "source": self.source,
            "targets": list(self.targets),
            "description": self.description,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> HonestyRoute:
        return cls(
            source=data["source"],
            targets=list(data["targets"]),
            description=data.get("description", ""),
        )


@dataclass(slots=True)
class PluginManifest:
    """Metadata a plugin declares so the runtime can discover and route to it.

    Each plugin package provides exactly one ``PluginManifest`` instance.
    """

    plugin_id: str
    """Unique plugin identifier (e.g. ``"docwen_plugin_markdown"``)."""

    name: str
    """Human-readable plugin name."""

    version: str
    """Semver version string."""

    description: str = ""
    """Short description of what this plugin does."""

    author: str = ""
    """Plugin author."""

    routes: list[RouteSpec] = field(default_factory=list)
    """Every conversion route this plugin claims to handle."""

    requires: list[str] = field(default_factory=list)
    """Other plugin ids this plugin depends on (optional)."""

    platforms: tuple[str, ...] = ("windows", "linux")
    """Product platforms on which the plugin's routes are supported."""

    capability_rules: list[RouteCapabilityRule] = field(default_factory=list)
    """Typed dependency/platform/limitation rules for executable routes."""

    optimization_resources: list[OptimizationResourceSpec] = field(default_factory=list)
    """Public optimization resources bound to action routes in this manifest."""

    extra: dict[str, Any] = field(default_factory=dict)
    """Extension point for plugin-specific metadata."""

    def to_dict(self) -> dict[str, Any]:
        return {
            "plugin_id": self.plugin_id,
            "name": self.name,
            "version": self.version,
            "description": self.description,
            "author": self.author,
            "routes": [r.to_dict() for r in self.routes],
            "requires": list(self.requires),
            "platforms": list(self.platforms),
            "capability_rules": [rule.to_dict() for rule in self.capability_rules],
            "optimization_resources": [resource.to_dict() for resource in self.optimization_resources],
            "extra": dict(self.extra),
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> PluginManifest:
        return cls(
            plugin_id=data["plugin_id"],
            name=data["name"],
            version=data["version"],
            description=data.get("description", ""),
            author=data.get("author", ""),
            routes=[RouteSpec.from_dict(r) for r in data.get("routes", [])],
            requires=list(data.get("requires", [])),
            platforms=tuple(data.get("platforms", ("windows", "linux"))),
            capability_rules=[RouteCapabilityRule.from_dict(rule) for rule in data.get("capability_rules", [])],
            optimization_resources=[
                OptimizationResourceSpec.from_dict(resource) for resource in data.get("optimization_resources", [])
            ],
            extra=dict(data.get("extra", {})),
        )
