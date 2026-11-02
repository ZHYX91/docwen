"""Resource applicability is declared explicitly, separately from translations."""

import pytest

from docwen_core.models.manifest import OptimizationResourceSpec

pytestmark = pytest.mark.unit


def test_locale_declaration_round_trip_and_legacy_absence() -> None:
    resource = OptimizationResourceSpec("future", "Future resource", "future-action", ("en_US", "ja_JP"))
    assert OptimizationResourceSpec.from_dict(resource.to_dict()) == resource
    legacy = {"id": "legacy", "name": "Legacy", "action_name": "legacy-action"}
    assert OptimizationResourceSpec.from_dict(legacy).applicable_locales == ()


@pytest.mark.parametrize("value", [None, "zh_CN", ["zh-CN"], [""], ["zh_CN", "zh_CN"], [42]])
def test_invalid_resource_locale_declarations_are_rejected(value: object) -> None:
    with pytest.raises(ValueError, match="applicable_locales"):
        OptimizationResourceSpec.from_dict(
            {"id": "resource", "name": "Resource", "action_name": "action", "applicable_locales": value}
        )
