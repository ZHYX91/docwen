"""A layout manifest can belong to a resource-only result directory."""

from dataclasses import replace

import pytest

from docwen_core.models import (
    ArtifactBundleValidationError,
    BundleDraft,
    BundleDraftArtifact,
    BundleEntry,
    BundleRelation,
    validate_artifact_bundle_draft,
)

pytestmark = pytest.mark.contract


def _draft() -> BundleDraft:
    return BundleDraft(
        artifacts=(
            BundleDraftArtifact("table", "resource", "table.csv", "table.csv", "text/csv", "result/table.csv"),
            BundleDraftArtifact(
                "manifest",
                "resource",
                "docwen-node.json",
                "docwen-node.json",
                "application/vnd.docwen.document-node+json",
                "result/docwen-node.json",
            ),
        ),
        entries=(BundleEntry("table", "supplementary", 0, True),),
        relations=(BundleRelation("resource_of", "manifest", "table", "manifest", 0),),
        layout_schema="docwen.document_node.v1",
    )


def test_resource_result_owns_its_typed_layout_manifest() -> None:
    validate_artifact_bundle_draft(_draft())


@pytest.mark.parametrize("change", ["untyped", "different_name", "not_preferred", "ordinary_resource"])
def test_manifest_resource_owner_does_not_relax_other_resource_relations(change: str) -> None:
    draft = _draft()
    if change == "untyped":
        draft = replace(
            draft, artifacts=(draft.artifacts[0], replace(draft.artifacts[1], media_type="application/json"))
        )
    elif change == "different_name":
        draft = replace(draft, artifacts=(draft.artifacts[0], replace(draft.artifacts[1], suggested_name="other.json")))
    elif change == "not_preferred":
        draft = replace(draft, entries=(replace(draft.entries[0], preferred=False),))
    else:
        draft = replace(draft, relations=(replace(draft.relations[0], role="image"),))
    with pytest.raises(ArtifactBundleValidationError, match="incompatible"):
        validate_artifact_bundle_draft(draft)
