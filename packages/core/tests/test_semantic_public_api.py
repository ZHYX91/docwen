"""Direct import contracts for the provider-neutral semantic public API."""

from __future__ import annotations

import pytest

import docwen_core
import docwen_core.models as public_models
from docwen_core import (
    SemanticBibliographyEntry as RootSemanticBibliographyEntry,
)
from docwen_core import SemanticBibliographyRun as RootSemanticBibliographyRun
from docwen_core import (
    SemanticCitationCluster as RootSemanticCitationCluster,
)
from docwen_core import (
    SemanticCitationItem as RootSemanticCitationItem,
)
from docwen_core import (
    is_portable_semantic_id as root_is_portable_semantic_id,
)
from docwen_core.models import (
    SemanticBibliographyEntry as ModelsSemanticBibliographyEntry,
)
from docwen_core.models import SemanticBibliographyRun as ModelsSemanticBibliographyRun
from docwen_core.models import (
    SemanticCitationCluster as ModelsSemanticCitationCluster,
)
from docwen_core.models import (
    SemanticCitationItem as ModelsSemanticCitationItem,
)
from docwen_core.models import (
    is_portable_semantic_id as models_is_portable_semantic_id,
)
from docwen_core.models.semantic_document import (
    SemanticBibliographyEntry,
    SemanticBibliographyRun,
    SemanticCitationCluster,
    SemanticCitationItem,
    is_portable_semantic_id,
)

pytestmark = pytest.mark.contract


def test_new_semantic_symbols_are_identical_across_public_import_boundaries() -> None:
    assert RootSemanticBibliographyEntry is ModelsSemanticBibliographyEntry is SemanticBibliographyEntry
    assert RootSemanticBibliographyRun is ModelsSemanticBibliographyRun is SemanticBibliographyRun
    assert RootSemanticCitationCluster is ModelsSemanticCitationCluster is SemanticCitationCluster
    assert RootSemanticCitationItem is ModelsSemanticCitationItem is SemanticCitationItem
    assert root_is_portable_semantic_id is models_is_portable_semantic_id is is_portable_semantic_id


def test_new_semantic_symbols_are_declared_by_both_public_all_lists() -> None:
    symbol_names = {
        "SemanticBibliographyEntry",
        "SemanticBibliographyRun",
        "SemanticCitationCluster",
        "SemanticCitationItem",
        "is_portable_semantic_id",
    }

    assert symbol_names <= set(docwen_core.__all__)
    assert symbol_names <= set(public_models.__all__)
