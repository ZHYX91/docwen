"""Repo-level integration tests.

Cross-plugin integration tests live here — notably the
``Markdown -> DOCX -> Markdown`` numbering round-trip, which drives
both ``docwen_plugin_markdown`` (MD -> DOCX) and
``docwen_plugin_document`` (DOCX -> MD) through the real runtime
orchestration, rather than deep-importing either plugin from the
other's package tests.

These tests honour the package boundary: they import only the
application/runtime composition root and the shared core models,
never the internals of a plugin they do not own.
"""
