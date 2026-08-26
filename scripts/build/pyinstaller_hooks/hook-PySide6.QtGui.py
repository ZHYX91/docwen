"""Collect the QtGui runtime without unused network-dependent plugins.

PyInstaller's stock QtGui hook deliberately collects every image-format,
generic-input, platform and platform-input-context plugin.  Optional plugins
in the PySide6 wheel pull QtPdf, QtQml/QtQuick or QtNetwork into an otherwise
widgets-only DocWen bundle.  DocWen does not use Qt's PDF image reader, TUIO
touch input, bundled virtual keyboard, or VNC platform backend, so omit those
plugin entry points while preserving the rest of the stock hook result.
"""

from __future__ import annotations

from pathlib import Path

from PyInstaller.utils.hooks.qt import add_qt6_dependencies

_OMITTED_PLUGIN_STEMS = frozenset({"qpdf", "qtuiotouchplugin", "qtvirtualkeyboardplugin", "qvnc"})


def _plugin_stem(source: str) -> str:
    stem = Path(source).name.partition(".")[0].casefold()
    return stem.removeprefix("lib")


hiddenimports, binaries, datas = add_qt6_dependencies(__file__)
binaries = [entry for entry in binaries if _plugin_stem(entry[0]) not in _OMITTED_PLUGIN_STEMS]
