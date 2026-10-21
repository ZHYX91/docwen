"""Keep python-docx template paths traversable in frozen applications.

The library opens ``docx/parts/../templates`` through its module ``__file__``.
On POSIX the intermediate ``parts`` directory must exist even when all XML
resources were collected. Keep these modules in the PYZ and on disk so the
library's unmodified resource lookup works for both GUI and CLI payloads.
"""

from PyInstaller.utils.hooks import collect_data_files

datas = collect_data_files("docx")
module_collection_mode = {"docx.parts": "pyz+py"}
