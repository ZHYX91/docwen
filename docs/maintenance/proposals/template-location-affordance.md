# Template file location

The main template selector now offers **Open file location** on its context menu when the runtime registry provided a physical `TemplateInfo.path`. The selector only emits that injected path; it does not probe the filesystem or open URLs itself. The main window's existing `_open_path(..., open_parent=True)` action performs the platform-specific reveal.

The menu does not construct a filesystem path from display names or canonical template IDs. A missing source path removes the location action but leaves **Manage templates** available. Existing user template management, selection, and default policy are unchanged.

This fixes the mismatch between the documented context-menu shortcut and the current UI.
