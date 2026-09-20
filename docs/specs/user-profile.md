# User profile paths

The composed GUI, CLI and Machine entry points resolve their profile once before loading configuration or templates. The resulting immutable paths also identify the GUI instance. Reloading configuration, worker threads and later environment changes cannot select another profile within that process. Restart to change a profile.

## Selection

| Selection | User configuration | Template data and state | Default log directory |
|---|---|---|---|
| `DOCWEN_DATA_DIR=<root>` | `<root>/configs` | `<root>` | `<root>/logs` |
| Extracted frozen archive | `<executable-dir>/data/configs` | `<executable-dir>/data` | `<executable-dir>/data/logs` |
| Windows package identity, including MSIX | Platform user-data directory plus `configs` | Platform user-data directory | Platform user-data directory plus `logs` |
| Source or installed Python execution | Platform user-config directory plus `configs` | Platform user-data directory | Platform user-log directory |

Platform directories use the application name `docwen` without an author component. An unknown Windows package-identity probe uses the protected-installation behavior. It never authorizes writing beside an executable. Package resources, including base configuration and built-in templates, remain read-only.

`DOCWEN_CONFIG_DIR` selects the exact user-configuration directory independently of the other components. `DOCWEN_LOG_DIR` selects a log root and writes under its `logs` child. A truthy `DOCWEN_LOG_TO_TEMP` selects the system temporary directory's `docwen/logs`, unless `DOCWEN_LOG_DIR` is also supplied. These two logging overrides disable the GUI directory controls; selecting a whole profile does not. Without a logging environment override, saved custom or temporary logging preferences still apply.

Whitespace-only environment values are absent. Relative profile selections are expanded against the startup working directory and canonicalized once. A GUI launched by Machine receives the startup environment with explicit profile paths made absolute, even if its working directory differs. Consumers preserve the relevant platform home and profile-directory variables, the explicit DocWen profile selectors, and temporary-directory preferences within a bounded environment. They resolve relative selectors before changing the child working directory.

GUI instance identity includes both the canonical data and configuration directories. Separate configurations over the same template data use separate GUI instances. Log preferences do not create another instance. A package upgrade keeps its user-data identity; a copied archive uses the copy's own data unless an explicit environment selection overrides it.

Windows single-instance ownership uses a process-owned named pipe in the same user scope as GUI control. Changing `TEMP`, `TMP` or `TMPDIR` cannot create another owner for the same profile. Normal shutdown closes the ownership handle after control has stopped; process termination releases it without deleting a lock file.

## Copying or importing existing data

Close processes that use the source and destination profiles. For an archive, copy the application's `data` directory together with the executable directory. To import a current profile into a new location, create a new empty directory and copy `configs/`, `templates/` and `template-state.json` into it. Select that directory with `DOCWEN_DATA_DIR` before starting DocWen or its consumer. Logs are optional and are not needed to preserve settings.

If source execution used separate platform configuration and data directories, copy the explicitly chosen user-configuration directory to the new root's `configs/` and the templates and state from the chosen data directory. Existing independent directory overrides remain explicit; selecting DATA does not cancel CONFIG or LOG. Absolute custom output/log paths inside copied settings retain their saved meaning and are not rewritten.

Template files and `template-state.json` must be copied together to preserve canonical template IDs, order, default and disabled state. User TOML overrides retain their rules. This procedure accepts the current configuration and state formats; it does not rename old keys, merge another destination profile, or scan alternative directories. For selective changes, use template import and the proofreading rule editor's existing previewed import.

## Failure behavior

Resolving a profile does not create directories or modify the environment. An existing file where a required profile directory belongs stops startup, reports the selected path and leaves it unchanged. Unreadable configuration or template state does not cause selection of a different directory. Unwritable settings/template operations report failure at their selected component and do not fall back to another profile.

If the selected file-log directory cannot be opened, file logging remains unavailable and the error is reported through console logging and runtime logging state. DocWen does not silently create a second log directory. The logging page displays an active file only when one exists, reports the failure, and does not invent a home-directory fallback.
