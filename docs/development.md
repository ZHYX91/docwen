# Development

[简体中文](development.zh-CN.md)

DocWen uses Python 3.12 and a uv workspace for core, application, runtime, apps,
plugins and bundle packages. Run these commands from the repository root.

## Prepare the environment

CI and production builds pin **uv 0.12.0**. Dependency and lock maintenance may use
compatible **uv 0.12.x** without changing the production toolchain.

```sh
uv sync --frozen --all-extras
```

Activate `.venv\Scripts\activate` on Windows or `source .venv/bin/activate` on
macOS/Linux. Do not use `pip install -e`: pip does not apply this workspace's scoped
dependency exclusion and can install two OpenCV distributions over the same `cv2`
files. See [external dependencies](external-dependencies.md).

## Source runs and checks

QA and source launchers keep generated data in an explicitly owned engineering
directory. Existing maintainer workspaces are reused. For a standalone clone,
initialize a new directory outside the repository:

```sh
python tools/workspace_root.py --init ../docwen-workspace --repository .
python tools/qa.py --workspace-root ../docwen-workspace
python tools/run_source.py --workspace-root ../docwen-workspace gui
python tools/run_source.py --workspace-root ../docwen-workspace cli -- --version
```

Alternatively, set `DOCWEN_WORKSPACE_ROOT` to that directory's absolute path. The
standard `repos/docwen` layout and linked worktrees can discover a registered
workspace. Initialization creates the selected directory or registers the repository
with an existing root; it does not require private maintainer plans.

The launcher leases a run directory and passes `DOCWEN_RUNTIME_ROOT` to the product.
Normal installed applications use their application cache. CI may provide explicit
`--pytest-runtime-root`, `--own-pytest-runtime` and report directories.

`python tools/qa.py` runs formatting, lint, type checks and the fast test suite.
Select `--suite full` when a full run is required. Use the owning checks for a small
change; follow [testing guidance](testing.md) for packaged and interactive scope.

```sh
ruff format --check .
ruff check .
python tools/typecheck.py
```

For local formatting, use `ruff format .`; review `ruff check . --fix` changes before
keeping them. Pyright covers `packages/` and excludes test/build/resource data.
Optional hooks use `pre-commit install` and `pre-commit run --all-files`.

## Application resources

`assets/icon.svg` is the sole application-icon source. Regenerate and verify its
committed PNG and ICO derivatives after editing it:

```sh
python scripts/maintenance/generate_app_icons.py
python scripts/maintenance/generate_app_icons.py --check
```

Follow [GUI visual ownership](specs/gui-behavior.md#visual-system--视觉规范): baseline
dimensions belong in shared tokens, static geometry uses `set_metric`, and responsive
text layout uses font metrics. Use the existing licensed icon collection or canonical
brand source. MSIX packaging requires Windows SDK MakePri beside MakeAppx and must
include `resources.pri`.

## Change ownership

- Keep shared contracts and pure logic in core, orchestration in application/runtime,
  and format behavior in plugins. Declare direct library dependencies in
  \`packages/core/pyproject.toml\`. Shared DOCX semantics use \`python-docx\` and \`lxml\`.
- Update owning tests and documentation when behavior, configuration, output or release
  requirements change. Keep compatibility decisions explicit; new compatibility shims
  or secondary state sources need an explicit public contract.
- Preserve inputs and unrelated local changes. Use reviewed, scoped text edits; avoid
  in-place shell substitutions that mishandle Windows paths, encoding or line endings.
- Update imports, references and packaging after renames. Let Git recognize renames;
  do not remove and recreate unrelated content to make a move look clean.

See [contributing](../CONTRIBUTING.md), [architecture](architecture.md) and
[documentation style](maintenance/docs-style-guide.md).
