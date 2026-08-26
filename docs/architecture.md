# Architecture / 架构

DocWen uses a layered workspace. Dependencies point inward toward stable contracts; composition belongs to the bundle package.

DocWen 采用分层 workspace。依赖指向稳定的内层契约，最终组合由 bundle 包负责。

```text
core <- plugins
  ^       ^
  |       |
runtime  application
   \       /
     bundle
     /   \
   cli   gui
```

## Layers / 分层

- `docwen_core`: data contracts, shared parsing, OCR, links, formula and pure services.
- `docwen_application`: request admission, pre-conversion orchestration and use-case control.
- `docwen_runtime`: plugin registry, route resolution, workspaces, configuration, security, IPC and output finalization.
- `docwen_plugin_*`: format-specific route implementations; plugins depend on core contracts, not application or runtime internals.
- `docwen_cli` and `docwen_gui`: presentation and interaction layers.
- `docwen_bundle`: composition root, entry points and adapter wiring.

## Execution flow / 执行主链

1. CLI or GUI builds a request from user input and a configuration snapshot.
2. Application admits the request, owns protective input copies and coordinates optional pre-conversion.
3. Runtime resolves a manifest route and invokes the owning plugin.
4. The plugin returns typed results and artifact declarations.
5. Runtime finalizes outputs transactionally and emits one truthful terminal result.

## Boundaries / 边界

- Core does not import application, runtime, apps, plugins or bundle.
- Plugins do not import application/runtime/apps/bundle.
- GUI and CLI do not deep-import plugin implementations or runtime internals.
- Network policy, configuration persistence and output placement each have one owner.
- GUI view models own presentation state; widgets render or forward interaction and do not become persistence sources.

These rules are enforced by Import Linter, repository guards and Pyright.

这些规则由 Import Linter、仓库门禁和 Pyright 持续检查。
