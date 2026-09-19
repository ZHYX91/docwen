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

Application conversion planning separates typed requests and results (`conversion_contracts`), capability
bindings (`conversion_capabilities`), option validation (`conversion_options`), runtime request construction
(`conversion_requests`), and the plan/accept/execute/cancel lifecycle (`conversion_service`). Plans own copies
of their input options and public result data; mutating a returned discovery or plan object cannot alter a
later execution. Input integrity and active capabilities are checked again at acceptance.

`conversion_routes` composes the same preconversion chain used by the controller and resolves every step
against the runtime catalog. `optimization_selection` requires all selected inputs to have an available
chain and intersects their route options. GUI localization and settings remain presentation concerns;
missing or unavailable Office bridge routes cannot be hidden by an available final DOCX optimizer.

应用层分别维护请求/结果类型、能力绑定、参数校验、运行请求构造与计划/接纳/执行/取消生命周期。
计划持有自己的参数和结果快照，外部修改发现结果或计划对象不会改变执行；接纳时仍重验输入完整性和当前能力。
转换路线计划与控制器使用同一前置转换链，逐步核对真实运行时路线。优化选择要求每个输入的完整链可用，
批量选项取交集；GUI 只负责显示与交互，不能因最终 DOCX 优化器可用而掩盖前置 Office 转换不可用。

- Core does not import application, runtime, apps, plugins or bundle.
- Plugins do not import application/runtime/apps/bundle.
- GUI and CLI do not deep-import plugin implementations or runtime internals.
- Network policy, configuration persistence and output placement each have one owner.
- GUI view models own presentation state; widgets render or forward interaction and do not become persistence sources.

These rules are enforced by Import Linter, repository guards and Pyright.

这些规则由 Import Linter、仓库门禁和 Pyright 持续检查。
