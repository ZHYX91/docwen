# Plugin manifest / 插件清单

Each plugin exports one manifest and one implementation that conforms to the core plugin protocol.

每个插件导出一份 manifest 和一个符合 core plugin protocol 的实现。

## Manifest fields / 字段

- stable plugin ID and display metadata;
- supported routes or action-only operations;
- canonical source and target formats plus aliases;
- option schema, defaults and validation constraints;
- capability and external-dependency metadata;
- typed optimization resources (`id`, display `name`, internal `action_name`, and explicit `applicable_locales`); runtime derives public scopes from the final canonical action routes;
- output/artifact expectations where required.

## Invariants / 不变量

- Registration rejects duplicate IDs and conflicting routes.
- Manifest options and runtime consumption remain synchronized.
- Plugins depend on core contracts, not runtime/application/app internals.
- A route is not advertised when its implementation is missing.
- External availability is reported separately from route declaration.
- Optimization IDs are public resource identifiers, not runtime action names. Each declaration must bind to
  at least one action route in the same manifest. Runtime derives scopes and route bindings from those routes;
  manifests must not declare parallel source-format, target-format, or scope catalogs.
- Duplicate optimization IDs, missing action routes, and unknown source categories
  make runtime capability projection fail closed.

`applicable_locales` declares the interface languages for which the optimization resource is intended,
using exact canonical locale identifiers such as `zh_CN`. It is not a list of translated UI labels.
An omitted or empty declaration does not grant GUI applicability. GUI selection intersects this declaration
with the actual input category, available conversion plan, and enabled-resource policy; it never infers
language support from a resource ID. A saved default remains stored when inapplicable, but is not executed.

This metadata is projected only through the existing in-process GUI capability catalog, as
`optimization_applicable_locales` keyed by real resource IDs. The public `docwen.optimizations` v1 resource
shape, CLI explicit optimization requests, and Machine resource/capability discovery remain unchanged.

`applicable_locales` 明确声明资源适用的界面语言，不等同于界面文案的翻译覆盖。缺省或空声明不在 GUI
提供该优化；筛选还须满足输入类别、实际可用转换链及启用策略。不适用的已保存默认值保留，但不执行。
此元数据只进入已有的进程内 GUI 投影，不改变公开 v1 资源字段或 CLI／Machine 的显式优化能力。

The runtime registry and `docwen inspect --json` expose the composed route view. The canonical optimization
view is `docwen resources list optimizations --json`; consumers must not reconstruct it from extensions,
categories, configuration files, or action-name conventions.
