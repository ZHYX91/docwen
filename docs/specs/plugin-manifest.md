# Plugin manifest / 插件清单

Each plugin exports one manifest and one implementation that conforms to the core plugin protocol.

每个插件导出一份 manifest 和一个符合 core plugin protocol 的实现。

## Manifest fields / 字段

- stable plugin ID and display metadata;
- supported routes or action-only operations;
- canonical source and target formats plus aliases;
- option schema, defaults and validation constraints;
- capability and external-dependency metadata;
- typed optimization resources (`id`, display `name`, and internal `action_name`); runtime derives public scopes from the final canonical action routes;
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

The runtime registry and `docwen inspect --json` expose the composed route view. The canonical optimization
view is `docwen resources list optimizations --json`; consumers must not reconstruct it from extensions,
categories, configuration files, or action-name conventions.
