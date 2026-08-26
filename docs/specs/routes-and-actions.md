# Routes and actions / 路线与操作

Plugin manifests are the executable source of truth for routes, actions, sources, targets and options. Inspect the current composition with:

插件 manifest 是 route、action、source、target 和 option 的可执行事实源。当前组合可通过以下命令检查：

```powershell
docwen inspect --json
docwen resources list formats --json
```

## Rules / 规则

- Every advertised source/target pair resolves to exactly one registered plugin route.
- Options must be declared in the manifest schema and consumed by the owning request path.
- Action-only routes do not masquerade as ordinary source/target conversions.
- Aliases normalize at the route boundary; result metadata uses canonical identifiers.
- Unsupported routes fail before destructive or external work.
- GUI and CLI derive available actions from the same manifest/runtime composition.
- Public optimization resource IDs are declared separately from internal action names. Runtime projection
  validates their route/scope bindings; consumers must not equate the two strings or rebuild applicability.

## Runtime capability projection / 运行时能力投影

Manifests may attach route capability rules for supported platforms, required and optional dependency
gates, and stable limitation identifiers. `resources list formats --json` projects those rules over the
currently loaded composition. Every projected route includes its operation (`conversion` or `action`),
source, target, owning plugin, availability state, missing gates and limitations. Action routes remain
explicit matrix entries; they do not become public CLI commands.

The projection is fail-closed: an unknown required gate makes the affected route unavailable. A known
optional gate may leave the route available while reporting a limitation. A runtime that initialized with
zero routes returns a successful empty matrix, while runtime initialization/query failure is a typed error.

Route additions require manifest tests, option-consumption tests, entry-point coverage and an update to [Capabilities](../capabilities.md).
