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

## Gongwen metadata and body / 公文元数据与正文

Numeric table cells alone are not evidence of a copy identifier. A table value needs an adjacent
explicit copy-ID label, or must lead the document with following official-header evidence
(security, urgency or document number). Ordinary data tables remain in the body. Leading zeroes
in legitimate copy identifiers remain strings.

Ordinary Word lists remain Markdown lists: numbering counters, starts/restarts, nesting and
continuation lines are retained. Ordered markers normalize to decimal Markdown markers and bullets
to `-`. Gongwen heading-number removal/replacement applies only to recognised headings, not to
ordinary list items.

表格单元格是纯数字不能单独证明其为份号；需要相邻的明确份号标签，或位于文首且后续有密级、
紧急程度、发文字号等公文版头依据。普通数据表留在正文，合法份号保留前导零字符串。
普通 Word 列表保留计数、起始/重启、嵌套与续行；有序标记规范为 Markdown 十进制数字，无序标记为
`-`。公文标题序号清理/替换只作用于识别为标题的段落，不删除普通列表标记。
