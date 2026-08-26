# Documentation style / 文档规范

Repository documentation describes current behavior. Historical explanations belong to Git history and release changelogs.

仓库文档只描述当前行为。历史解释由 Git 历史和版本更新日志承担。

## Structure / 结构

- The root `README.md` is the sole English user guide. Translations remain separate files under `docs/user-guides/`.
- Localized guides use BCP 47 names such as `README.zh-CN.md`. Each public README starts with `# DocWen`, followed by the same native-language navigation using `·` and absolute GitHub `blob/main` links.
- Product, architecture, specification and maintenance documents contain Chinese and English in one file where prose is maintained long term.
- Maintainer execution plans stay outside the public repository. Public documents contain only durable current behavior, contracts and verification instructions.
- Screenshots used by current guides/specifications live under `docs/assets/screenshots/`.

- 根目录 `README.md` 是唯一英文用户手册；其他语言译本分别保存在 `docs/user-guides/`。
- 本地化手册使用 `README.zh-CN.md` 等 BCP 47 文件名。所有公开 README 均以 `# DocWen` 开头，随后使用相同的原生语言导航、`·` 分隔符和 GitHub `blob/main` 绝对链接。
- 维护者执行计划不进入公开仓库；计划中的稳定结论必须先吸收到相应当前文档。

## Writing / 写法

- State the current owner, entry point, behavior and boundary.
- Avoid stage IDs, test-count snapshots, completion narratives and comparison with retired implementations.
- Describe compatibility as a current contract, not as a migration story.
- Link to current files and stable symbols; avoid private absolute paths.
- Update both language sections in the same change.

## Required checks / 必需检查

1. All local links and referenced files resolve.
2. Current documentation contains no completed-plan or acceptance-report dependency.
3. Route/config/output/UI changes update their owning document.
4. Documentation guards and `git diff --check` pass.
5. User guides do not expose internal guard names or implementation chronology.
