# Public API boundaries / 公共 API 边界

Public imports are the symbols exported by package `__init__` modules, declared protocols, entry points and documented CLI/schema contracts. Other symbols are private implementation details.

公共 API 包括包级 `__init__` 导出、已声明 protocol、入口点以及文档化 CLI/schema 契约；其余符号均为私有实现。

## Rules / 规则

- Applications use core protocols and injected ports instead of plugin/runtime internals.
- Plugins use core contracts and their own package-local implementation.
- GUI and CLI do not import concrete plugin modules.
- Private modules may change without compatibility guarantees.
- Compatibility re-exports require an explicit supported import contract and one owning implementation.
- New public symbols need documentation, typing and direct tests.

Import Linter and repository guards enforce package-level dependency boundaries.
