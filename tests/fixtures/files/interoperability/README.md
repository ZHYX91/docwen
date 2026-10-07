# Independent dialect corpora

These byte-identical fixtures freeze the source plugins' independent syntax oracles.
CI reads only these checked-in copies and never a sibling checkout or installed plugin.

| Fixture | Upstream repository / commit | SHA-256 |
| --- | --- | --- |
| structural-tables.json | ZHYX91/obsidian-structural-tables / fc8559003ca95318646b6ddd42634666d0fe5275 | ee46512fc5100d0c1691a804636e24221f970f3ae1e0a82b55484c646e60293c |
| number-suite.json | ZHYX91/obsidian-number-suite / 6bb8c417f098fca27aa1a2731e41f3401c2fd102 | 09099dc4c94b7af6523fbe8f014bd7384ebac793223dcc2e5a93214ceb687e76 |

Both upstream paths are `tests/fixtures/interoperability-syntax-contract.json`.
The consumer tests map each upstream oracle to DocWen's native representation:
spreadsheet cells retain Markdown source, DOCX AST cells contain parsed inline text,
invalid merge geometry produces DocWen diagnostics, and an invalid delimiter remains
literal source. Note labels become opaque per-type identities while reference order
and repeated references remain observable. These are static consumer regressions,
not Office, Obsidian, package or publication acceptance.

The Number Suite corpus explicitly leaves the Structural Tables reference-alias pipe
escaping order unresolved; that case is not promoted to a passed contract here.
