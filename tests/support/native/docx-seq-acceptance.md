# DOCX caption SEQ / REF native acceptance

This is a reproducible host acceptance fixture, **not a claim that native field updates pass**. It exercises the
ordinary Markdown caption writer with two generated captions and two generated REF fields. One template contains
two fixed captions using the native `SEQ Table` identifier, before and after the body anchor. Generated captions use
one separately owned `DW_T_<scope>` counter; their REF fields keep the original target bookmarks. The paired control
has the same fixed content with literal numbers. Both files are produced by the real converter; context adapters only
supply request configuration, progress and a private staging directory.

## Create and inspect

Run from this checkout with the locked development dependencies and its own source packages available:

```sh
uv run --frozen --all-extras python -m tests.support.native.docx_seq_acceptance create /absolute/new/seq-case
git rev-parse HEAD
```

The destination must not already exist. Keep `source.md`, both templates, both `*-before-update.docx` outputs and
`before-update.json`. The manifest records package hashes, instructions and cached results and explicitly retains
`native_acceptance: pending`. The helper refuses to import a converter from another checkout.

For each supported Word/WPS host, copy the two before-update outputs to new files. Record the host/version, OS,
source commit and original hashes. Open without repairing the package; record the initial visible numbers and any
repair message. Update all document fields using the host's normal operation, save to the new names, close and reopen.
Record generated caption values, both REF values, the fixed template captions and screenshots before and after.
Repeat the update once to check stability. Pagination, print preview and PDF/printing remain separate observations.

```sh
uv run --frozen --all-extras python -m tests.support.native.docx_seq_acceptance inspect /absolute/seq-case/shared-counter-after-update.docx
uv run --frozen --all-extras python -m tests.support.native.docx_seq_acceptance inspect /absolute/seq-case/control-after-update.docx
```

Save each JSON observation alongside the host/version evidence. A pre-update source assertion or python-docx reopen
does not calculate SEQ and cannot replace these observations. The current source contract gives the two generated
captions/REFs caches `1` and `2`; the shared-counter output also contains the template's cached `1` and `2` under its
independent `Table` identifier. Both pairs must remain `1` and `2` after repeated native updates. Compare the actual
native results with those intended generated values and the control. If native
update changes generated numbering or REF meaning because of template fields, record that as a confirmed host-level
failure requiring a counter-ownership decision. Do not silently add a reset switch, unlink fields or weaken proof.

## Evidence status / 证据状态

本样本用于核验模板既有同名 SEQ 与新题注在原生更新域后的影响。源码回归只证明：输出包含哪些指令与缓存、
生成题注 ownership 能区分模板题注、REF 指向相应生成目标。它没有启动 Word/WPS，也没有证明更新域、分页、
打印或 PDF 行为。当前状态保持待原生验收；两份对照文件及更新前 JSON 不能被记为宿主验收通过。
