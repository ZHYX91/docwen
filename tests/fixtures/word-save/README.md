# Native Word save fixtures

These two small DOCX files contain only synthetic test data: text formatting, a
merged table, named captions and references, an H7 heading, an endnote, and a
generated image. The tests read the fixed packages without launching Office or
requiring a source Markdown sidecar.

The input was created by
[`word_saved_fixture.py`](../../support/native/word_saved_fixture.py) using the
current Markdown converter. It uses caption ownership map version 2 with a
`sequence_scope` shared by its generated SEQ fields. The exact generated input
SHA-256 was
`bd8068412fa0feba2d2f45bd5fc3e4757ece465e0ed1f2a7ca1f3963beeec98c`;
the preparation helper SHA-256 was
`a3239715281b880e52e50d2c0726d1bf5950da0ab98bca1c1af06605ac2b40cb`.

Microsoft Word 16.0, build 16.0.20326, on Windows 10 edited and saved the files
through its native application API on 2026-11-12:

- `edited-markdown-extensions.docx`: Word replaced the ordinary body paragraph
  `待 Word 替换的普通正文。` with
  `Word 实测追加：从独立 DOCX 的当前内容回转。`, then saved the writable copy.
- `expanded-caption-control.docx`: Word opened a separate copy of the saved
  positive fixture and appended an ordinary paragraph inside the final Figure
  content control. The new paragraph explicitly uses `wdStyleNormal`. Native
  assertions proved a change from two to three paragraphs while preserving the
  original paragraphs, field codes/results, and image counts. After saving, the
  original two paragraph XML trees remained identical except for revision and
  paragraph identity attributes. Conversion must report
  `one caption and one logical object` and produce no artifacts.

Word's native package changes are preserved, including renamed styles, custom
XML parts, content-control IDs, and `w:sdtEndPr` with its optional ending run
formatting. No document or semantic XML was repaired after Word saved it.
Only author names and creation/modification timestamps in `docProps/core.xml`
were removed; every other package part is byte-identical to its native saved
source. ZIP entry timestamps and permissions were normalized.

| File | Native saved SHA-256 | Checked-in SHA-256 |
| --- | --- | --- |
| `edited-markdown-extensions.docx` | `571662dcc4ad5ffd0f5e05093ba7f3023d0f7db737500f30946a18f0a77786ea` | `df8a2696a8be5a0bbd03cbfd00b59414dd278b8c5c3c7ac0c00ab1fae75ee944` |
| `expanded-caption-control.docx` | `b6c981a9392c5fb5d7878b5e314df2a10862f3f82047c9656cacdfe9d50e0403` | `3df66b0dd1af19a791e512f54d5a49eed0590bdef364652719cc46f7dc6fba2a` |

To refresh these fixtures, generate a fresh input with the helper, repeat both
native Word edits on writable copies, and check the original and redacted
packages with the three fixed-fixture cases in
[`test_markdown_extension_policy.py`](../../integration/test_markdown_extension_policy.py).
The helper only prepares source bytes; running it alone does not constitute a
native Word save. These fixtures exercise semantic round-tripping and rejection
of an expanded content control; they do not claim visual layout acceptance.
