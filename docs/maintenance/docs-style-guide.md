# Documentation style

[简体中文](docs-style-guide.zh-CN.md)

## Location and names

- Root files are entry points: `README.md`, `CONTRIBUTING.md`, `SECURITY.md`,
  `CHANGELOG.md` and the existing license/notice files. Development instructions have
  one authority in `docs/development.md` and its language version.
- `docs/README.md` indexes guides, development, architecture and specifications.
  README translations live in `docs/i18n/README.<locale>.md`; task guides live in
  `docs/guides/`. Screenshots live in `docs/assets/screenshots/`.
- Public product READMEs retain `# DocWen` and the shared native-language navigation
  with `·` separators and absolute GitHub `blob/main` links. Keep existing locale codes.
- Ordinary document names use lowercase English words separated by hyphens. Keep
  conventional `README.md` names and existing license filenames. Dates belong to
  historical records, not long-lived guides.
- Full language versions are separate: English `name.md`, Chinese `name.zh-CN.md`.
  Technical terms, quoted interface labels and examples may contain multiple languages.
  Engineering references do not require newly invented translations.
- Accepted technical decisions stay with the source in `docs/architecture/decisions/`.
  Local plans, migration receipts, run logs and historical acceptance records belong
  to the maintainer workspace. A standalone clone contains the instructions needed
  to use, build, test and release the product.

## Writing and maintenance

- Describe current behavior, owner and limits. Distinguish planned behavior explicitly.
  Keep stage IDs, completion narratives and test-count snapshots out of current guides;
  current instructions must not depend on private plans or historical acceptance records.
- Give each changing fact one authoritative location; link to it instead of copying it.
- Keep user guides focused on tasks and implementation details in their owning reference.
- Update maintained language pairs together. Preserve navigation and explicit limits;
  source checks do not establish packaged or physical acceptance.
- Moving a document requires updating links, repository URLs, packaging and guards.
  Preserve license texts and historical original bytes.

Check local links, documentation guards and `git diff --check` after editing.
