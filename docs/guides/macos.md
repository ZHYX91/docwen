# macOS installation and first launch

[简体中文](macos.zh-CN.md)

DocWen's macOS packages target **Apple Silicon (arm64), macOS 14 or later**.
Intel Macs are not a package target. Download the matching assets from
[DocWen Releases](https://github.com/ZHYX91/docwen/releases). Select a release that actually lists
`macos-arm64` assets. A platform mentioned
in the source documentation does not add packages to an earlier immutable release.

The packages retain **ad-hoc code signatures**. They are **not Developer ID signed
or notarized by Apple**. A valid ad-hoc signature does not identify a trusted publisher
to Gatekeeper. SHA-256 checksums and GitHub build provenance do not replace that trust.

## Install the desktop app

1. Download `DocWen-<version>-macos-arm64.tar.gz` and the release's `SHA256SUMS.txt`.
   To compare the archive digest, run `shasum -a 256` followed by the downloaded archive's path.
2. Extract the archive and move the complete `DocWen.app` to Applications or your user's Applications folder.
   Keep the app bundle intact; do not move files out of its Contents directory.
3. Double-click `DocWen.app`. If macOS blocks it because the developer cannot be verified
   or Apple cannot check it for malware, first confirm the source and checksum.
4. If you trust this copy, open **System Settings → Privacy & Security → Open Anyway**,
   then confirm **Open**. This creates an exception for this app. The button normally
   appears after an attempted launch; availability depends on macOS and administrator policy.

Follow [Apple's current opening instructions](https://support.apple.com/102445).
Do not disable Gatekeeper globally. A malware, revoked-software or damaged-app warning
is not the same as an unidentified-developer warning: stop and investigate the download.
Updating to a new copy can require another approval. Apple signing and notarization may be added later.

## Command line and optional dependencies

The desktop app includes its matching CLI at:

```bash
/Applications/DocWen.app/Contents/Resources/payload/DocWenCLI --help
```

For a CLI-only installation, extract `DocWenCLI-<version>-macos-arm64.tar.gz`,
keep the entire extracted directory together and run `./DocWenCLI --help` from it.
The standalone CLI is also ad-hoc signed and subject to macOS security policy.
If the system blocks it, follow Apple's per-app instructions where available;
managed computers may require administrator assistance.

No Python installation is needed for either package. LibreOffice is optional and must
be installed separately for legacy Office formats and Office-to-PDF routes; macOS
packages do not use Windows Word/WPS COM automation. OCR models are bundled for offline use.
HEIC/HEIF support depends on the image backend included in the installed build. Check it with:

```bash
/Applications/DocWen.app/Contents/Resources/payload/DocWenCLI doctor --json
```

Installing `pillow-heif` into system Python does not add it to a packaged app. A source
checkout obtains it through `uv sync --frozen --all-extras`; if that backend reports a
missing native `libheif`, install it with `brew install libheif` and check the source
environment with `uv run --frozen docwen doctor --json`. Discovery reports backend
availability; it is not a decoding test of every HEIC/HEIF variant.

## Settings and updates

Packaged GUI and CLI share `~/Library/Application Support/docwen` by default, including
`configs` and `logs`. The app does not write settings inside its signed bundle.
Quit DocWen before replacing the complete app with a newer copy. Replacing or moving
the application preserves this user profile. `DOCWEN_DATA_DIR` remains an explicit
whole-profile override; do not point it into the app bundle.

CI checks extracted package contents, signatures and selected functionality. Offscreen
GUI checks do not prove interactive Finder, browser quarantine, manual security approval,
or every Mac/Office version. Those observations are recorded separately from automated results.
