# macOS installation and first launch / macOS 安装与首次打开

DocWen's macOS packages target **Apple Silicon (arm64), macOS 14 or later**.
Intel Macs are not a package target. Download the matching assets from
[DocWen Releases](https://github.com/ZHYX91/docwen/releases); an older release may not contain Mac assets.

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
HEIC/HEIF support remains conditional on the installed image backend.

## Settings and updates

Packaged GUI and CLI share `~/Library/Application Support/docwen` by default, including
`configs` and `logs`. The app does not write settings inside its signed bundle.
Quit DocWen before replacing the complete app with a newer copy. Replacing or moving
the application preserves this user profile. `DOCWEN_DATA_DIR` remains an explicit
whole-profile override; do not point it into the app bundle.

## 中文说明

- 仅提供 Apple 芯片 arm64、macOS 14 及以上的包；不提供 Intel 包。旧版本可能没有 Mac 附件。
- 保留 ad-hoc 签名，**未使用 Developer ID 签名，也未经过 Apple 公证**。
- 下载并核对 SHA-256，解压后将完整 `DocWen.app` 移入“应用程序”。
- 首次双击后若提示无法验证开发者或无法检查恶意软件，确认来源可信后，到
  **系统设置 → 隐私与安全性 → 仍要打开 → 打开**。受管理设备可能不允许设置例外。
  参阅 [Apple 官方说明](https://support.apple.com/zh-cn/102445)，不要全局关闭系统安全检查。
- “将损坏电脑”、恶意软件或“已损坏”提示不能直接视作普通首次打开提醒，应先核查下载与签名。
- 配置和日志默认保存在 `~/Library/Application Support/docwen`；退出应用后整体替换 `.app` 即可更新，保留用户资料。
- GUI/CLI 不需要安装 Python。旧 Office 格式及 Office 转 PDF 需另行安装 LibreOffice；独立 CLI 也受系统安全策略限制。

CI checks extracted package contents, signatures and selected functionality. Offscreen
GUI checks do not prove interactive Finder, browser quarantine, manual security approval,
or every Mac/Office version. Those observations are recorded separately from automated results.
