# macOS 安装与首次打开

[English](macos.md)

以下为安装速查，完整命令和更新说明见英文指南。

- 仅提供 Apple 芯片 arm64、macOS 14 及以上的包；不提供 Intel 包。旧版本可能没有 Mac 附件。
- 保留 ad-hoc 签名，**未使用 Developer ID 签名，也未经过 Apple 公证**。
- 选择附件中实际列出 `macos-arm64` 的发行版；源码文档中的平台说明不会给旧版补添安装包。
- 下载并核对 SHA-256，解压后将完整 `DocWen.app` 移入“应用程序”。
- 首次双击后若提示无法验证开发者或无法检查恶意软件，确认来源可信后，到
  **系统设置 → 隐私与安全性 → 仍要打开 → 打开**。受管理设备可能不允许设置例外。
  参阅 [Apple 官方说明](https://support.apple.com/zh-cn/102445)，不要全局关闭系统安全检查。
- “将损坏电脑”、恶意软件或“已损坏”提示不能直接视作普通首次打开提醒，应先核查下载与签名。
- 配置和日志默认保存在 `~/Library/Application Support/docwen`；退出应用后整体替换 `.app` 即可更新，保留用户资料。
- GUI/CLI 不需要安装 Python。旧 Office 格式及 Office 转 PDF 需另行安装 LibreOffice；独立 CLI 也受系统安全策略限制。
- HEIC/HEIF 取决于成品包内的图片后端。使用 `/Applications/DocWen.app/Contents/Resources/payload/DocWenCLI doctor --json` 检查；给系统 Python 安装 `pillow-heif` 不会改变成品包。源码环境通过 `uv sync --frozen --all-extras` 安装依赖；仅在后端提示缺少原生库时使用 `brew install libheif`，再以 `uv run --frozen docwen doctor --json` 复查。能力检测不等于逐一验证所有 HEIC/HEIF 编码变体。

CI 检查解包内容、签名及选定功能；离屏 GUI 检查不证明 Finder 交互、下载隔离标记、人工安全批准或所有 Mac/Office 版本均已验收。
