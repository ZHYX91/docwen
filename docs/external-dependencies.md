# External dependencies / 外部依赖

DocWen prefers bundled or process-local components. External applications are discovered through explicit adapters and are never treated as silently available.

DocWen 优先使用随包或进程本地组件。外部应用通过明确 adapter 发现，不得被静默假定为可用。

## Office backends / Office 后端

- Microsoft Office and WPS may provide Windows COM conversion routes.
- LibreOffice provides registered-install or explicitly selected executable routes.
- Adapter priority and supported source/target pairs are route-specific.
- Cancellation owns the child process/profile lifecycle and must clean request-owned profiles.

## Network boundary / 网络边界

The supported GUI and CLI entry points deny all DNS/name resolution and AF_INET/AF_INET6 `bind`, `connect`, `connect_ex`, `sendto`, and `sendmsg` operations in the main DocWen Python process. This protects against accidental egress by in-process dependencies without changing Windows named pipes or Unix-domain sockets. Separately launched processes are not governed by the CPython audit hook. Microsoft Office, WPS, LibreOffice and the dedicated Office helper therefore retain their own and the operating system's network policy. Production source is also checked to prevent general-purpose Python/Qt network-client imports. This is not an operating-system sandbox and does not claim to contain hostile native code.

The guarantee is intentionally limited to CPython events that can be audited. Ordinary payload writes on an IP socket that was already connected or inherited before the guard became active do not emit a `socket.send` audit event and are outside the claimed boundary. Supported packaged entry points activate the guard before product and Qt imports, reject unused bundled Qt networking components, and do not intentionally create such a socket; this remains an accidental-egress defence rather than hostile-code containment.

受支持的 GUI 和 CLI 入口会阻止 DocWen Python 主进程内的全部 DNS/名称解析及 AF_INET/AF_INET6 的 `bind`、`connect`、`connect_ex`、`sendto`、`sendmsg` 操作，在不改变 Windows 命名管道或 Unix 域套接字的前提下防止进程内依赖意外联网。单独启动的进程不受 CPython 审计钩子管理，因此 Microsoft Office、WPS、LibreOffice 及专用 Office helper 仍遵循自身和操作系统的网络策略。生产源码还通过静态门禁止引入通用 Python/Qt 网络客户端。该机制不是操作系统级沙箱，也不宣称能够约束恶意原生代码。

该保证只覆盖 CPython 能够审计的事件。若某个 IP 套接字在守卫启用前已经连接或被继承，其普通载荷写入不会产生 `socket.send` 审计事件，因此不属于承诺边界。受支持的打包入口会在产品和 Qt 模块导入前启用守卫、拒绝携带未使用的 Qt 网络组件，并且不会有意创建这类套接字；这仍是防止依赖意外出站的纵深防御，而不是对恶意代码的隔离。

## OCR

RapidOCR models are stored under `models/rapidocr/`. Language selection is request-scoped. OCR is best effort: unavailable engines, recognition failure and no-text results remain distinguishable and produce user-visible warnings where applicable.

RapidOCR 模型位于 `models/rapidocr/`，语言选择按请求隔离。OCR 属于尽力而为能力：引擎不可用、识别失败和未检测到文字必须保持可区分，并在适用场景给出可见警告。

### OpenCV distribution ownership / OpenCV 分发所有权

`rapidocr-onnxruntime==1.4.4` declares `opencv-python`, while the PDF-to-DOCX fallback declares
`opencv-python-headless`. Those distributions install the same `cv2` paths and must never coexist.
DocWen explicitly selects `opencv-python-headless==4.13.0.92`: the product uses PySide6 for windows,
and neither its production sources nor the exercised OCR/PDF paths require OpenCV HighGUI. The root
uv configuration removes only RapidOCR 1.4.4's conflicting dependency edge. A RapidOCR version
change therefore fails the exact contract until the declaration and runtime behavior are reviewed again.

`rapidocr-onnxruntime==1.4.4` 声明 `opencv-python`，而 PDF→DOCX 兜底声明
`opencv-python-headless`；两者会安装同一组 `cv2` 路径，禁止共存。DocWen 明确选择
`opencv-python-headless==4.13.0.92`：窗口由 PySide6 提供，产品源码及已验证的 OCR/PDF
路径均不需要 OpenCV HighGUI。根 uv 配置仅移除 RapidOCR 1.4.4 的冲突依赖边；一旦
RapidOCR 版本变化，精确合同会先失败，必须重新审查声明与真实运行行为。

The authoritative environment check is the frozen uv lock plus the single-owner/runtime contract in
`tests/test_repo/test_opencv_distribution_contract.py`. A generic metadata-only `pip check` does not
understand uv's scoped exclusion and will repeat RapidOCR's upstream declaration; installing the second
OpenCV wheel to silence that message would reintroduce nondeterministic file ownership.

权威环境门是冻结的 uv lock 与 `tests/test_repo/test_opencv_distribution_contract.py` 的唯一所有者／
运行合同。通用、仅看包元数据的 `pip check` 不理解 uv scoped exclusion，会重复 RapidOCR 的上游
声明；不得为了消除该提示而安装第二个 OpenCV wheel，否则会重新引入不确定的文件所有权。

## Platform limits / 平台限制

Windows-only COM routes are unavailable on POSIX. LibreOffice-backed routes require a compatible executable and fonts. Packaging checks verify bundled resources; `docwen doctor` reports the current machine's actual capability state.

`docwen resources list formats --json` also projects dependency and platform gates onto every loaded
conversion/action route. Unavailable routes remain visible with `available: false` and typed limitation
identifiers, so consumers can distinguish an unsupported route from a successful empty composition.
