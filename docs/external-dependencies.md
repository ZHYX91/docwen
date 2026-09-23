# External dependencies / 外部依赖

DocWen prefers bundled or process-local components. External applications are discovered through explicit adapters and are never treated as silently available.

DocWen 优先使用随包或进程本地组件。外部应用通过明确 adapter 发现，不得被静默假定为可用。

## Office backends / Office 后端

- Microsoft Office and WPS may provide Windows COM conversion routes.
- LibreOffice provides registered-install or explicitly selected executable routes.
- DOC/WPS/RTF → DOCX pre-conversion is backend-agnostic: `software.default_priority.word_processors` controls the attempt order across WPS Writer, Microsoft Word, and LibreOffice. A WPS source does **not** require WPS Writer; Microsoft Word remains a legal Windows COM candidate and failure proceeds to the next legal backend.
- Backend selection consumes the admitted source format and configured priority; it must not infer that a `.wps` filename binds conversion to one vendor application.
- Adapter priority and supported source/target pairs are route-specific. ODT is narrower and deliberately excludes WPS Writer.
- Cancellation owns the child process/profile lifecycle and must clean request-owned profiles.

For DOC/WPS/RTF → DOCX pre-conversion, the Windows COM candidates are therefore compatibility backends rather than format owners. The configured order is policy; successful opening/export by the selected backend is the runtime fact. This distinction is important because a `.wps` suffix alone is not evidence that WPS Writer must be installed.

对于 DOC/WPS/RTF → DOCX 预转换，Windows COM 后端是“兼容性后端”，不是格式所有者：`software.default_priority.word_processors` 决定 WPS 文字、Microsoft Word、LibreOffice 的尝试顺序。WPS 来源**不要求必须安装 WPS 文字**；Microsoft Word 仍是合法候选，某个后端失败后继续尝试下一个合法后端。后端选择依据已准入的来源格式和用户配置，不得仅凭 `.wps` 扩展名把转换绑定到某一家软件。ODT 的候选范围更窄，明确排除 WPS 文字。

## Network boundary / 网络边界

The supported GUI and CLI entry points deny all DNS/name resolution and AF_INET/AF_INET6 `bind`, `connect`, `connect_ex`, `sendto`, and `sendmsg` operations in the main DocWen Python process. This protects against accidental egress by in-process dependencies without changing Windows named pipes or Unix-domain sockets. Separately launched processes are not governed by the CPython audit hook. Microsoft Office, WPS, LibreOffice and the dedicated Office helper therefore retain their own and the operating system's network policy. Production source is also checked to prevent general-purpose Python/Qt network-client imports. This is not an operating-system sandbox and does not claim to contain hostile native code.

The guarantee is intentionally limited to CPython events that can be audited. Ordinary payload writes on an IP socket that was already connected or inherited before the guard became active do not emit a `socket.send` audit event and are outside the claimed boundary. Supported packaged entry points activate the guard before product and Qt imports, reject unused bundled Qt networking components, and do not intentionally create such a socket; this remains an accidental-egress defence rather than hostile-code containment.

受支持的 GUI 和 CLI 入口会阻止 DocWen Python 主进程内的全部 DNS/名称解析及 AF_INET/AF_INET6 的 `bind`、`connect`、`connect_ex`、`sendto`、`sendmsg` 操作，在不改变 Windows 命名管道或 Unix 域套接字的前提下防止进程内依赖意外联网。单独启动的进程不受 CPython 审计钩子管理，因此 Microsoft Office、WPS、LibreOffice 及专用 Office helper 仍遵循自身和操作系统的网络策略。生产源码还通过静态门禁止引入通用 Python/Qt 网络客户端。该机制不是操作系统级沙箱，也不宣称能够约束恶意原生代码。

该保证只覆盖 CPython 能够审计的事件。若某个 IP 套接字在守卫启用前已经连接或被继承，其普通载荷写入不会产生 `socket.send` 审计事件，因此不属于承诺边界。受支持的打包入口会在产品和 Qt 模块导入前启用守卫、拒绝携带未使用的 Qt 网络组件，并且不会有意创建这类套接字；这仍是防止依赖意外出站的纵深防御，而不是对恶意代码的隔离。

## OCR and table structure / OCR 与表格结构

RapidOCR models are stored under `models/rapidocr/`. Language selection is request-scoped. A successful OCR result is normal completion; the UI may show an informational quality notice, but success does not create a warning. No-text and operationally degraded outcomes remain typed and user-visible.

Table structure recognition reuses the same admitted OCR boxes/text/confidence and the existing NumPy/OpenCV/ONNX Runtime stack. DocWen embeds only the CPU SLANet inference/matching subset adapted from RapidTable/PaddleOCR under Apache-2.0; it does not add RapidTable's downloader/configuration dependencies or run a second OCR pass. The pinned `slanet-plus.onnx` resource is installed as `models/rapidtable/slanet-plus.onnx`. Production builds acquire it before packaging, enforce the frozen SHA-256, and the runtime never downloads it. `DOCWEN_RAPIDTABLE_MODEL` may point source/development runs to the same model file.

RapidOCR 模型位于 `models/rapidocr/`，语言选择按请求隔离。OCR 成功属于正常完成：界面可以显示普通质量说明，但不会因此把任务降为警告；未检测到文字和运行故障仍保留类型化状态并向用户明确展示。

表格结构识别复用同一次 OCR 已取得的文字框、文本和置信度，并继续使用项目现有 NumPy/OpenCV/ONNX Runtime。DocWen 仅内置由 RapidTable/PaddleOCR Apache-2.0 代码适配的 CPU SLANet 推理/匹配子集，不引入其下载器、配置依赖，也不会重复执行 OCR。固定的 `slanet-plus.onnx` 在打包后位于 `models/rapidtable/slanet-plus.onnx`；生产构建在打包前取得模型并核对冻结 SHA-256，运行时绝不联网下载。源码/开发环境可用 `DOCWEN_RAPIDTABLE_MODEL` 指向同一模型。

### OpenCV distribution ownership / OpenCV 分发所有权

`rapidocr-onnxruntime==1.4.4` declares `opencv-python`, while the PDF-to-DOCX fallback declares
`opencv-python-headless`. Those distributions install the same `cv2` paths and must never coexist.
DocWen explicitly selects `opencv-python-headless==4.13.0.92`: the product uses PySide6 for windows,
and neither its production sources nor the exercised OCR/PDF paths require OpenCV HighGUI. The root
uv configuration removes only RapidOCR 1.4.4's conflicting dependency edge. The embedded SLANet
table runtime imports DocWen's already-selected headless `cv2` and therefore adds no second OpenCV
distribution. A RapidOCR version change therefore fails the exact contract until the declaration and
runtime behavior are reviewed again.

`rapidocr-onnxruntime==1.4.4` 声明 `opencv-python`，而 PDF→DOCX 兜底声明
`opencv-python-headless`；两者会安装同一组 `cv2` 路径，禁止共存。DocWen 明确选择
`opencv-python-headless==4.13.0.92`：窗口由 PySide6 提供，产品源码及已验证的 OCR/PDF
路径均不需要 OpenCV HighGUI。根 uv 配置仅移除 RapidOCR 1.4.4 的冲突依赖边；一旦
RapidOCR 版本变化，精确合同会先失败，必须重新审查声明与真实运行行为。

The upstream macOS headless wheel retains the Cocoa backend; its package name does not promise
`GUI: NONE` on macOS. Windows and Linux wheels use `NONE`. DocWen still requires one headless
package owner and does not call HighGUI. See the [upstream build policy](https://github.com/opencv/opencv-python/blob/4.x/setup.py).

上游 macOS headless 包保留 Cocoa 后端，包名不表示 macOS 的构建信息必须为 `GUI: NONE`；
Windows 与 Linux 包使用 `NONE`。DocWen 仍严格要求唯一 headless 包所有者，且不调用 HighGUI。

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
