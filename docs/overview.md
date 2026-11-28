# Overview

[简体中文](overview.zh-CN.md)

DocWen is a local-first desktop and command-line document conversion toolkit. It provides GUI and CLI entry points over one application/runtime/plugin composition and does not require a network service for normal conversion.

## Scope

- Documents: Markdown, DOCX, DOC, ODT, RTF, WPS and related PDF export routes.
- Spreadsheets: XLSX, XLS, ODS, CSV, TSV and ET routes.
- Layout and images: PDF, XPS, OFD and common raster formats.
- Imports: HTML, MHTML, EPUB, ENEX, PPTX and PPT to Markdown where declared.
- OCR, proofreading, numbering, template rendering, batch conversion and aggregate actions.

The exact supported surface is defined by [Capabilities](capabilities.md) and [Routes and actions](specs/routes-and-actions.md). Unsupported or best-effort behavior must be reported truthfully by the owning route.

## Safety

Supported GUI and CLI composition roots activate one CPython audit guard for the full lifetime of the main DocWen Python process. It denies all DNS/name resolution and AF_INET/AF_INET6 `bind`, `connect`, `connect_ex`, `sendto`, and `sendmsg` operations used by in-process Python dependencies while leaving Windows named pipes and Unix-domain sockets available. Native code that bypasses CPython audit events, separately launched processes, external Office/WPS/LibreOffice applications and the dedicated Office helper are outside this boundary. The guard is defence in depth against accidental dependency egress, not an operating-system sandbox. Source files are treated as immutable inputs; final outputs are placed through the runtime finalizer with collision and containment checks.
An inherited or already-connected IP socket can perform an ordinary payload write without a `socket.send` audit event, so that case is explicitly outside the guarantee; supported packaged entries activate before product/Qt imports and do not intentionally establish one.
