# Mermaid images / Mermaid 图片

DocWen keeps Mermaid code blocks by default. To export them as images, install an optional local rendering environment and choose **Settings → Markdown syntax → Mermaid → Render as image**. The main installer does not include Node.js, Mermaid CLI or Chrome, and DocWen does not download them automatically.

DocWen 默认保留 Mermaid 代码块。需要图片时，请自行安装本地渲染环境，在**设置 → Markdown 语法 → Mermaid → 渲染为图片**中选择图片模式。主安装包不内置 Node.js、Mermaid CLI 或浏览器，应用也不会自动安装。

## Local installation / 本地安装

The Windows combination exercised during development is Node.js **24.18.0**, Mermaid CLI **11.17.0**, Mermaid **11.17.2**, Puppeteer **25.11.0**, Chrome headless shell **153.0.8010.36**. The version gate accepts CLI and Mermaid **11.16+ within 11.x**; other accepted combinations still require a successful test render on the user's machine. Linux/macOS browser and font availability must be checked locally.

本轮 Windows 实测组合为上述版本。CLI 和 Mermaid 本身都需为 11.16 以上的 11.x；版本匹配仍需通过本机测试渲染，不能等同于浏览器已可用。泳道使用 `swimlane-beta`，语法见 [Mermaid 官方说明](https://mermaid.js.org/syntax/swimlanes)。

1. Install [Node.js](https://nodejs.org/en/download), including npm. Reopen DocWen after changing PATH.
2. In a dedicated directory, install the tested packages:

   ```sh
   npm install --save-exact @mermaid-js/mermaid-cli@11.17.0 mermaid@11.17.2 puppeteer@25.11.0
   ```

3. If the package manager blocked Puppeteer's installation script, install its matching browser from that same directory:

   ```sh
   npx --no-install puppeteer browsers install chrome-headless-shell
   ```

4. In DocWen, select `node_modules/.bin/mmdc.cmd` (Windows) or `node_modules/.bin/mmdc` (Linux/macOS). Apply/OK saves the path; Cancel discards changes. A blank path enables automatic discovery through PATH and common npm installation locations. `DOCWEN_MERMAID_CLI` is an optional environment fallback.
5. Choose **Detect again**, then **Test rendering**. Only a completed PNG test displays **Available**. Test rendering keeps the selected output mode and does not convert an open document.

先安装 Node.js，再在专用目录运行上述命令；如果安装脚本被包管理器拦截，需要手动下载配套浏览器。随后在 DocWen 选择本地 CLI 路径，点“重新检测”和“测试渲染”。只有测试图成功才显示可用。路径按常规 Apply/OK 保存，Cancel 放弃；测试不会改变代码/图片模式。

Browser downloads may require a working network/proxy. Run the browser-install command as the same OS user that runs DocWen; if using `PUPPETEER_CACHE_DIR`, expose the same value to both installation and DocWen. Do not fix missing browser errors by disabling the browser sandbox. See [Puppeteer installation and blocked scripts](https://pptr.dev/guides/installation).

浏览器下载失败时先检查网络和代理；安装和运行 DocWen 应使用同一系统账户。设置了 `PUPPETEER_CACHE_DIR` 时两者必须一致。可用性提示会区分依赖缺失、版本不兼容、浏览器缺失及测试失败。

## Rendering contract / 渲染边界

- Images use a local browser, strict Mermaid settings and a 30-second limit per diagram. Cancellation terminates the owned process tree. No online rendering service is used. External/local-file images and unserved network resources are rejected; self-contained diagram data is supported.
- PNGs preserve aspect ratio and fit the template's body section, columns and paragraph indents. Large reductions produce `MD2DOCX-MERMAID-SCALED`; inspect text readability and split very large diagrams if needed. Chinese labels use installed Microsoft YaHei / Noto Sans CJK SC where available.
- Invalid syntax, missing dependencies, browser failure and timeout retain the original visible code block, with an indexed `MD2DOCX-MERMAID-FALLBACK` warning. Other diagrams continue. Cancellation remains cancellation. The source Markdown is unchanged.
- A rendered diagram is an ordinary DOCX image. DOCX → Markdown exports an image and does not reconstruct Mermaid source. The exact resolved-v4 input route continues to preserve fenced source.

图片按模板正文节、分栏与段落缩进等比缩放，明显缩小会提示检查文字。每张图单独计时，失败时保留可见代码并报告图表序号和原因，其他图继续；取消不会伪装成成功回退。反向转换只处理普通图片，不还原 Mermaid 源码。
