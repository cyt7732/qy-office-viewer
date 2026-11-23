# Changelog / Release List

本项目遵循语义化版本规范（[Semantic Versioning](https://semver.org/)），记录所有正式版本的变更内容。

---

## [v1.2.2] - 2026-11-23

### 新增与优化 (Added & Fixed)
- **补齐 macOS 与 Linux 系统原生联动与防挂起加固**：
  - **macOS**：`pbcopy` 显式注入 UTF-8 Locale 环境变量（`buildMacUtf8Env`），修复 GUI 子进程环境下复制中文路径乱码的问题；`/api/action/reveal` 通过 `open -R` 在访达（Finder）中高亮定位并返回专属提示文案。
  - **Linux**：`/api/action/reveal` 新增三级定位链路——优先通过 Freedesktop D-Bus `org.freedesktop.FileManager1.ShowItems`（`gdbus` → `dbus-send`）在默认文件管理器中高亮选中目标文件，不支持时自动回退 `xdg-open` 打开父目录；`/api/action/copy-clipboard` 按 Wayland/X11 会话自适应切换 `wl-copy` / `xclip` / `xsel`，并使用 `stdio: ['pipe', 'ignore', 'ignore']` + `stdin` 错误监听杜绝剪贴板守护进程持锁挂死与 `EPIPE` 异常。
- **前端跨平台细节适配**：
  - `app.js` 按客户端系统动态区分「资源管理器 / 访达（Finder） / 文件管理器」提示文案，且 `normPath()` 仅在 Windows 忽略路径大小写，macOS/Linux 保留大小写精度；
  - `index.html` 默认定位提示改为中性跨平台文案。
- **跨平台同步脚本与文档**：新增纯 LF 换行的 `sync.sh`（macOS / Linux 一键同步部署），同步补齐 `sync.ps1` 的 `CHANGELOG.md` 复制及 `README.md` 三端部署指引。

---

## [v1.2.1] - 2026-11-23

### 修复 (Fixed)
- **修复 Windows 无法唤起 WPS / 默认程序与资源管理器定位的回归问题**：
  - 修复 `main.mjs` 中通过 `spawn(..., { detached: true })` 启动 `pwsh.exe` / `powershell.exe` 时，因 `DETACHED_PROCESS` 导致控制台宿主未执行脚本即静默退出（退出码 0）的问题，改为 `detached: false` + `windowsHide: true`，并结合 `.NET ProcessStartInfo`（`UseShellExecute = $true`）可靠触发 Win32 `ShellExecuteEx`；
  - 新增子进程错误监听、退出码与 CLIXML 错误信息回传、3 秒防阻塞超时保护，以及 `pwsh` 启动失败自动回退至系统 `powershell.exe` 的兜底机制；
  - 资源管理器定位（`/api/action/reveal`）改为直接以 `windowsVerbatimArguments: true` 唤起 `explorer.exe /select,"<path>"`（移除多余 PowerShell 中转与 `windowsHide: true`），兼顾毫秒级响应、特殊字符路径兼容与零命令注入风险；
  - 修复 `resolveWinPowerShellExe()` 在 WindowsApps Store 版 `pwsh.exe`（App Execution Alias）上 `existsSync` 误判失效的问题，改用 `existsSync || lstatSync` 准确识别。

---

## [v1.2.0] - 2026-11-22

### 修复 (Fixed)
- **Word 红头与彩色排版保留**：移除 `styles.css` 中对 `section.docx` 后代元素的 `color: #111827 !important` 强制覆盖，改由页面容器提供默认基色，完整保留红头标题、印章及高亮标注色彩。
- **杜绝同名相近文件静默张冠李戴**：移除后端 `resolveActualPath` 模糊自动替换逻辑，改为严格精确匹配；路径未命中时返回 `exists: false` 及同目录相似候选列表，在前端明确提示并由用户点击确认。
- **PDF 中文 CMap、懒渲染、划词复制与竞态隔离**：
  - 补齐 `pdfjs-dist@3.11.174` 全套 168 个离线 `.bcmap` 文件及 `/api/vendor/cmap` 安全路由，解决非嵌字中文 PDF 丢字问题；
  - 引入 `IntersectionObserver` 可视区懒渲染、`TextLayer` 透明文字层（支持划词复制）、缩放防抖清晰重绘与 `Ctrl/Cmd + 滚轮` 缩放；
  - 新增 `activeRenderToken` 并发取消令牌并在切换时销毁上一个 `PDFDocumentProxy`，彻底杜绝快速切换导致的串页与内存泄漏。
- **Excel 吸顶表头、脱机拖拽与 CSV 编码修复**：
  - 渲染表格时自动将首行提升为 `<thead><th>`，让表头在纵向滚动时真正吸顶；
  - 修复脱机拖拽 Excel/CSV 时未构建 `sheetsMap` 与工作表标签导致的空白 Bug，统一多端表格挂载逻辑；
  - 新增 BOM 检测 + UTF-8 严格校验 + `GB18030` 自动回退解码，兼容无 BOM UTF-8 与中文 GBK/CP936 CSV。
- **公文落款后处理防误伤与换行控制**：缩小 `fixDocxSignatures` 扫描范围至正文末尾 15 个非空段落（排除表格/页眉/页脚/列表及含句中标点的句子），取消激进全局扫描与强制 `nowrap` 溢出。
- **Windows 特殊字符直开与文件访问安全加固**：
  - Windows 下「打开 / 定位 / 复制剪贴板」全面改用 `pwsh`（回退 `powershell.exe`）`-EncodedCommand` + 单引号安全转义，彻底消除 `cmd.exe /c start` 命令注入风险及 `&`、`^`、`%`、`,` 路径断链问题；
  - `/api/file/raw` 与 `/api/file/meta` 补充预览扩展名白名单校验（非白名单返回 HTTP 403）。
- **纯文本预览补齐**：实现 `.txt`、`.md`、`.json`（自动格式化缩进）预览与拖放解析。

### 优化与重构 (Changed & Performance)
- **后端 `fs.watch` 变更追踪 + 热重载状态保持**：新增 `/api/file/watch-state` 内存态监听接口，移除鼠标移入与普通滚轮发请求的事件监听；热重载同一文档时自动保留滚动位置、当前 Sheet 标签与缩放比例。
- **按需懒加载与静态资源内存缓存**：移除 `index.html` 首屏阻塞加载的全家桶 `<script>`，改为按文档类型动态加载；后端对 `vendor/` 与 `cmaps/` 启用进程内存缓存。
- **主题适配与代码瘦身**：移除锁死主题的 `localStorage` 内联覆盖与硬编码视口暗色底色，清理死代码并将纯逻辑抽离至 `lib.mjs`。

---

## [v1.1.0] - 2026-11-22

### 新增 (Added)
- **PDF 纯离线高清流式渲染**：集成 Mozilla PDF.js（`pdf.min.js` + `pdf.worker.min.js`），支持高 DPI Canvas 渲染、白底卡片阴影、自适应视口宽度、缩放控制及悬浮页码徽标。
- **Excel 服务端无损 CJK 解析通道**：新增 `/api/file/spreadsheet` 接口，在 Node.js 侧加载完整版 `xlsx.full.min.js` 解析工作簿（降低浏览器编码乱码风险），并保留前端 SheetJS 自动降级回退能力。
- **自适应动态心跳与热重载**：引入分阶段文件变更检测（活跃期 2.5s / 平稳期 8s / 静置 90s 自动休眠），结合窗口聚焦、页面可见性与交互事件实时感知文件改动并自动刷新。
- **智能文件名路径对齐**：新增 `resolveActualPath()`，当传入路径因标点或空格微小差异未直接命中时，支持在同目录按扩展名与字符相似度匹配候选文件。
- **公文落款排布辅助处理**：新增 `fixDocxSignatures()` 对 `GB/T 9704-2012` 常见成文日期与署名段落进行对齐修正。
- **文件拖放预览**：支持从系统拖入本地 `.docx` / `.xlsx` / `.xls` / `.csv` / `.pdf` 文件进行预览。

### 优化与变更 (Changed)
- **紧凑顶栏与窄屏适配**：顶栏收束至 35px 高度，宽度小于 360px 时自动折叠品牌名称与工具按钮文字，最大化文档阅读区。
- **Excel 白底仿真表格视觉**：统一工作表标签栏与白色网格风格，增加隔行高亮、悬停反馈与长内容单元格自动换行。
- **主题变量同步保活**：补充宿主 CSS 变量同步逻辑，避免重载后丢失基础配色。
- **同步插件元数据与版本**：统一将 `plugin.json`、`sidecars/panel/package.json` 及 `/api/status` 版本更新为 `1.1.0`，补充 `displayName` 与 `repository` 字段。

### 安全与清理 (Security & Chore)
- **移除本机硬编码路径**：清理 `app.js` 中硬编码的本地绝对路径回退逻辑，全部统一通过 Sidecar `/vendor/*` 路由加载。
- **精简仓库体积**：清理未引用的重复大文件（`vendor/pdf.worker.js`、`vendor/xlsx.mini.min.js`）及根目录临时调试截图。
- **彻底清理历史测试文件**：从 Git 全历史中移除早期调试提交的测试公文 `.docx`，确保仓库历史洁净。

---

## [v1.0.0] - 2026-10-10

### 新增 (Added)
- **Antigravity Sidecar 扩展基础架构**：完成 `plugin.json`、`sidecars/panel/sidecar.json`、`main.mjs` 后端服务与辅助侧边栏前端骨架搭建。
- **Sidecar 胶囊 URL 路由**：支持 `sidecar://qy-office-viewer/panel/?file=<encoded-path>` 直达预览，自动过滤非 Office 文件。
- **Word（`.docx`）离线渲染引擎**：集成 `docx-preview` 与 `JSZip`，支持多页白底纸张仿真、横向全宽自适应缩放、`ResizeObserver` 动态响应与缩放工具栏。
- **Excel 表格初版预览**：集成 SheetJS（`xlsx.mini.min.js`）实现基础表格读取与多工作表切换。
- **系统原生联动（P0）**：
  - 非阻塞极速唤起系统默认程序打开文档（支持 Windows / macOS / Linux）；
  - 在系统文件管理器中高亮定位目标文件；
  - 跨平台剪贴板路径复制（含系统原生命令兜底）；
  - 打开与定位操作采用静默触发，减少多余弹窗打扰。
- **部署辅助**：新增 `sync.ps1` 本地同步脚本与插件 SVG 矢量图标（`assets/logo.svg`）。
