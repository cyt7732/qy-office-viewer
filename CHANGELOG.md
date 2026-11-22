# Changelog / Release List

本项目遵循语义化版本规范（[Semantic Versioning](https://semver.org/)），记录所有正式版本的变更内容。

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
