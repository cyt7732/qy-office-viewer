# QY Office Viewer

> 基于 **Google Antigravity UI Extension & Sidecar SDK** 构建的侧边栏 Office 文档离线预览插件。  
> 专为智能体生成与处理本地 Office 文档的场景设计，实现「对话胶囊一键直达侧边栏预览 + 毫秒级唤起本地 Office/WPS 程序」的无缝衔接。

---

## 支持格式一览

| 格式 | 扩展名 | 侧边栏离线预览 | 唤起本地默认程序 / 定位 / 复制路径 |
| :--- | :--- | :---: | :---: |
| **Word 文档** | `.docx` | 支持（`docx-preview` 仿真纸张排版） | 支持 |
| **Word 旧格式** | `.doc` | 暂不支持（提示一键本地打开） | 支持 |
| **Excel 表格** | `.xlsx` / `.xls` / `.csv` | 支持（Node 服务端 SheetJS 无损解析 + 多 Sheet 切换） | 支持 |
| **PDF 文档** | `.pdf` | 支持（PDF.js 高 DPI Canvas 多页流式渲染） | 支持 |
| **PowerPoint** | `.pptx` / `.ppt` | 暂不支持（提示一键本地打开） | 支持 |

---

## 核心特性

### 1. 纯离线多格式侧边栏预览
- **Word（`.docx`）仿真排版**：内置 `docx-preview` 与 `JSZip`，白底纸张流式呈现段落、表格、页眉页脚与图片；默认横向自适应视口宽度，支持缩放（放大 / 缩小 / `1:1` / 自适应宽度）与窄屏响应。
- **Excel（`.xlsx` / `.xls` / `.csv`）多工作表浏览**：内置完整版 SheetJS（`xlsx.full.min.js`），优先经由 Node.js 服务端解析（保障 CJK/CP936 编码），支持多 Sheet 标签切换、隔行高亮、长单元格自动换行；服务异常时自动降级前端本地解析。
- **PDF（`.pdf`）高保真流式渲染**：集成 Mozilla PDF.js（离线 Worker），按设备像素比动态计算渲染分辨率，白底卡片阴影展示、自适应视口宽度并附悬浮页码徽标。
- **本地拖放直览**：支持将本地 Office 文件直接拖入侧边栏即时预览。

### 2. 系统本地程序无缝联动
- **一键本地打开**：非阻塞瞬时唤起 Microsoft Office、WPS Office 或系统默认关联程序（支持 Windows / macOS / Linux）。
- **文件管理器高亮定位**：在 Windows 资源管理器或 macOS 访达中直接定位并高亮选中目标文件。
- **路径复制**：内置剪贴板复制及系统原生命令兜底，一键复制物理绝对路径。

### 3. Antigravity Sidecar 胶囊协议与智能热重载
- **标准胶囊协议**：在 Antigravity 对话中点击以下格式的 Markdown 胶囊即可在辅助侧边栏展开并定位对应文档：
  ```markdown
  [文件名.docx](sidecar://qy-office-viewer/panel/?file=<URI编码后的文件绝对路径>)
  ```
- **自适应热重载**：
  - 活跃期（前 25 秒）每 2.5 秒轻量检查磁盘修改时间戳与大小，智能体改稿保存后自动刷新；
  - 平稳期（25～90 秒）自动降频，静置超 90 秒后零请求休眠；
  - 窗口重新聚焦或重新点击胶囊时立即同步最新文件内容。
- **紧凑顶栏与主题适配**：35px 紧凑顶栏（窄于 360px 自动折叠文字标签），通过 `/preload.js` 自动同步 Antigravity 宿主主题变量。

---

## 项目结构

```text
qy-office-viewer/
├── plugin.json                        # Antigravity 插件清单
├── CHANGELOG.md                       # 版本发布与变更记录（Release List）
├── README.md                          # 项目说明文档
├── sync.ps1                           # Windows 本地一键同步部署脚本
├── assets/
│   └── logo.svg                       # 插件与侧边栏矢量图标
└── sidecars/
    └── panel/                         # 侧边栏 Sidecar 服务与前端
        ├── sidecar.json               # Sidecar 入口、进程与 AUX_PANE 配置
        ├── package.json               # ES Module 声明
        ├── main.mjs                   # Node.js 后端（静态路由、文件读取、表格解析、系统动作）
        ├── index.html                 # 侧边栏页面骨架
        ├── app.js                     # 前端渲染调度、缩放与热重载逻辑
        ├── styles.css                 # 主题自适应与文档仿真排版样式
        └── vendor/                    # 纯离线前端渲染引擎（零外网依赖）
            ├── docx-preview.min.js
            ├── jszip.min.js
            ├── xlsx.full.min.js
            ├── pdf.min.js
            └── pdf.worker.min.js
```

---

## 安装与部署

插件通过 `%USERPROFILE%\.gemini\config\plugins\qy-office-viewer` 加载，提供两种部署方式：

### 方式一：NTFS Junction 软链接（推荐，开发调试零同步成本）

```powershell
New-Item -ItemType Junction `
  -Path "$env:USERPROFILE\.gemini\config\plugins\qy-office-viewer" `
  -Target "D:\Project\qy-office-viewer"
```

### 方式二：使用同步脚本复制部署

```powershell
pwsh -NoProfile -File .\sync.ps1
```

### 启用步骤

1. 打开 Antigravity 左侧 **Customizations** 面板；
2. 切换到 **Installed**，确认 **`qy-office-viewer`**（Office 预览）已启用；
3. 在对话中点击 `sidecar://qy-office-viewer/panel/?file=...` 文档胶囊，或直接将本地文档拖入侧边栏即可预览。

---

## 版本历史（Release List）

完整版本变更记录见 [CHANGELOG.md](./CHANGELOG.md) 及仓库 [Releases](https://github.com/cyt7732/qy-office-viewer/releases)。

| 版本 | 发布日期 | 核心亮点 |
| :--- | :--- | :--- |
| **[v1.1.0](https://github.com/cyt7732/qy-office-viewer/releases/tag/v1.1.0)** | 2026-11-22 | 新增 PDF.js 离线高清多页渲染、Node 服务端无损 CJK 表格解析、自适应心跳热重载、紧凑顶栏重构与安全脱敏 |
| **[v1.0.0](https://github.com/cyt7732/qy-office-viewer/releases/tag/v1.0.0)** | 2026-10-10 | 首个完整可用版本：Sidecar 胶囊路由、Word 离线仿真排版与自适应宽度、Excel 初版预览、系统程序非阻塞直开/定位/复制 |

---

## 后续演进规划（Roadmap）

- [ ] 修复 Word 彩色文本/红头标题被全局深色样式覆盖的问题，优化公文落款识别准确性
- [ ] 补齐 PDF.js 中文 CMap 资产、可视区懒渲染与 `TextLayer` 划词复制/全文搜索
- [ ] 接入本地 LibreOffice 无头转换缓存，解锁 `.pptx` / `.ppt` / `.doc` / `.wps` 全格式预览及 Word 精准打印视图切换
- [ ] 将文件变更监测从心跳轮询升级为 `fs.watch` + SSE 推送，并在热重载后保持阅读滚动位置
- [ ] 补全顶部最近/会话文档快速切换器，以及 `.md` / `.txt` / `.json` 纯文本预览

---

## 依赖许可

- `docx-preview` — Volter9 (MIT License)
- `JSZip` — Stuart Knightley (MIT License)
- `SheetJS / xlsx` — SheetJS (Apache 2.0 License)
- `PDF.js` — Mozilla Foundation (Apache 2.0 License)
- `Antigravity Sidecar SDK` — Google Antigravity Team
