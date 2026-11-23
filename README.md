# QY Office Viewer

> 基于 **Google Antigravity UI Extension & Sidecar SDK** 构建的侧边栏 Office 文档离线预览插件。  
> 专为智能体生成与处理本地 Office 文档的场景设计，实现「对话胶囊一键直达侧边栏预览 + 毫秒级唤起本地 Office/WPS 程序」的无缝衔接。

---

## 支持格式一览

| 格式 | 扩展名 | 侧边栏离线预览 | 唤起本地默认程序 / 定位 / 复制路径 |
| :--- | :--- | :---: | :---: |
| **Word 文档** | `.docx` | 支持（`docx-preview` 仿真纸张排版，保留红头与彩色样式） | 支持 |
| **Word 旧格式** | `.doc` | 暂不支持（提示一键本地打开） | 支持 |
| **Excel 表格** | `.xlsx` / `.xls` / `.csv` | 支持（服务端 SheetJS 解析 + 自动编码识别 + 表头吸顶 + 多 Sheet 切换） | 支持 |
| **PDF 文档** | `.pdf` | 支持（PDF.js 离线中文 CMap + 懒渲染 + 划词复制 + 清晰缩放） | 支持 |
| **文本与配置** | `.txt` / `.md` / `.json` | 支持（等宽字体渲染，JSON 自动美化缩进） | 支持 |
| **PowerPoint** | `.pptx` / `.ppt` | 暂不支持（提示一键本地打开） | 支持 |

---

## 核心特性

### 1. 纯离线多格式侧边栏预览
- **Word（`.docx`）仿真排版**：内置 `docx-preview` 与 `JSZip`，白底纸张流式呈现段落、红头/彩色文字、表格、页眉页脚与图片；默认横向自适应视口宽度，支持按钮及 `Ctrl/Cmd + 滚轮` 缩放与窄屏响应。
- **Excel（`.xlsx` / `.xls` / `.csv`）多工作表浏览**：内置完整版 SheetJS（`xlsx.full.min.js`），支持 BOM + UTF-8 / `GB18030` 自动解码、多 Sheet 标签切换、表头吸顶冻结、隔行高亮、长单元格自动换行；服务异常时自动降级前端本地解析。
- **PDF（`.pdf`）高保真懒渲染与文字选择**：集成 Mozilla PDF.js 与完整离线 CMap（168 份 `.bcmap`），支持 `IntersectionObserver` 可视区懒渲染、`TextLayer` 划词复制、缩放防抖清晰重绘与悬浮页码徽标。
- **纯文本与本地拖放直览**：支持 `.txt` / `.md` / `.json` 预览，以及将本地文件直接拖入侧边栏即时预览。

### 2. 系统本地程序无缝联动
- **一键本地打开**：非阻塞瞬时唤起系统默认程序（Windows `ProcessStartInfo` 唤起 Office / WPS；macOS `open`；Linux `xdg-open`）。
- **文件管理器高亮定位**：
  - Windows：直调 `explorer.exe /select,"<path>"` 高亮选中文件；
  - macOS：通过 `open -R` 在访达（Finder）中定位并高亮选中文件；
  - Linux：优先通过 Freedesktop D-Bus `org.freedesktop.FileManager1.ShowItems`（`gdbus` / `dbus-send`）在默认文件管理器中高亮选中文件，不支持时自动回退 `xdg-open` 打开所在目录。
- **路径复制**：优先使用浏览器剪贴板 API，后端按平台自动兜底（Windows `Set-Clipboard`；macOS 强制 UTF-8 环境变量的 `pbcopy`；Linux 按 Wayland/X11 会话自适应 `wl-copy` / `xclip` / `xsel`）。

### 3. Antigravity Sidecar 胶囊协议与智能热重载
- **标准胶囊协议**：在 Antigravity 对话中点击以下格式的 Markdown 胶囊即可在辅助侧边栏展开并定位对应文档：
  ```markdown
  [文件名.docx](sidecar://qy-office-viewer/panel/?file=<URI编码后的文件绝对路径>)
  ```
- **智能变更监听与保位热重载**：
  - 后端基于 `fs.watch` 监听文件改动并提供内存态 `/api/file/watch-state`，前端温和感知并在文档保存后自动热重载，同时保留滚动位置、工作表标签与缩放状态；
  - 精准路径校验：当目标文件不存在时不再静默替换，而是明确提示并展示同目录相似候选供用户主动点击打开。
- **按需加载与安全加固**：按文档类型懒加载离线渲染库；Windows 原生打开/剪贴板采用 PowerShell `-EncodedCommand` 参数安全唤起、文件定位直调 `explorer.exe` 原字面参数，macOS/Linux 采用无 shell 数组传参；静态/文件接口带扩展名白名单校验。

---

## 项目结构

```text
qy-office-viewer/
├── plugin.json                        # Antigravity 插件清单
├── CHANGELOG.md                       # 版本发布与变更记录（Release List）
├── README.md                          # 项目说明文档
├── sync.ps1                           # Windows 本地一键同步部署脚本
├── sync.sh                            # macOS / Linux 本地一键同步部署脚本
├── assets/
│   └── logo.svg                       # 插件与侧边栏矢量图标
└── sidecars/
    └── panel/                         # 侧边栏 Sidecar 服务与前端
        ├── sidecar.json               # Sidecar 入口、进程与 AUX_PANE 配置
        ├── package.json               # ES Module 声明
        ├── lib.mjs                    # 路径校验、CSV 编码识别与系统安全调用工具
        ├── main.mjs                   # Node.js 后端（缓存静态路由、fs.watch 监听、表格解析、系统动作）
        ├── index.html                 # 侧边栏页面骨架（按需懒加载引擎）
        ├── app.js                     # 前端渲染调度、PDF TextLayer 与热重载保位逻辑
        ├── styles.css                 # 主题自适应与文档仿真排版样式
        └── vendor/                    # 纯离线前端渲染引擎与 CMap（零外网依赖）
            ├── cmaps/                 # PDF.js 3.11.174 官方离线 CMap（168 份 .bcmap）
            ├── docx-preview.min.js
            ├── jszip.min.js
            ├── xlsx.full.min.js
            ├── pdf.min.js
            └── pdf.worker.min.js
```

---

## 安装与部署

插件通过 `~/.gemini/config/plugins/qy-office-viewer`（Windows 为 `%USERPROFILE%\.gemini\config\plugins\qy-office-viewer`）加载，提供软链接与脚本同步两种部署方式：

### 方式一：目录软链接（推荐，开发调试零同步成本）

- **Windows（PowerShell）**：
  ```powershell
  New-Item -ItemType Junction `
    -Path "$env:USERPROFILE\.gemini\config\plugins\qy-office-viewer" `
    -Target "D:\Project\qy-office-viewer"
  ```
- **macOS / Linux（Bash / Zsh）**：
  ```bash
  mkdir -p ~/.gemini/config/plugins
  ln -s /path/to/qy-office-viewer ~/.gemini/config/plugins/qy-office-viewer
  ```

### 方式二：使用同步脚本复制部署

- **Windows（PowerShell 7）**：
  ```powershell
  pwsh -NoProfile -File .\sync.ps1
  ```
- **macOS / Linux（Bash）**：
  ```bash
  chmod +x ./sync.sh && ./sync.sh
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
| **[v1.2.2](https://github.com/cyt7732/qy-office-viewer/releases/tag/v1.2.2)** | 2026-11-23 | 补齐 macOS（访达高亮、UTF-8 `pbcopy`）与 Linux（D-Bus `FileManager1` 文件高亮回退链、Wayland/X11 剪贴板防挂起）跨平台能力，新增 `sync.sh` 与三端部署文档 |
| **[v1.2.1](https://github.com/cyt7732/qy-office-viewer/releases/tag/v1.2.1)** | 2026-11-23 | 修复 Windows 无法唤起 WPS/默认程序与资源管理器定位的回归问题，定位改为直调 `explorer.exe` 原字面参数并修复 Store 版 `pwsh` 识别与失败自动回退 |
| **[v1.2.0](https://github.com/cyt7732/qy-office-viewer/releases/tag/v1.2.0)** | 2026-11-22 | 修复 Word 红头涂黑、静默误匹配、Excel 吸顶与 CSV 编码/拖拽、PDF CMap/懒渲染/划词复制、Windows 特殊字符命令注入、`fs.watch` 保位热重载与按需懒加载 |
| **[v1.1.0](https://github.com/cyt7732/qy-office-viewer/releases/tag/v1.1.0)** | 2026-11-22 | 新增 PDF.js 离线高清多页渲染、Node 服务端无损 CJK 表格解析、自适应心跳热重载、紧凑顶栏重构与安全脱敏 |
| **[v1.0.0](https://github.com/cyt7732/qy-office-viewer/releases/tag/v1.0.0)** | 2026-10-10 | 首个完整可用版本：Sidecar 胶囊路由、Word 离线仿真排版与自适应宽度、Excel 初版预览、系统程序非阻塞直开/定位/复制 |

---

## 后续演进规划（Roadmap）

- [x] 修复 Word 彩色文本/红头标题被全局深色样式覆盖的问题，优化公文落款识别准确性
- [x] 补齐 PDF.js 中文 CMap 资产、可视区懒渲染与 `TextLayer` 划词复制
- [x] 将文件变更监测从高频轮询升级为 `fs.watch` 后端监听，并在热重载后保持阅读滚动位置
- [x] 补全 `.md` / `.txt` / `.json` 纯文本预览
- [ ] 接入本地 LibreOffice 无头转换缓存，解锁 `.pptx` / `.ppt` / `.doc` / `.wps` 全格式预览及 Word 精准打印视图切换
- [ ] 增加 PDF 页内全文搜索（`Ctrl+F`）、页码直达跳转，以及 Excel 表内搜索与顶部会话文档快速切换器

---

## 依赖许可

- `docx-preview` — Volter9 (MIT License)
- `JSZip` — Stuart Knightley (MIT License)
- `SheetJS / xlsx` — SheetJS (Apache 2.0 License)
- `PDF.js` — Mozilla Foundation (Apache 2.0 License)
- `Antigravity Sidecar SDK` — Google Antigravity Team
