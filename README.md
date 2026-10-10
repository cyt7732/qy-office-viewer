# QY Office Viewer (千羽 Office 文档全能预览与系统直开助手)

> 基于 **Google Antigravity 2.22.0+ UI Extension & Sidecar SDK** 规范打造的侧边栏文档中枢插件。  
> 专为解决 AI 智能体生成 Office 二进制工件（Word / Excel / PDF / PPT）后无法直接调用本地程序与难以快速预览的痛点而生。

---

## 🌟 核心特性 (Key Features)

### 1. 🚀 系统原生操作中枢【P0 级核心】
- **一键调用系统默认程序直开**：
  - 面板顶部常驻醒目按钮：**`在 Word / WPS / Excel 中打开`**；
  - 后台 Node.js 原生调用 Win32 / macOS 指令，毫秒级唤起本地已安装的 Microsoft Office、WPS Office 或默认阅读器，绝无弹窗阻塞与卡顿假死。
- **一键在系统文件管理器中高亮定位**：
  - 点击 **`定位文件`** 按钮，立即在 Windows 资源管理器（或 macOS 访达）中打开目标文件夹，并**精准高亮选中**该文件。
- **一键复制物理绝对路径**：
  - 快速复制文件完整路径至剪贴板，方便在终端、代码或沟通中引用。
- **一键联动 AI 对话润色**：
  - 点击 **`AI 润色`**，自动将当前预览的文件路径与修改诉求推回当前会话，由智能体继续精进优化。

### 2. 📖 侧边栏多格式仿真直接阅览【P1 级核心】
- **Word (`.docx`) 仿真排版**：
  - 纯前端离线集成 `docx-preview` 引擎，以仿 A4 实体纸张形态直接渲染标题层级、字体、颜色、段落缩进与表格样式；
  - 附带 50% ~ 200% 平滑缩放控制器与一键还原。
- **Excel (`.xlsx` / `.xls` / `.csv`) 交互式表格**：
  - 支持多工作表（Sheet 1, Sheet 2...）标签页秒级切换；
  - 表头吸顶置顶、交替斑马纹网格与响应式滚动浏览。
- **PDF (`.pdf`) 极速阅读**：
  - 直接挂载 Chromium 原生 PDF 阅读内核，畅享高保真矢量排版与翻页体验。
- **代码与纯文本 (`.txt`, `.md`, `.json`)**：
  - 等宽字体清爽阅读。

### 3. 🔍 智能工件自动嗅探 (Artifact Sniffer)
- 侧边栏展开或切换会话时，自动识别当前会话工件目录（`brain/<id>`）与当前活跃工作区最新生成的 Office 文档；
- 自动按修改时间倒序排列并**默认载入最新文件**，实现“生成完毕打开侧边栏即见成果”的丝滑体验；
- 提供手动路径输入栏与全局拖拽（Drag & Drop）直览。

---

## 🛠️ 项目结构 (Project Structure)

```text
qy-office-viewer/
├── plugin.json                 # Antigravity 插件根清单
├── assets/
│   └── logo.svg                # 官方侧边栏矢量徽标
├── sidecars/
│   └── panel/                  # 辅助侧面板扩展核心
│       ├── sidecar.json        # 侧面板入口与进程配置
│       ├── package.json        # ES Module 规范声明
│       ├── main.mjs            # Node.js 后台服务（原生指令、文件嗅探、流式路由）
│       ├── index.html          # 侧边栏前端骨架
│       ├── app.js              # 前端交互与渲染调度
│       ├── styles.css          # Antigravity 主题自适应样式表
│       └── vendor/             # 纯离线前端渲染依赖
│           ├── docx-preview.min.js
│           ├── jszip.min.js
│           └── xlsx.mini.min.js
├── .gitignore                  # Git 忽略配置
└── README.md                   # 项目说明文档
```

---

## 🚀 安装与调试指引 (Installation & Setup)

Antigravity 插件通过 `%USERPROFILE%\.gemini\config\plugins\` 发现。  
推荐使用 Windows NTFS Junction（软链接）进行双向热联开发：

### 1. 建立软链接（以管理员身份或普通 PowerShell 运行）
```powershell
New-Item -ItemType Junction -Path "$env:USERPROFILE\.gemini\config\plugins\qy-office-viewer" -Target "D:\Project\qy-office-viewer"
```

### 2. 在 Antigravity 中启用
1. 打开 Antigravity 界面左侧导航栏的 **Customizations**（自定义）面板；
2. 切换到 **Installed**（已安装）标签页；
3. 找到 **`qy-office-viewer`** 并将其状态设为 **Enabled**；
4. 随后在任意对话界面右上角点击 **⋮**（三点菜单），在 **Extensions** 中点击 **`Office 预览与直开`** 即可在右侧面板打开！

---

## 📄 依赖许可 (Licenses & Acknowledgments)
- `docx-preview`: Volter9 (MIT License)
- `JSZip`: Stuart Knightley (MIT License)
- `SheetJS / xlsx`: Apache 2.0 License
- Antigravity Sidecar SDK: Google Antigravity Team
