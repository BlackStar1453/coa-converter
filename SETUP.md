# COA Converter Web — 安装与使用指南

将供应商 COA (Certificate of Analysis) PDF 中的检测数据提取并填充到 XLSX/DOCX 模板中。提供 Web 界面（局域网可访问）和 Claude Code Skill 两种使用方式。

## 快速开始

### 1. 克隆仓库

```bash
git clone https://github.com/BlackStar1453/coa-converter.git
cd coa-converter
```

> 可以 clone 到任意目录，项目使用相对路径，不依赖固定安装位置。

### 2. 一键安装

```bash
bash install.sh
```

安装脚本会自动完成：
- 安装 Homebrew（如未安装，仅 macOS）
- 安装 Python 3（如未安装，通过 Homebrew / apt）
- 创建虚拟环境（`.venv/`）并安装依赖
- 创建 `input/` 和 `output/` 工作目录
- 安装 Claude Code Skills（`/coa-to-template` + `/coa-fix-output`）到 `~/.claude/commands/`

> 全新 Mac 也能直接运行，无需提前安装任何工具。安装过程中可能需要输入电脑密码。

### 3. 放置模板文件

模板文件已包含在仓库的 `templates/` 目录中（10 个模板：5 XLSX + 5 DOCX）。如需更新模板，直接替换该目录下的文件即可。

### 4. 启动 Web 服务

```bash
bash run.sh
```

终端会显示两个地址：
- **本机访问**：`http://127.0.0.1:5050`
- **局域网访问**：`http://<你的IP>:5050`（同一网络下的其他设备均可访问）

## 使用方式

### 方式一：Web 界面

1. 启动服务后，在浏览器打开上述地址
2. 上传 COA PDF 文件
3. 选择目标模板
4. 点击转换，完成后下载结果文件

### 方式二：Claude Code Skill

在 Claude Code 中直接使用：

```
/coa-to-template <pdf路径> <模板路径> [输出路径]
```

示例：
```
/coa-to-template ./input/sample.pdf ./templates/Key\ In\ COA\ -\ Assay.xlsx
```

Skill 会自动完成：PDF 数据提取 → 模板填充 → AI 验证 → 错误修复 → 供应商注册。

如果转换结果有误，使用修复 Skill：
```
/coa-fix-output <错误描述>
```

示例：
```
/coa-fix-output Lot No. 填到了 Mfg. Date 的位置
```

## 项目结构

```
coa-converter-web/
├── converter/              # 核心转换引擎
│   ├── coa_converter.py    #   PDF 解析 + 数据提取
│   ├── template_detector.py #   模板布局自动检测
│   ├── xlsx_filler.py      #   XLSX 模板填充
│   ├── docx_filler.py      #   DOCX 模板填充
│   ├── supplier_checker.py #   供应商识别
│   └── supplier_registry.json # 已验证供应商注册表
├── static/                 # Web 前端（HTML/CSS/JS）
├── templates/              # 模板文件（XLSX + DOCX）
├── .claude/commands/       # Claude Code Skill 定义
├── app.py                  # Web 服务器
├── converter_service.py    # 转换服务封装
├── job_manager.py          # 任务队列管理
├── terminal_launcher.py    # Claude Code 验证启动器
├── install.sh              # 一键安装脚本
├── run.sh                  # 启动脚本
├── requirements.txt        # Python 依赖
├── input/                  # 待转换的 PDF（运行时创建，不入库）
└── output/                 # 转换结果输出（运行时创建，不入库）
```

## 环境要求

- macOS / Linux（Windows 需使用 PowerShell）
- Claude Code（仅使用 Skill 方式时需要）
- Python 3 和 Homebrew 会在安装时自动安装（如缺失）

## 常见问题

**Q: 局域网其他设备无法访问？**
- 确认设备在同一网络下
- 检查 macOS 防火墙设置，允许 Python 的入站连接
- 在"系统设置 → 网络 → 防火墙"中添加例外

**Q: 转换报错找不到模块？**
- 运行 `bash install.sh` 重新安装依赖
- 或手动安装：`pip install pdfplumber openpyxl PyMuPDF python-docx`

**Q: 如何更新模板？**
- 直接替换 `templates/` 目录下的文件，无需重启服务
