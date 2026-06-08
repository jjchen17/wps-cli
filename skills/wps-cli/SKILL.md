---
name: wps-cli
description: "通过 WPS Office COM 自动化操作 Word/Excel/PPT/PDF 文档。76 条 CLI 命令，100% 格式保真，--json 输出适配 AI Agent。When you need to: create/edit Word documents (.docx), manipulate Excel spreadsheets (.xlsx) with formulas/charts/data-validation, build PowerPoint presentations (.pptx), process PDF files, batch convert office formats, fill Word templates with {{key}} placeholders, validate document formatting, or diagnose WPS COM issues. 当用户需要操作办公文档、填写模板、生成报表、处理PDF、诊断WPS环境时自动触发。"
---
# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)

# WPS CLI — AI Agent 办公文档操作

在终端里指挥 WPS Office 干活 — **真格式、真排版**。不模拟文件格式，直接驱动真实的 WPS 引擎。

## 工作决策树

```
用户请求操作办公文档？
├── Word (.docx/.doc/.wps)  → writer 子命令
│   ├── 查看/诊断     → writer view <file> [summary|issues|outline|annotated|stats]
│   ├── 模板填充     → writer merge template.docx --data '{...}'
│   ├── 查找替换     → writer replace <file> <old> <new>
│   ├── 表格操作     → writer table-get / table-insert
│   ├── 路径定位     → writer get <file> "<path>"
│   ├── 表单域/控件  → writer formfield-* / contentcontrol-*
│   └── 文档验证     → writer validate <file>
├── Excel (.xlsx/.xls/.csv)  → calc 子命令
│   ├── 查看/诊断     → calc view <file> [summary|issues|sheets]
│   ├── 单元格读写   → calc cell-get / cell-set / cell-range
│   ├── 公式设置     → calc cell-formula
│   ├── 图表创建     → calc chart-create
│   ├── 条件格式     → calc conditional-format-*
│   ├── 数据验证     → calc data-validation-*
│   ├── 迷你图       → calc sparkline-add
│   └── 排序         → calc sort
├── PPT (.pptx/.ppt)         → impress 子命令
│   ├── 查看/诊断     → impress view <file> [summary|issues|slides]
│   ├── 幻灯片管理   → impress slide-list / slide-add / slide-delete
│   ├── 文本/图片    → impress text-get / text-set / image-insert
│   └── 导出         → impress export-pdf
├── PDF (.pdf)               → pdf 子命令
│   ├── 合并/拆分    → pdf merge / split / extract-pages
│   ├── 水印         → pdf watermark
│   └── 信息         → pdf info
├── 格式转换                  → export 子命令
│   ├── 单文件       → export convert <file> <format>
│   └── 批量         → export batch "*.docx" -t pdf
├── 批量执行                  → batch 子命令
│   └── JSON数组     → wps batch -c '[...]'
└── 环境诊断                  → wps doctor / version
```

## 核心概念

### 三层架构

```
CLI 层 (typer) → Service 层 (业务逻辑) → COM Backend (win32com) → WPS Office 桌面应用
```

- **CLI 层**：参数解析 + JSON 输出格式化
- **Service 层**：业务逻辑，通过抽象接口调用后端
- **COM Backend**：封装 COM 细节，可替换为 LibreOffice / WebOffice

### SessionManager

管理 WPS 进程生命周期（启动/关闭/复用），线程安全（RLock），UUID 会话 ID。

### 路径定位语法

统一 1-based 路径，支持三种风格：
```
/section[1]/paragraph[3]          # Word：第1节第3段
/sheet["Sheet1"]/cell["C12"]      # Excel：Sheet1 的 C12
$Sheet1:A1                        # Excel 风格简写
/slide[1]/shape[2]                # PPT：第1张第2个形状
```

## 使用须知（AI Agent 必读）

1. **命令是同步阻塞的**：每次 COM 操作需数秒（启动 WPS + 打开文件），设计工作流时预留等待时间。
2. **文件扩展名白名单**：writer 仅接受 `.doc/.docx/.wps/.rtf/.txt/.html`；calc 仅接受 `.xls/.xlsx/.xlsm/.et/.csv`；impress 仅接受 `.ppt/.pptx/.pps/.dps`。
3. **公式安全限制**：`calc cell-formula` 禁止 SHELL/DDE/HYPERLINK/WEBSERVICE 等危险函数。
4. **路径限制**：不支持 UNC 路径和符号链接，仅本地绝对/相对路径。
5. **仅 Windows**：依赖 COM 和 win32com，WPS Office 2019+ 必须安装。
6. **全局选项**：所有命令支持 `--json` / `-j` 输出机器可读 JSON；`--help` 查看帮助。

## 诊断工作流

遇到问题时按此顺序排查：
1. `wps doctor` — 环境诊断（Python/pywin32/WPS 版本/COM 注册表）
2. `wps doctor --fix` — 自动修复 COM 注册问题（需管理员权限）
3. `wps doctor --report` — 生成脱敏报告粘贴到 GitHub Issue
4. `wps writer|calc|impress validate <file>` — 文档级验证

## 参考文档

以下参考文档按需加载，避免占用上下文窗口：

- [commands.md](references/commands.md) — 完整 76 条命令速查表
- [patterns.md](references/patterns.md) — 8 大常见使用模式（模板填充、报表生成、批量转换等）
- [mcp.md](references/mcp.md) — MCP 服务器配置与 27 个工具列表
- [json-schema.md](references/json-schema.md) — JSON 输出格式、错误 Schema、退出码语义
