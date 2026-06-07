# WPS CLI Skills

> AI Agent 教学文件：让 Claude Code、Cursor 等 AI 工具理解并正确使用 wps-cli 的全部能力。
> 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

## 概述

**wps-cli** 是一个通过 COM 自动化驱动 WPS Office 桌面端的命令行工具。它允许 AI Agent 在 Windows 环境下以编程方式操作 Word、Excel、PowerPoint 和 PDF 文档，无需人工交互。

### 核心能力

- **Word (Writer)**：创建、读取、修改、查找替换、表格操作、图片插入、页面布局、导出 PDF
- **Excel (Calc)**：单元格读写、公式设置、区域操作、图表创建、排序筛选、工作表管理
- **PowerPoint (Impress)**：幻灯片管理、文本读写、图片插入、备注管理、导出 PDF
- **PDF**：元信息读取、合并、拆分、提取页面、水印添加
- **格式转换**：Word/Excel/PPT 之间的格式互转（含 PDF）

### 限制

| 限制项 | 说明 |
|--------|------|
| 平台 | **仅 Windows**（依赖 COM 和 win32com） |
| WPS 版本 | 需要 WPS Office 2019 或更高版本 |
| 并发 | 单进程模式，同一会话对象不应跨线程并发使用 |
| 网络 | 无需网络连接，所有操作在本地完成 |

## 架构

### 三层架构

```
CLI 层 (typer)
    │
    ▼
Service 层 (业务逻辑)
    │
    ▼
COM Backend (win32com)
    │
    ▼
WPS Office 桌面应用
```

### 关键组件

- **SessionManager**：管理 WPS 进程生命周期（启动/关闭/复用）
- **WriterService**：Word 文档的所有操作逻辑
- **CalcService**：Excel 工作簿的所有操作逻辑（含公式安全防护）
- **ImpressService**：PPT 演示文稿的所有操作逻辑
- **PdfService**：PDF 处理逻辑（基于 Writer 的 PDF 导出能力）
- **ExportService**：跨应用格式转换
- **StyleEngine**：预设样式管理（公文格式、报告格式等）

### 安全设计

- **COM 宏自动执行禁用**：`AutomationSecurity = msoAutomationSecurityForceDisable`，防止恶意宏
- **公式注入防护**：`CalcService.cell_formula` 禁止 `SHELL()`, `DDE()`, `HYPERLINK()`, `WEBSERVICE()` 等危险函数
- **路径遍历防护**：所有输入路径经过符号链接、UNC、目录遍历检查
- **长度限制**：文本替换上限 1000 字符，水印文字上限 100 字符，glob 结果上限 200

## 命令速查

> 所有命令均支持 `--json` / `-j` 标志输出 JSON 格式，便于 AI 解析。
> 全局选项：`--help` 查看帮助，`--json` 输出 JSON。

### 环境诊断

```bash
wps doctor              # 诊断环境：检查 Python、pywin32、WPS 组件
wps doctor --report     # 输出脱敏 Markdown 诊断报告（适合粘贴到 GitHub Issue）
wps version             # 输出版本号
```

### Word 操作 (writer)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps writer info <file>` | 文档元信息（页数/字数/字符数/段落数/作者） | `wps writer info report.docx --json` |
| `wps writer count <file>` | 字数统计 | `wps writer count report.docx --json` |
| `wps writer replace <file> <old> <new>` | 查找替换文本 | `wps writer replace doc.docx "旧" "新"` |
| `wps writer replace <file> <old> <new> -w` | 通配符替换 | `wps writer replace doc.docx "张?" "李?" -w` |
| `wps writer table-get <file> -i 1` | 读取第 1 个表格 | `wps writer table-get data.docx -i 1 --json` |
| `wps writer table-insert <file> -r 3 -c 4 -d '[["A","B"]]'` | 插入 3x4 表格 | `wps writer table-insert doc.docx -r 3 -c 4 -d '[[...]]'` |
| `wps writer image-insert <file> -i <image>` | 插入图片 | `wps writer image-insert doc.docx -i photo.png` |
| `wps writer image-insert <file> -i <image> -w 200 -h 150` | 插入图片并指定尺寸 | `wps writer image-insert doc.docx -i photo.png -w 200 -h 150` |
| `wps writer page-setup <file>` | 页面布局设置 | `wps writer page-setup doc.docx --width 210 --height 297` |
| `wps writer export-pdf <file> -o <output>` | 导出 PDF | `wps writer export-pdf doc.docx -o output.pdf` |
| `wps writer style-apply <file> <preset>` | 应用预设样式 | `wps writer style-apply doc.docx "公文正文"` |
| `wps writer style-apply <file> "" -l` | 列出可用样式预设 | `wps writer style-apply doc.docx "" -l` |
| `wps writer new -o <path>` | 新建空白文档 | `wps writer new -o blank.docx` |

**可用样式预设**：公文标题、公文一级标题、公文二级标题、公文正文、报告标题、报告正文

### Excel 操作 (calc)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps calc info <file>` | 工作簿元信息 | `wps calc info data.xlsx --json` |
| `wps calc sheet-list <file>` | 列出所有工作表 | `wps calc sheet-list data.xlsx --json` |
| `wps calc cell-get <file> <ref>` | 读取单元格值 | `wps calc cell-get data.xlsx A1 --json` |
| `wps calc cell-get <file> <ref> -s <sheet>` | 指定工作表读取 | `wps calc cell-get data.xlsx B3 -s Sheet2 --json` |
| `wps calc cell-set <file> <ref> <value>` | 设置单元格值 | `wps calc cell-set data.xlsx A1 "Hello"` |
| `wps calc cell-range <file> <ref>` | 读取区域 | `wps calc cell-range data.xlsx A1:D10 --json` |
| `wps calc cell-formula <file> <ref> <formula>` | 设置公式 | `wps calc cell-formula data.xlsx B10 "=SUM(B1:B9)"` |
| `wps calc chart-create <file> -d <range> -t <type>` | 创建图表 | `wps calc chart-create data.xlsx -d A1:C10 -t pie --title "销售占比"` |
| `wps calc sort <file> -b <col> --order <asc/desc>` | 排序 | `wps calc sort data.xlsx -b A --order desc` |
| `wps calc export-csv <file> -o <output>` | 导出 CSV | `wps calc export-csv data.xlsx -o data.csv` |
| `wps calc new -o <path>` | 新建工作簿 | `wps calc new -o blank.xlsx` |

**图表类型**：`bar`（柱状图）、`line`（折线图）、`pie`（饼图）、`scatter`（散点图）、`area`（面积图）

### PPT 操作 (impress)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps impress info <file>` | 演示文稿元信息 | `wps impress info slides.pptx --json` |
| `wps impress slide-list <file>` | 列出所有幻灯片 | `wps impress slide-list slides.pptx --json` |
| `wps impress slide-add <file> -l 1 -t "标题"` | 新增幻灯片 | `wps impress slide-add slides.pptx -l 1 -t "新页面"` |
| `wps impress slide-delete <file> <index>` | 删除幻灯片 | `wps impress slide-delete slides.pptx 3` |
| `wps impress text-get <file> -s <idx>` | 获取幻灯片文本 | `wps impress text-get slides.pptx -s 2 --json` |
| `wps impress text-set <file> -s <idx> -p title -t "标题"` | 设置占位符文本 | `wps impress text-set slides.pptx -s 1 -p title -t "新标题"` |
| `wps impress image-insert <file> -s <idx> -i <image>` | 插入图片 | `wps impress image-insert slides.pptx -s 2 -i logo.png` |
| `wps impress export-pdf <file> -o <output>` | 导出 PDF | `wps impress export-pdf slides.pptx -o slides.pdf` |
| `wps impress new -o <path>` | 新建演示文稿 | `wps impress new -o blank.pptx` |

**占位符类型**：`title`（标题）、`body`（正文）、`subtitle`（副标题）

### PDF 操作 (pdf)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps pdf info <file>` | PDF 元信息 | `wps pdf info document.pdf --json` |
| `wps pdf merge <file1> <file2> ... -o <output>` | 合并 PDF | `wps pdf merge a.pdf b.pdf c.pdf -o merged.pdf` |
| `wps pdf extract-pages <file> <pages> -o <output>` | 提取页面 | `wps pdf extract-pages doc.pdf "1-3,5,7-9" -o extracted.pdf` |
| `wps pdf split <file> -e <N> -d <dir>` | 每 N 页拆分 | `wps pdf split doc.pdf -e 10 -d ./split/` |
| `wps pdf watermark <file> <text> -o <output>` | 添加水印 | `wps pdf watermark doc.pdf "机密" -o watermarked.pdf` |

### 格式转换 (export)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps export convert <file> <format> -o <output>` | 单文件格式转换 | `wps export convert doc.docx pdf -o doc.pdf` |
| `wps export batch <glob> -t <format> -d <dir>` | 批量转换 | `wps export batch "*.docx" -t pdf -d ./pdfs/` |

**支持的转换方向**：
- Writer → pdf, docx, doc, rtf, txt, html
- Calc → xlsx, csv
- Impress → pptx, ppt, pdf

### MCP 服务器 (mcp)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps mcp serve` | 启动 MCP stdio 服务器 | `wps mcp serve` |
| `wps mcp install -t claude` | 安装 MCP 到 Claude Code | `wps mcp install -t claude` |
| `wps mcp status` | 检查 MCP 注册状态 | `wps mcp status --json` |

### AI 工具集成 (install)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps install skill -t claude` | 安装 SKILL.md 到 Claude Code | `wps install skill -t claude` |
| `wps install mcp -t claude` | 安装 MCP 配置到 Claude Code | `wps install mcp -t claude` |
| `wps install all-tools -t claude` | 一键安装 SKILL.md + MCP | `wps install all-tools -t claude` |

## MCP 工具列表

通过 MCP 服务器暴露给 AI Agent 的工具（JSON-RPC 2.0 over stdio）：

| 工具名 | 功能 | 关键参数 |
|--------|------|---------|
| `writer_info` | Word 文档元信息 | file |
| `writer_replace` | 查找替换 | file, old_text, new_text, wildcard, case_sensitive |
| `writer_count` | 字数统计 | file |
| `writer_table_get` | 读取表格 | file, index |
| `writer_table_insert` | 插入表格 | file, rows, cols, data_json |
| `writer_export_pdf` | 导出 PDF | file, output |
| `writer_image_insert` | 插入图片 | file, image, width, height |
| `writer_page_setup` | 页面布局 | file, width_mm, height_mm, margin_* |
| `calc_info` | 工作簿元信息 | file |
| `calc_cell_get` | 读取单元格 | file, ref, sheet |
| `calc_cell_set` | 设置单元格 | file, ref, value, sheet |
| `calc_range_get` | 读取区域 | file, ref, sheet |
| `calc_cell_formula` | 设置公式 | file, ref, formula, sheet |
| `calc_chart_create` | 创建图表 | file, data_range, chart_type, title, sheet |
| `calc_sheet_list` | 列出工作表 | file |
| `impress_info` | PPT 元信息 | file |
| `impress_slide_list` | 列出幻灯片 | file |
| `impress_text_get` | 读取幻灯片文本 | file, slide_idx |
| `impress_text_set` | 设置幻灯片文本 | file, slide_idx, placeholder, text |
| `impress_export_pdf` | PPT 导出 PDF | file, output |
| `pdf_info` | PDF 元信息 | file |
| `pdf_merge` | 合并 PDF | files_json, output |
| `pdf_extract_pages` | 提取页面 | file, pages, output |
| `pdf_watermark` | 添加水印 | file, text, output |
| `pdf_split` | 拆分 PDF | file, every, output_dir |
| `export_convert` | 格式转换 | file, output_format, output |

## 常见模式

### 模式 1：模板填充

```bash
# 1. 打开模板文档
# 2. 替换占位符
wps writer replace template.docx "{{name}}" "张三"
wps writer replace template.docx "{{date}}" "2026-01-15"
# 3. 导出 PDF
wps writer export-pdf template.docx -o report.pdf
```

### 模式 2：数据报表生成

```bash
# 1. 创建或打开 Excel 工作簿
# 2. 写入表头和公式
wps calc cell-set report.xlsx A1 "项目"
wps calc cell-set report.xlsx B1 "金额"
wps calc cell-formula report.xlsx B10 "=SUM(B2:B9)"
# 3. 创建图表
wps calc chart-create report.xlsx -d A1:B9 -t bar --title "财务报表"
# 4. 导出 CSV 用于进一步处理
wps calc export-csv report.xlsx -o report.csv
```

### 模式 3：批量格式转换

```bash
# 批量将 docx 转为 pdf
wps export batch "reports/*.docx" -t pdf -d reports/pdf/
```

### 模式 4：PPT 内容维护

```bash
# 1. 查看幻灯片结构
wps impress slide-list deck.pptx --json
# 2. 修改特定幻灯片的文本
wps impress text-set deck.pptx -s 1 -p title -t "年度总结"
wps impress text-set deck.pptx -s 1 -p body -t "2026年工作回顾..."
# 3. 导出 PDF 用于分发
wps impress export-pdf deck.pptx -o deck.pdf
```

### 模式 5：PDF 处理流水线

```bash
# 1. 从不同来源获取 PDF
# 2. 提取需要的页面
wps pdf extract-pages source.pdf "1-3,7" -o extracted.pdf
# 3. 添加水印
wps pdf watermark extracted.pdf "内部资料" -o final.pdf
# 4. 合并到最终文档
wps pdf merge final.pdf appendix.pdf -o complete.pdf
```

## JSON 输出格式

所有命令使用 `--json` 标志后输出统一 JSON 结构：

**成功响应：**
```json
{
  "success": true,
  "command": "writer.info",
  "data": {
    "path": "report.docx",
    "pages": 10,
    "words": 2500,
    "characters": 3600,
    "paragraphs": 45,
    "author": "张三",
    "created": "2026-01-15",
    "modified": "2026-01-20"
  }
}
```

**错误响应：**
```json
{
  "success": false,
  "command": "writer.info",
  "error": {
    "type": "FileNotFoundErrorCli",
    "message": "文件不存在: ~/report.docx",
    "code": 21,
    "suggestion": "请检查文件路径是否正确",
    "context": {"path": "~/report.docx"}
  }
}
```

**退出码语义：**

| 退出码 | 含义 |
|--------|------|
| 0 | 成功 |
| 1 | 通用错误 |
| 10 | WPS 未安装/未启动 |
| 11 | 会话管理失败 |
| 20 | 文件操作失败（路径、权限） |
| 21 | 文件未找到 |
| 30 | COM 调用失败 |
| 40 | 不支持的格式 |
| 50 | 参数校验失败 |
| 60 | 操作超时 |
| 61 | 部分批量操作失败 |

## 注意事项

### AI Agent 使用须知

1. **命令执行是同步的**：每个 COM 操作可能需要数秒（启动 WPS + 打开文件），请耐心等待。
2. **文件路径必须存在**：所有输入文件需要提前确认存在，否则会得到退出码 21 错误。
3. **扩展名限制**：`writer` 命令仅接受 `.doc/.docx/.wps/.rtf/.txt/.html` 文件；`calc` 命令仅接受 `.xls/.xlsx/.xlsm/.et/.csv` 文件；`impress` 命令仅接受 `.ppt/.pptx/.pps/.dps` 文件。
4. **通配符替换计数**：使用 `-w` 通配符模式时，替换计数返回 -1（不确定），因为同一位置可能被不同通配规则反复匹配。
5. **公式安全限制**：`calc cell-formula` 不允许使用 SHELL()、DDE()、HYPERLINK()、WEBSERVICE() 等危险函数，这是为防止恶意公式注入。
6. **UNC 路径不支持**：所有文件路径必须是本地路径，不支持 `\\server\share` 格式。
7. **符号链接不支持**：出于安全考虑，不支持通过符号链接访问文件。
8. **MCP 环境下**：通过 MCP 调用时，所有工具调用会在独立的 WPS 会话中完成，自动管理生命周期。

### WPS 版本兼容

- WPS 2019 及以上版本全面兼容
- 部分高级功能（如特定图表类型）可能需要 WPS 2021+
- 不同 WPS 版本对 COM 接口的实现略有差异，遇到问题请运行 `wps doctor --report` 生成诊断报告

### 故障排查优先级

1. 运行 `wps doctor` 确认环境正常
2. 确认文件路径正确且文件未被其他程序占用
3. 确认 WPS 已安装且能正常启动
4. 检查文件扩展名是否在支持列表中
5. 对于批量操作失败，检查单文件操作是否正常

## MCP 配置指南

### Claude Code

在 `~/.claude/mcp.json` 中添加：
```json
{
  "wps-cli": {
    "command": "wps",
    "args": ["mcp", "serve"]
  }
}
```

或运行 `wps mcp install -t claude` 自动配置。

### Cursor

在 `.cursor/mcp.json` 中添加相同配置，或运行 `wps mcp install -t cursor`。

### 验证安装

运行 `wps mcp status` 检查各 AI 工具的配置状态。
