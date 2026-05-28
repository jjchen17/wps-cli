# WPS CLI

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/python-%3E%3D3.10-blue.svg)](https://python.org)
[![Platform](https://img.shields.io/badge/platform-Windows-blue.svg)](https://www.microsoft.com/windows)

[English](#english) | 中文

---

## 中文

WPS Office 全量 CLI 工具 — 通过 COM 自动化驱动 WPS 桌面端，覆盖 Word/Excel/PPT/PDF。

> 格式与手动操作完全一致，不是模拟文件格式，而是驱动真实的 WPS 程序。

### 特点

- **格式 100% 保真** — 通过 COM 自动化驱动真实 WPS 程序，不是模拟文件格式
- **覆盖四大应用** — Word、Excel、PPT、PDF 全部支持
- **AI Agent 友好** — 所有命令支持 `--json` 输出
- **三层架构** — CLI / Service / Backend 解耦，可扩展

### 安装

**系统要求：**
- Windows 10/11
- WPS Office 2019+
- Python 3.10+

**安装命令：**

```bash
pip install wps-cli
```

**从源码安装：**

```bash
git clone https://github.com/jjchen17/wps-cli.git
cd wps-cli
pip install -e .
```

**验证安装：**

```bash
wps doctor
```

### 快速开始

```bash
# 新建 Word 文档
wps writer new -o report.docx

# 替换文本
wps writer replace report.docx "旧文本" "新文本" --json

# 读取 Excel 单元格
wps calc cell get budget.xlsx "B3" --json

# 创建图表
wps calc chart create budget.xlsx --data "A1:C10" --type bar --title "销售趋势"

# 列出 PPT 幻灯片
wps impress slide list deck.pptx --json

# 合并 PDF
wps pdf merge cover.pdf body.pdf --output final.pdf

# 格式转换
wps export convert report.docx pdf
wps export batch "*.docx" --to pdf --output-dir ./pdf/
```

### 命令体系

```
wps
├── writer          Word 文档操作
│   ├── new / info
│   ├── replace / count
│   ├── table_insert / table_get
│   ├── image_insert
│   ├── page_setup
│   └── export_pdf
├── calc            Excel 表格操作
│   ├── new / info
│   ├── sheet_list
│   ├── cell_get / cell_set / cell_range / cell_formula
│   ├── chart_create
│   ├── sort
│   └── export_csv
├── impress         PPT 演示操作
│   ├── new / info
│   ├── slide_list / slide_add / slide_delete
│   ├── text_set / text_get
│   ├── image_insert
│   └── export_pdf
├── pdf             PDF 处理
│   ├── info
│   ├── merge / extract_pages / split
│   └── watermark
├── export          格式转换
│   ├── convert
│   └── batch
├── version         版本信息
└── doctor          环境诊断
```

### 全局选项

| 选项 | 说明 |
|------|------|
| `--json` | JSON 输出（AI Agent 友好） |
| `--tsv` | TSV 输出（管道友好） |
| `--quiet` | 静默模式 |
| `--dry-run` | 预览模式 |

### 架构

```
CLI 层 (Typer)
    │
    ▼
业务层 (Service)
    │
    ▼
COM 后端层 (Backend)
    │
    ▼
WPS Office 桌面端
```

三层解耦：CLI 层只做参数解析和输出格式化，业务层通过抽象接口调用后端，后端层封装 COM 细节。

### 开发

```bash
git clone https://github.com/jjchen17/wps-cli.git
cd wps-cli
pip install -e ".[dev]"
pytest
```

### 反馈

欢迎提交 Issue 或发送邮件至 **948881912@qq.com**。

### 许可证

[MIT](LICENSE)

---

## English

A full-featured CLI tool for WPS Office — drives WPS desktop applications via COM automation, covering Word/Excel/PPT/PDF.

> Formatting is identical to manual operation. This is not format simulation — it drives the real WPS application.

### Why

### Installation

**Requirements:**
- Windows 10/11
- WPS Office 2019+
- Python 3.10+

```bash
pip install wps-cli
```

### Quick Start

```bash
# Create Word document
wps writer new -o report.docx

# Replace text
wps writer replace report.docx "old" "new" --json

# Read Excel cell
wps calc cell get budget.xlsx "B3" --json

# Create chart
wps calc chart create budget.xlsx --data "A1:C10" --type bar --title "Sales"

# List PPT slides
wps impress slide list deck.pptx --json

# Merge PDFs
wps pdf merge cover.pdf body.pdf --output final.pdf

# Format conversion
wps export convert report.docx pdf
```

### Architecture

```
CLI Layer (Typer)
    │
    ▼
Service Layer
    │
    ▼
COM Backend Layer
    │
    ▼
WPS Office Desktop
```

### License

[MIT](LICENSE)
