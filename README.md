# WPS CLI

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/python-%3E%3D3.10-blue.svg)](https://python.org)
[![Platform](https://img.shields.io/badge/platform-Windows-blue.svg)](https://www.microsoft.com/windows)

WPS Office 全量 CLI 工具 — 通过 COM 自动化驱动 WPS 桌面端，覆盖 Word/Excel/PPT/PDF。

> 格式与手动操作完全一致，不是模拟文件格式，而是驱动真实的 WPS 程序。

[安装](#安装) · [快速开始](#快速开始) · [命令体系](#命令体系) · [架构](#架构) · [开发](#开发)

## 为什么需要它

| 方案 | 格式保真度 | 功能覆盖 |
|------|-----------|---------|
| python-docx / openpyxl | 低，很多格式不支持 | 仅文件操作 |
| cli-anything-wps | 100%（COM 驱动） | 47 个命令 |
| **wps-cli（本项目）** | **100%（COM 驱动）** | **146+ 个命令** |

## 安装

### 系统要求

- Windows 10/11
- WPS Office 2019+
- Python 3.10+

### 安装命令

```bash
pip install wps-cli
```

### 从源码安装

```bash
git clone https://github.com/yourname/wps-cli.git
cd wps-cli
pip install -e .
```

### 验证

```bash
wps doctor
```

## 快速开始

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

## 命令体系

```
wps
├── writer          Word 文档
│   ├── new / info
│   ├── replace / count
│   ├── table_insert / table_get
│   ├── image_insert
│   ├── page_setup
│   └── export_pdf
├── calc            Excel 表格
│   ├── new / info
│   ├── sheet_list
│   ├── cell_get / cell_set / cell_range / cell_formula
│   ├── chart_create
│   ├── sort
│   └── export_csv
├── impress         PPT 演示
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

## 架构

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

**三层解耦**：
- CLI 层只做参数解析和输出格式化
- 业务层通过抽象接口调用后端
- 后端层封装 COM 细节，可替换为 LibreOffice / WebOffice

## 与 cli-anything-wps 对比

| 维度 | cli-anything-wps | wps-cli |
|------|-----------------|---------|
| 命令数量 | 47 | 146+ |
| 页眉页脚 | 无 | 完整支持 |
| 目录 | 无 | 插入/更新 |
| 批注修订 | 无 | 完整支持 |
| 数据透视表 | 无 | 支持 |
| 条件格式 | 无 | 支持 |
| PDF 工具 | 无 | 15+ 命令 |
| 模板系统 | 无 | 完整支持 |
| REPL 模式 | 有 | 有 |
| 输出格式 | JSON | JSON/YAML/Table/TSV |

## 开发

```bash
# 克隆
git clone https://github.com/yourname/wps-cli.git
cd wps-cli

# 安装开发依赖
pip install -e ".[dev]"

# 运行测试
pytest

# 代码检查
ruff check src/
```

### 项目结构

```
wps-cli/
├── src/wps_cli/
│   ├── main.py              # Typer 入口
│   ├── cli/                 # CLI 命令层
│   │   ├── common.py        # 公共工具
│   │   ├── writer.py        # Writer 命令
│   │   ├── calc.py          # Calc 命令
│   │   ├── impress.py       # Impress 命令
│   │   ├── pdf.py           # PDF 命令
│   │   └── export.py        # 导出命令
│   ├── services/            # 业务层
│   │   ├── writer_service.py
│   │   ├── calc_service.py
│   │   ├── impress_service.py
│   │   ├── pdf_service.py
│   │   ├── export_service.py
│   │   ├── style_engine.py
│   │   └── session_manager.py
│   ├── backends/            # COM 后端层
│   │   ├── base.py          # 抽象基类
│   │   └── wps_com.py       # WPS COM 实现
│   └── utils/
│       ├── path_utils.py
│       └── platform_check.py
├── tests/
├── pyproject.toml
├── LICENSE
└── README.md
```

## 许可证

[MIT](LICENSE)
