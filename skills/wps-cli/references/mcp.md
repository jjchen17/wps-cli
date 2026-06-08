# MCP 服务器配置

## 概述

wps-cli 内置 MCP (Model Context Protocol) 服务器，通过 JSON-RPC 2.0 over stdio 将全部文档操作能力暴露给 AI Agent。零外部依赖。

## 自动安装（推荐）

```bash
wps mcp serve                          # 手动启动 MCP stdio 服务器
wps mcp install --target claude        # 一键注册到 Claude Code
wps mcp install --target cursor        # 一键注册到 Cursor
wps mcp status                         # 检查注册状态
```

## 手动配置

在各 AI 工具的 MCP 配置文件中添加：

```json
{
  "wps-cli": {
    "command": "wps",
    "args": ["mcp", "serve"]
  }
}
```

### 各工具配置文件路径

| 工具 | MCP 配置文件 |
|------|-------------|
| Claude Code | `~/.claude/mcp.json` |
| Cursor | `.cursor/mcp.json` |
| VS Code / Cline | `~/.vscode/mcp.json` |
| Windsurf | `.windsurf/mcp.json` |
| Codex CLI | `.agents/mcp.json` |
| Hermes Agent | `.hermes/mcp.json` |
| MiniMax CLI | `.minimax/mcp.json` |
| OpenCode | `.opencode/mcp.json` |
| NanoBot | `.nanobot/mcp.json` |
| ZeroClaw | `.zeroclaw/mcp.json` |
| OpenClaw | `.openclaw/mcp.json` |

## MCP 工具列表（27 个）

### Writer 工具 (8)
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

### Calc 工具 (7)
| 工具名 | 功能 | 关键参数 |
|--------|------|---------|
| `calc_info` | 工作簿元信息 | file |
| `calc_cell_get` | 读取单元格 | file, ref, sheet |
| `calc_cell_set` | 设置单元格 | file, ref, value, sheet |
| `calc_range_get` | 读取区域 | file, ref, sheet |
| `calc_cell_formula` | 设置公式 | file, ref, formula, sheet |
| `calc_chart_create` | 创建图表 | file, data_range, chart_type, title, sheet |
| `calc_sheet_list` | 列出工作表 | file |

### Impress 工具 (5)
| 工具名 | 功能 | 关键参数 |
|--------|------|---------|
| `impress_info` | PPT 元信息 | file |
| `impress_slide_list` | 列出幻灯片 | file |
| `impress_text_get` | 读取幻灯片文本 | file, slide_idx |
| `impress_text_set` | 设置幻灯片文本 | file, slide_idx, placeholder, text |
| `impress_export_pdf` | PPT 导出 PDF | file, output |

### PDF 工具 (5)
| 工具名 | 功能 | 关键参数 |
|--------|------|---------|
| `pdf_info` | PDF 元信息 | file |
| `pdf_merge` | 合并 PDF | files_json, output |
| `pdf_extract_pages` | 提取页面 | file, pages, output |
| `pdf_watermark` | 添加水印 | file, text, output |
| `pdf_split` | 拆分 PDF | file, every, output_dir |

### Export 工具 (2)
| 工具名 | 功能 | 关键参数 |
|--------|------|---------|
| `export_convert` | 格式转换 | file, output_format, output |
| `export_batch` | 批量转换 | glob_pattern, output_format, output_dir |
