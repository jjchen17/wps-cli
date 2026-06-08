# 命令速查

> 所有命令均支持 `--json` / `-j` 标志输出 JSON 格式。
> 全局选项：`--help` 查看帮助，`--json` 输出 JSON。

## 环境诊断

```bash
wps doctor                 # 诊断环境：Python、pywin32、WPS 组件
wps doctor --fix           # 自动修复 COM 注册
wps doctor --report        # 输出脱敏 Markdown 诊断报告
wps doctor --verbose       # 详细诊断
wps version                # 输出版本号
```

## Word 操作 (writer)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps writer new -o <path>` | 新建空白文档 | `wps writer new -o blank.docx` |
| `wps writer info <file>` | 文档元信息 | `wps writer info report.docx --json` |
| `wps writer count <file>` | 字数统计 | `wps writer count report.docx --json` |
| `wps writer replace <file> <old> <new>` | 查找替换 | `wps writer replace doc.docx "旧" "新"` |
| `wps writer replace <file> <old> <new> -w` | 通配符替换 | `wps writer replace doc.docx "张?" "李?" -w` |
| `wps writer replace <file> <old> <new> -c` | 区分大小写 | `wps writer replace doc.docx "WPS" "wps" -c` |
| `wps writer replace <file> <old> <new> --regex` | 正则替换 | `wps writer replace doc.docx "\\d+" "N" --regex` |
| `wps writer table-get <file> -i 1` | 读取第1个表格 | `wps writer table-get data.docx -i 1 --json` |
| `wps writer table-insert <file> -r 3 -c 4 -d '[["A","B"]]'` | 插入表格 | `wps writer table-insert doc.docx -r 3 -c 4 -d '[[...]]'` |
| `wps writer image-insert <file> -i <image>` | 插入图片 | `wps writer image-insert doc.docx -i photo.png` |
| `wps writer image-insert <file> -i <image> -w 200 -h 150` | 插入图片指定尺寸 | `wps writer image-insert doc.docx -i photo.png -w 200 -h 150` |
| `wps writer page-setup <file>` | 页面布局 | `wps writer page-setup doc.docx --width 210 --height 297` |
| `wps writer style-apply <file> <preset>` | 应用预设样式 | `wps writer style-apply doc.docx "公文正文"` |
| `wps writer style-apply <file> "" -l` | 列出样式预设 | `wps writer style-apply doc.docx "" -l` |
| `wps writer export-pdf <file> -o <output>` | 导出 PDF | `wps writer export-pdf doc.docx -o output.pdf` |
| `wps writer merge <file> -o <out> --data '{...}'` | 模板合并 | `wps writer merge template.docx -o out.docx --data '{"name":"张三"}'` |
| `wps writer view <file> <mode>` | 语义视图 | `wps writer view doc.docx summary` |
| `wps writer view <file> <mode> --type <subtype>` | 按类型过滤 | `wps writer view doc.docx issues --type formula_eval_error` |
| `wps writer validate <file>` | 文档验证 | `wps writer validate doc.docx --json` |
| `wps writer refresh <file>` | 刷新字段 | `wps writer refresh doc.docx` |
| `wps writer get <file> <path>` | 路径定位 | `wps writer get doc.docx "/section[1]/paragraph[3]"` |
| `wps writer formfield-list <file>` | 列出表单域 | `wps writer formfield-list doc.docx --json` |
| `wps writer formfield-get <file> <name>` | 读取表单域 | `wps writer formfield-get doc.docx "Text1" --json` |
| `wps writer formfield-set <file> <name> <value>` | 设置表单域 | `wps writer formfield-set doc.docx "Text1" "新值"` |
| `wps writer contentcontrol-list <file>` | 列出内容控件 | `wps writer contentcontrol-list doc.docx --json` |
| `wps writer contentcontrol-set <file> <tag> <value>` | 设置内容控件 | `wps writer contentcontrol-set doc.docx "tag1" "新值"` |

**view 模式**：`summary` / `issues` / `outline` / `annotated` / `stats`

**样式预设**：公文标题、公文一级标题、公文二级标题、公文正文、报告标题、报告正文

## Excel 操作 (calc)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps calc new -o <path>` | 新建工作簿 | `wps calc new -o blank.xlsx` |
| `wps calc info <file>` | 工作簿元信息 | `wps calc info data.xlsx --json` |
| `wps calc sheet-list <file>` | 列出所有工作表 | `wps calc sheet-list data.xlsx --json` |
| `wps calc cell-get <file> <ref>` | 读取单元格值 | `wps calc cell-get data.xlsx A1 --json` |
| `wps calc cell-get <file> <ref> -s <sheet>` | 指定工作表的单元格 | `wps calc cell-get data.xlsx B3 -s Sheet2 --json` |
| `wps calc cell-set <file> <ref> <value>` | 设置单元格值 | `wps calc cell-set data.xlsx A1 "Hello"` |
| `wps calc cell-set <file> <ref> <value> -s <sheet>` | 指定工作表设置值 | `wps calc cell-set data.xlsx A1 "Hi" -s Sheet2` |
| `wps calc cell-range <file> <ref>` | 读取区域 | `wps calc cell-range data.xlsx A1:D10 --json` |
| `wps calc cell-formula <file> <ref> <formula>` | 设置公式 | `wps calc cell-formula data.xlsx B10 "=SUM(B1:B9)"` |
| `wps calc chart-create <file> -d <range> -t <type> --title <title>` | 创建图表 | `wps calc chart-create data.xlsx -d A1:C10 -t pie --title "销售占比"` |
| `wps calc sort <file> -b <col> --order <asc/desc>` | 排序 | `wps calc sort data.xlsx -b A --order desc` |
| `wps calc export-csv <file> -o <output>` | 导出 CSV | `wps calc export-csv data.xlsx -o data.csv` |
| `wps calc view <file> <mode>` | 语义视图 | `wps calc view data.xlsx summary` |
| `wps calc validate <file>` | 文档验证 | `wps calc validate data.xlsx --json` |
| `wps calc refresh <file>` | 刷新公式/透视表 | `wps calc refresh data.xlsx` |
| `wps calc get <file> <path>` | 路径定位 | `wps calc get data.xlsx '/sheet["Sheet1"]/cell["C12"]'` |
| `wps calc conditional-format-add <file> ...` | 添加条件格式 | 支持 cellvalue/databar/colorscale/iconset 等 |
| `wps calc conditional-format-list <file>` | 列出条件格式 | `wps calc conditional-format-list data.xlsx --json` |
| `wps calc conditional-format-delete <file> <index>` | 删除条件格式 | `wps calc conditional-format-delete data.xlsx 1` |
| `wps calc data-validation-add <file> ...` | 添加数据验证 | 支持 list/whole/decimal/date/time/textlength/custom |
| `wps calc data-validation-list <file>` | 列出数据验证 | `wps calc data-validation-list data.xlsx --json` |
| `wps calc sparkline-add <file> ...` | 添加迷你图 | 支持 line/column/stacked100 |

**图表类型**：`bar`（柱状图）、`line`（折线图）、`pie`（饼图）、`scatter`（散点图）、`area`（面积图）

**view 模式**：`summary` / `issues` / `sheets`

## PPT 操作 (impress)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps impress new -o <path>` | 新建演示文稿 | `wps impress new -o blank.pptx` |
| `wps impress info <file>` | 演示文稿元信息 | `wps impress info slides.pptx --json` |
| `wps impress slide-list <file>` | 列出所有幻灯片 | `wps impress slide-list slides.pptx --json` |
| `wps impress slide-add <file> -l 1 -t "标题"` | 新增幻灯片 | `wps impress slide-add slides.pptx -l 1 -t "新页面"` |
| `wps impress slide-delete <file> <index>` | 删除幻灯片 | `wps impress slide-delete slides.pptx 3` |
| `wps impress text-get <file> -s <idx>` | 获取幻灯片文本 | `wps impress text-get slides.pptx -s 2 --json` |
| `wps impress text-set <file> -s <idx> -p title -t "标题"` | 设置占位符文本 | `wps impress text-set slides.pptx -s 1 -p title -t "新标题"` |
| `wps impress image-insert <file> -s <idx> -i <image>` | 插入图片 | `wps impress image-insert slides.pptx -s 2 -i logo.png` |
| `wps impress export-pdf <file> -o <output>` | 导出 PDF | `wps impress export-pdf slides.pptx -o slides.pdf` |
| `wps impress view <file> <mode>` | 语义视图 | `wps impress view slides.pptx summary` |
| `wps impress validate <file>` | 文档验证 | `wps impress validate slides.pptx --json` |
| `wps impress refresh <file>` | 刷新链接 | `wps impress refresh slides.pptx` |
| `wps impress get <file> <path>` | 路径定位 | `wps impress get slides.pptx "/slide[1]/shape[2]"` |

**占位符类型**：`title`（标题）、`body`（正文）、`subtitle`（副标题）

**view 模式**：`summary` / `issues` / `slides`

## PDF 操作 (pdf)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps pdf info <file>` | PDF 元信息 | `wps pdf info document.pdf --json` |
| `wps pdf merge <f1> <f2> ... -o <output>` | 合并 PDF | `wps pdf merge a.pdf b.pdf -o merged.pdf` |
| `wps pdf extract-pages <file> <pages> -o <output>` | 提取页面 | `wps pdf extract-pages doc.pdf "1-3,5,7-9" -o extracted.pdf` |
| `wps pdf split <file> -e <N> -d <dir>` | 每 N 页拆分 | `wps pdf split doc.pdf -e 10 -d ./split/` |
| `wps pdf watermark <file> <text> -o <output>` | 添加水印 | `wps pdf watermark doc.pdf "机密" -o watermarked.pdf` |

## 格式转换 (export)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps export convert <file> <format> -o <output>` | 单文件转换 | `wps export convert doc.docx pdf -o doc.pdf` |
| `wps export batch <glob> -t <format> -d <dir>` | 批量转换 | `wps export batch "*.docx" -t pdf -d ./pdfs/` |

**支持转换**：Writer→pdf/docx/doc/rtf/txt/html；Calc→xlsx/csv；Impress→pptx/ppt/pdf

## 驻留模式 (resident)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps resident start --port 9123` | 启动后台 HTTP 服务 | `wps resident start` |
| `wps resident open <file> --type writer` | 打开文档 | `wps resident open report.docx --type writer` |
| `wps resident close <file>` | 关闭文档 | `wps resident close report.docx` |
| `wps resident status` | 查看服务状态 | `wps resident status --json` |
| `wps resident sessions` | 查看活跃会话 | `wps resident sessions --json` |
| `wps resident stop` | 停止服务 | `wps resident stop` |

## MCP 服务器 (mcp)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps mcp serve` | 启动 MCP stdio 服务器 | `wps mcp serve` |
| `wps mcp install -t claude` | 安装 MCP 配置 | `wps mcp install -t claude` |
| `wps mcp status` | 检查注册状态 | `wps mcp status --json` |

## AI 工具集成 (install)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps install skill -t claude` | 安装 SKILL.md | `wps install skill -t claude` |
| `wps install mcp -t claude` | 安装 MCP 配置 | `wps install mcp -t claude` |
| `wps install all-tools` | 一键安装所有 | `wps install all-tools` |

## 批量执行 (batch)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps batch -c '[...]'` | 直接传入 JSON 命令数组 | `wps batch -c '[{"command":"calc.cell-set",...}]'` |
| `wps batch -i cmds.json` | 从文件读取 | `wps batch -i commands.json` |
| `wps batch -j` | 从 stdin 读取 | `echo '[...]' \| wps batch -j` |
| `wps batch -c '[...]' -s` | 遇错停止 | `wps batch -c '[...]' --stop-on-error` |

## 文档序列化 (dump)

| 命令 | 说明 | 示例 |
|------|------|------|
| `wps dump <file>` | 全文档序列化 | `wps dump report.docx --json` |
| `wps dump <file> --path "/section[1]"` | 子树序列化 | `wps dump doc.docx --path "/section[1]"` |
