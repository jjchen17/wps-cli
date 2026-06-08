# 常见使用模式

## 模式 1：模板填充

使用 `{{key}}` 占位符实现"AI 设计一次模板，代码填充 N 次"。覆盖段落、表格单元格、页眉页脚中的占位符，保持原有格式不变。

```bash
# Step 1: 准备模板 template.docx（含 {{name}}, {{date}}, {{amount}} 等占位符）
# Step 2: 合并数据
wps writer merge template.docx -o output.docx --data '{"name":"张三","date":"2026-06-08","amount":"12,500"}'
# Step 3: 导出 PDF
wps writer export-pdf output.docx -o output.pdf
```

## 模式 2：数据报表生成

```bash
# Step 1: 创建 Excel 工作簿
wps calc new -o report.xlsx
# Step 2: 写入表头和明细
wps calc cell-set report.xlsx A1 "项目"
wps calc cell-set report.xlsx B1 "金额"
wps calc cell-set report.xlsx A2 "收入"
wps calc cell-set report.xlsx B2 10000
wps calc cell-set report.xlsx A3 "支出"
wps calc cell-set report.xlsx B3 7500
# Step 3: 设置汇总公式
wps calc cell-formula report.xlsx B4 "=SUM(B2:B3)"
# Step 4: 创建图表
wps calc chart-create report.xlsx -d A1:B3 -t bar --title "财务概览"
# Step 5: 导出 CSV
wps calc export-csv report.xlsx -o report.csv
```

## 模式 3：批量格式转换

```bash
# 批量 docx → pdf
wps export batch "reports/*.docx" -t pdf -d reports/pdf/

# 批量 xlsx → csv
wps export batch "data/*.xlsx" -t csv -d data/csv/
```

## 模式 4：PPT 内容维护

```bash
# Step 1: 查看幻灯片结构
wps impress slide-list deck.pptx --json
# Step 2: 修改特定幻灯片文本
wps impress text-set deck.pptx -s 1 -p title -t "年度总结"
wps impress text-set deck.pptx -s 1 -p body -t "2026年工作回顾..."
# Step 3: 添加新幻灯片
wps impress slide-add deck.pptx -l 1 -t "新章节"
# Step 4: 导出 PDF
wps impress export-pdf deck.pptx -o deck.pdf
```

## 模式 5：PDF 处理流水线

```bash
# Step 1: 提取需要的页面
wps pdf extract-pages source.pdf "1-3,7" -o extracted.pdf
# Step 2: 添加水印
wps pdf watermark extracted.pdf "内部资料" -o watermarked.pdf
# Step 3: 合并到最终文档
wps pdf merge watermarked.pdf appendix.pdf -o complete.pdf
```

## 模式 6：诊断-修复闭环（AI Agent 自愈）

```bash
# Step 1: 诊断文档问题
wps writer view report.docx issues --json
# Step 2: 根据诊断结果修复
wps writer replace report.docx "错误文本" "正确文本"
# Step 3: 验证修复
wps writer validate report.docx --json
# Step 4: 重新诊断确认
wps writer view report.docx issues --json
```

## 模式 7：批量命令执行

```bash
# 用单次命令操作同一文档的多个元素，避免多次 Open/Close
wps batch -c '[
  {"command": "calc.cell-set", "params": {"file": "data.xlsx", "ref": "A1", "value": "Q1"}},
  {"command": "calc.cell-set", "params": {"file": "data.xlsx", "ref": "B1", "value": 100}},
  {"command": "calc.cell-formula", "params": {"file": "data.xlsx", "ref": "C1", "formula": "=B1*1.1"}},
  {"command": "calc.chart-create", "params": {"file": "data.xlsx", "data_range": "A1:B1", "chart_type": "bar"}}
]'
```

## 模式 8：驻留模式高性能批量

```bash
# Step 1: 启动驻留进程（一次）
wps resident start --port 9123
# Step 2-N: 多次快速操作（5-10x 性能提升，无启动开销）
wps resident open report.docx --type writer
# ... 多次操作 ...
wps resident close report.docx
# 最后：停止驻留进程
wps resident stop
```
