# WPS CLI

[![Python](https://img.shields.io/badge/python-%3E%3D3.10-3776AB?logo=python&logoColor=white)](https://python.org)
[![Commands](https://img.shields.io/badge/commands-36-important)]()
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white)]()
[![WPS](https://img.shields.io/badge/requires-WPS%20Office%202019%2B-C00000)]()

在终端里指挥 WPS Office 干活 — 真格式、真排版、36 条命令。

> **不是"所见即所猜"，是所见即所得。** 不模拟文件格式，直接驱动真实的 WPS 引擎。

[English](README.en.md) · [更新日志](CHANGELOG.md) · [贡献指南](CONTRIBUTING.md)

---

## 为什么需要它

如果你用 python-docx 生成过带中文的 Word 文档，大概经历过这样的时刻：

- 明明代码里设了"**宋体**"，打开文件一看变成了 **Calibri**
- 用英文字段名操作中文段落，**排版永远对不齐**，调个缩进像在开盲盒
- 表格合并单元格后**边框消失**、页眉页脚完全不听使唤
- openpyxl 写个公式，打开 Excel **公式罢工**；条件格式更是千呼万唤出不来

这些库本质上是**在解析和拼接 XML 文件格式**，它们不知道"这个字体在 WPS 里实际渲染出来是什么样"，自然就无法保证最终效果。

**wps-cli 走了另一条路：不模拟文件格式，而是通过 COM 自动化直接驱动你电脑上正在运行的 WPS Office 本体。**

你设置"黑体三号加粗"，它就真的在 WPS 里执行一遍设置黑体三号加粗的操作。字体、间距、页码、目录、页眉页脚，每一个像素都和手动操作完全一致——因为本来就是同一个程序在干活。你调试时甚至可以盯着屏幕，看文档一步步"长出来"。

---

## 三大亮点

|  | | |
|---|---|---|
| **100% 格式保真** | 不是解析文件，是直接指挥 WPS 引擎干活 | 所见即所得，不是所见即所猜 |
| **AI Agent 原生支持** | 所有命令支持 `--json` 输出，统一 schema | 天生适配 LLM Agent 和自动化流水线 |
| **一把梭四件套** | Writer / Calc / Impress / PDF，36 个命令 | 办公自动化一个工具全搞定 |

---

## 一键安装

**系统要求**：Windows 10/11 · WPS Office 2019+ · Python 3.10+

```bash
pip install wps-cli
```

```bash
# 验证安装
wps doctor
```

> 国内网络环境推荐使用清华源加速：
> `pip install wps-cli -i https://pypi.tuna.tsinghua.edu.cn/simple`

---

## 三步体验

用三行命令，让 budget.xlsx 自动生成图表 + 提取数据 + 导出 PDF：

```bash
# Step 1: 一行命令，自动打开 Excel 插入柱状图
wps calc chart-create budget.xlsx --data "A1:C10" --type bar --title "2026年销售趋势"

# Step 2: 提取汇总数据，结构化 JSON 喂给下游脚本或 LLM
wps calc cell-get budget.xlsx C12 --json
# → {"success": true, "command": "calc.cell_get", "data": {"ref": "C12", "value": 285600}}

# Step 3: 一键导出 PDF，格式与 WPS 手动导出完全一致
wps export convert budget.xlsx pdf
```

三行命令，Excel 进去，图表 + 结构化数据 + PDF 出来 — 没有 GUI 点击，没有格式丢失。

---

## 命令体系

```
wps
├── writer          Word 文档
│   ├── new / info
│   ├── replace / count
│   ├── table-insert / table-get
│   ├── image-insert
│   ├── page-setup
│   ├── style-apply
│   └── export-pdf
├── calc            Excel 表格
│   ├── new / info
│   ├── sheet-list
│   ├── cell-get / cell-set / cell-range / cell-formula
│   ├── chart-create
│   ├── sort
│   └── export-csv
├── impress         PPT 演示
│   ├── new / info
│   ├── slide-list / slide-add / slide-delete
│   ├── text-set / text-get / image-insert
│   └── export-pdf
├── pdf             PDF 处理
│   ├── info
│   ├── merge / extract-pages / split
│   └── watermark
├── export          格式转换
│   ├── convert
│   └── batch
├── version         版本信息
└── doctor          环境诊断
```

| 全局选项 | 说明 |
|---------|------|
| `--json` / `-j` | 结构化 JSON 输出（适合 AI Agent / 管道） |

JSON 输出统一为：

```json
{
  "success": true,
  "command": "calc.cell_get",
  "data": { "...": "..." }
}
```

错误时：

```json
{
  "success": false,
  "command": "calc.cell_set",
  "error": {
    "type": "ValidationError",
    "message": "...",
    "code": 50,
    "suggestion": "...",
    "context": { "...": "..." }
  }
}
```

---

## 退出码

| 退出码 | 含义 |
|------:|------|
| 0  | 成功 |
| 1  | 通用错误 |
| 10 | WPS 未安装或未检测到 |
| 11 | 会话管理错误 |
| 20 | 文件操作失败 |
| 21 | 文件不存在 |
| 30 | COM 调用失败 |
| 40 | 不支持的格式 |
| 50 | 参数校验失败 |
| 60 | 操作超时 |
| 61 | 批量操作部分失败 |

---

## 安全性

wps-cli 默认开启以下加固，避免常见的 Office 自动化滥用路径：

- **强制禁用宏自动执行**：所有 `Documents.Open` / `Workbooks.Open` / `Presentations.Open` 调用均设置 `AutomationSecurity = msoAutomationSecurityForceDisable`，阻止 `Auto_Open` / `Document_Open` 类宏在打开文档时被触发。
- **公式注入防护**：`calc cell-formula` 拒绝包含 `SHELL` / `DDE` / `DDEAUTO` / `EXEC` / `CALL` / `REGISTER` / `HYPERLINK` 等危险函数的公式；`calc cell-set` 拒绝以 `=` 开头的值，避免二次公式注入。
- **路径边界限制**：所有命令的文件参数都会经过路径校验，禁止 UNC 路径与符号链接；`export batch` 的 glob 模式禁止使用绝对路径，匹配结果必须落在当前工作目录之下。
- **页码内存炸弹防护**：`pdf extract-pages` 的页码范围有上限，避免 `1-999999` 这类输入耗尽内存。
- **错误信息脱敏**：JSON 错误响应中的本地路径会被替换为 `<path>`，用户主目录会被替换为 `~`，避免日志/Agent 上下文泄露文件系统结构。

详见 [SECURITY.md](SECURITY.md)。

---

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

三层解耦：CLI 层只做参数解析和输出格式化，业务层通过抽象接口调用后端，后端层封装 COM 细节，可替换为 LibreOffice / WebOffice。

---

## 开发

```bash
git clone https://github.com/jjchen17/wps-cli.git
cd wps-cli
pip install -e ".[dev]" -i https://pypi.tuna.tsinghua.edu.cn/simple
pytest
```

```
src/wps_cli/
├── cli/            # CLI 命令层 (Typer)
├── services/       # 业务层 (Writer/Calc/Impress/PDF/Export)
├── backends/       # COM 后端层 (抽象基类 + WPS 实现)
└── utils/          # 工具函数
```

详见 [CONTRIBUTING.md](CONTRIBUTING.md)。

---

## 反馈

欢迎提交 [Issue](https://github.com/jjchen17/wps-cli/issues) 或发送邮件至 **[948881912@qq.com](mailto:948881912@qq.com)**。

## 许可证

[MIT](LICENSE)
