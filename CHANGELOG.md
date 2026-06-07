# Changelog

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 与 [Semantic Versioning](https://semver.org/lang/zh-CN/)。

## [Unreleased]

### Added — 🚀 全面升级（借鉴 OfficeCLI 设计理念）
- **MCP 服务器**：内置 JSON-RPC 2.0 over stdio MCP 服务器，零外部依赖，暴露 26 个 tool（Writer 8 + Calc 7 + Impress 5 + PDF 5 + Export 1）。
  - CLI 命令：`wps mcp serve|install|status`，支持一键注册到 Claude Code / Cursor / VS Code。
  - 设计借鉴：iOfficeAI/OfficeCLI (Apache 2.0)
- **SKILL.md**：~390 行中文 AI Agent 教学文件，覆盖全部命令速查、典型模式、JSON schema 和退出码语义。
  - CLI 命令：`wps install skill|mcp|all-tools`，自动检测并安装到主流 AI 工具配置目录。
- **模板合并引擎**：`{{key}}` 占位符替换，覆盖段落/表格/页眉页脚，保持原格式不变。
  - CLI 命令：`wps writer merge`。
  - 设计借鉴：iOfficeAI/OfficeCLI (Apache 2.0)
- **驻留模式**：HTTP 服务器后台驻留 COM 进程，连续操作性能提升 5-10x，stdlib http.server 实现。
  - CLI 命令：`wps resident start|stop|sessions|open|close`。
  - 设计借鉴：iOfficeAI/OfficeCLI (Apache 2.0)
- **统一路径解析器**：1-based 路径语法（`/section[1]/paragraph[3]`、`/slide[1]/shape[2]`），支持 Writer/Calc/Impress 和 Excel 风格简写（`$Sheet1:A1`）。
  - CLI 命令：`wps writer|calc|impress get <path>`。
  - 设计借鉴：iOfficeAI/OfficeCLI (Apache 2.0)
- **语义视图与诊断**：`summarize()` 文档结构摘要 + `diagnose()` 文档问题检测（图片 alt text、公式错误、字体一致性等）。
  - CLI 命令：`wps writer|calc|impress view <file> [summary|issues|outline]`。
  - 设计借鉴：iOfficeAI/OfficeCLI (Apache 2.0)

### Changed — 安全加固
- **公式注入绕过修复 (CRITICAL)**：`_check_formula_safe` 改用 `re.sub(r"\s+", "", ...)` 移除所有 Unicode 空白字符（修复换行符/制表符/全角空格绕过）。
- `harden()` 静默失败改为记录 `logging.warning`（防止宏保护缺失不知情）。
- 添加 `INDIRECT()` 到危险公式函数名单。
- 修复 `redact_path()` 对含空格 Windows 路径脱敏不完整的问题。
- 文件扩展名白名单：所有 CLI 命令按应用类型（writer/calc/impress/pdf）拒绝不匹配的扩展名，防止 `wps calc cell-set README.md ...` 这类把任意文件作为工作簿覆盖的攻击。
- 危险公式黑名单补全：增加 `WEBSERVICE` / `FILTERXML` / `RTD` / `IMPORTDATA` / `IMPORTHTML` / `IMPORTRANGE` / `IMPORTXML` / `IMPORTFEED` / `ENCODEURL` 与对应的 `_xlfn.` 兼容前缀。
- 通配符替换反向引用防护：`writer replace --wildcard` 拒绝 `\1`-`\9` 反向引用（防止内容指数级膨胀），并对查找/替换文本加 1000 字符长度上限。
- glob 数量上限：`export batch` 的匹配结果数量不超过 200（可配），防止 `**` 触发大量 COM 操作（DoS）。
- glob 模式拒绝 `..`：前置拦截，不再依赖 resolve 后的最终防线。
- 路径脱敏增强：`redact_path` 现在覆盖 UNC 路径（`<unc-path>`）和相对路径（`<rel-path>`）。

### Added — 反馈闭环
- `wps doctor --report` 输出脱敏的 markdown 环境报告，可直接粘贴到 GitHub Issue。报告不包含文件路径、用户名、机器名等。
- Issue 模板改为 GitHub Issue Forms（`bug_report.yml`），强制必填诊断报告字段。

### Added — 项目元数据
- README 顶部加入非官方项目免责声明（中英双版）。
- 依赖更新自动化：`.github/dependabot.yml` 启用 GitHub Actions 与 pip 依赖每周巡检。

### Changed
- CI 主测试矩阵从 `windows-latest` 切换到 `ubuntu-latest`（Mock 已完整覆盖），保留一个 `windows-latest` smoke job 验证 pywin32 与 entry point。
- CI 增加 `concurrency.cancel-in-progress` 与 pip cache。
- CI build 步骤加 `twine check dist/*`，提前拦截 README 渲染问题。
- 覆盖率门禁设为 45%（当前实测 47%，逐版本上调）。
- `text_replace` 算法重写：原实现在 `new` 包含 `old` 时返回 0（如 `"foo"→"foobar"`），现在用 Find API `Execute(Replace=0)` 逐次扫描计数，结果精确。

### Removed
- 删除死代码 `src/wps_cli/utils/platform_check.py`（生产代码无人调用，0% 覆盖率）。
- 删除 `WriterService.open = open_document` 兼容别名（无人使用且遮蔽内置 `open`）。

## [Previous: 第一轮重构]
- 后端层新增 `ComBackend.harden()`，所有连接到 WPS 的进程都会强制禁用宏自动执行（`AutomationSecurity = msoAutomationSecurityForceDisable`）并关闭 `DisplayAlerts`。
- 公式注入防护：`calc cell-formula` 拒绝包含 `SHELL` / `DDE` / `DDEAUTO` / `EXEC` / `CALL` / `REGISTER` / `HYPERLINK` 等危险函数的公式。
- 单元格值二次注入防护：`calc cell-set` 拒绝以 `=` 开头的值。
- 路径边界限制：CLI 层入口统一通过 `ensure_safe_input_path` / `ensure_safe_output_path` 校验；`export batch` 的 glob 模式禁止绝对路径与 UNC 路径，匹配结果必须落在当前工作目录之下。
- PDF 页码上限：`pdf extract-pages` 的页码值与范围跨度均设硬上限，防止内存炸弹。
- 错误信息脱敏：JSON 错误响应中的本地路径会被替换为 `<path>`，用户主目录会被替换为 `~`。

### Added — AI Agent 集成
- 所有命令的 JSON 输出统一外层结构 `{success, command, data}` / `{success, command, error}`。
- 错误响应包含 `type`、`message`、`code`、`suggestion`、`context` 五个字段。
- 语义化退出码：`0/1/10/11/20/21/30/40/50/60/61`，对应不同失败类别，便于 Agent 自动差异化恢复。
- 新增 `English README` (`README.en.md`)。

### Changed
- `WriterService.open` 重命名为 `open_document`，保留 `open` 作为兼容别名。
- `text_replace` 改为单次遍历实现，使用文本计数估算替换次数；通配符模式下返回 `-1` 表示未知。
- `SessionManager` 增加 `threading.RLock`，会话 ID 改用 `uuid` 生成，支持作为上下文管理器。
- `cli.calc.chart-create` 的 `--type` 参数底层重命名为 `chart_type`，避免遮蔽 Python 内置。
- `cli.export.convert` / `cli.export.batch` 的 `format` 参数底层重命名为 `target_format`。
- `doctor` 命令的异常处理细化：区分 COM 错误、属性错误与其他异常，输出更精确的诊断信息。
- 错误处理统一走 `WpsCliError` 体系，每个异常类自带 `exit_code`、`suggestion`、`context`。

### Fixed
- `calc.info` 和 `export.convert` 在打开工作簿时显式设置 `UpdateLinks=0` + `ReadOnly=True`（适用时），避免外部链接自动加载。
- `pdf.split` 拒绝 `every <= 0`。
- `pdf.watermark` 拒绝超过 100 字符的水印文字。
- `WpsComBackend.is_alive` 使用 `pythoncom.com_error` 精确捕获 COM 异常，不再吞掉 `AttributeError` 之外的程序错误。

### Tests
- 新增 `tests/test_services/test_calc_service.py`：覆盖公式注入、单元格值校验。
- 新增 `tests/test_services/test_pdf_service.py`：覆盖 `_parse_pages` 边界条件与水印长度限制。
- 新增 `tests/test_utils/test_path_utils.py`：覆盖路径遍历、UNC、glob 边界。
- 新增 `tests/test_cli/test_common.py`：覆盖统一 JSON Schema 与错误脱敏。

### Docs
- README 增加退出码表、安全章节、JSON Schema 示例。
- 新增 `CONTRIBUTING.md`、`SECURITY.md`、Issue / PR 模板。

## [0.1.0] - 2026-05-28

### Added
- 首个版本，提供 Writer / Calc / Impress / PDF / Export 五大子命令共 36 条命令。
- 三层解耦架构：CLI → Service → COM Backend。
- `--json` 输出支持。
- `wps doctor` 环境诊断命令。
