# 安全策略

## 报告漏洞

**请不要通过公开 GitHub Issue 报告安全漏洞。**

请发送邮件至 [948881912@qq.com](mailto:948881912@qq.com)，主题以 `[SECURITY]` 开头，包含：

- 受影响的版本
- 漏洞描述与威胁模型
- 复现步骤（或概念验证代码）
- 你建议的修复方式（可选）

我们会在 72 小时内回复，并在确认漏洞后协调披露时间。

## 支持的版本

| 版本 | 安全更新 |
|-----|---------|
| 0.1.x | ✅ |

## 默认安全加固

wps-cli 在每次连接 WPS 进程时自动应用以下加固，无需用户配置：

### 宏自动执行禁用

`ComBackend.harden()` 设置 `app.AutomationSecurity = msoAutomationSecurityForceDisable`，
阻止打开文档时触发 `Auto_Open` / `Document_Open` 等自动宏。

威胁场景：攻击者向 LLM Agent 投递包含恶意宏的文档，诱导 Agent 调用 `wps writer info`
打开文件。无加固时宏会以当前用户权限执行任意代码。

### 公式注入防护

`calc cell-formula` 命令在调用 COM 之前拒绝以下危险函数：

- `SHELL` / `EXEC` / `CALL` — 命令执行
- `DDE` / `DDEAUTO` — DDE 协议命令执行
- `REGISTER` — 加载外部 DLL
- `HYPERLINK` — 数据外泄

`calc cell-set` 拒绝以 `=` 开头的值，避免攻击者通过 `cell-set` 绕过 `cell-formula` 的校验。

### 路径边界限制

- 禁止 UNC 路径（`\\server\share`）—— 防止 NTLM 哈希泄露。
- 禁止符号链接 —— 防止 TOCTOU 与目录穿越。
- `export batch` 的 glob 模式必须是相对当前工作目录的相对路径，结果必须落在当前目录之下。

### 内存炸弹防护

`pdf extract-pages` 的页码字符串会被严格校验：

- 单个页码不超过 9999
- 单个区间跨度不超过 1000
- 总页码数不超过 1000

防止 `1-999999` 这种输入耗尽内存。

### 错误信息脱敏

JSON 错误响应中的本地路径在序列化前会经过 `redact_path()` 处理：

- 用户主目录 → `~`
- Windows 盘符路径 → `<path>`

避免 LLM Agent 上下文 / CI 日志泄露文件系统结构。

## 已知限制

- **WPS 桌面进程权限继承**：wps-cli 启动的 WPS 进程继承调用者的 Windows 安全上下文。如果你以管理员身份运行 wps-cli，WPS 进程也是管理员权限。生产环境建议用最低权限账号运行。
- **DCOM 暴露**：如果你手动开启了 KWPS / KET / KWPP 的 DCOM 访问，远程机器可能调用本地 WPS。这超出 wps-cli 控制范围，请检查 Windows COM 安全设置。
- **依赖供应链**：`pywin32` 等核心依赖若被投毒，影响所有 wps-cli 用户。我们会在 `pyproject.toml` 中固定版本上限。

## 安全补丁发布

- 高危漏洞：72 小时内发布补丁版本。
- 中危漏洞：在下一个常规版本中修复。
- 低危漏洞：累积到主要版本中修复。

补丁版本会在 [CHANGELOG.md](CHANGELOG.md) 中明确标注 `Security` 字样。
