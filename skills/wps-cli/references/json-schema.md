# JSON 输出格式

所有命令使用 `--json` / `-j` 标志后输出统一 JSON 结构。

## 成功响应

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

## 错误响应

```json
{
  "success": false,
  "command": "writer.info",
  "error": {
    "type": "FileNotFoundErrorCli",
    "message": "文件不存在: <path>",
    "code": 21,
    "suggestion": "请检查文件路径是否正确",
    "context": {"path": "<path>"}
  }
}
```

错误字段说明：
- `type`：异常类型标识符，便于 Agent 差异化处理
- `message`：人类可读错误描述（路径已脱敏）
- `code`：退出码（0-61），对应语义化退出码体系
- `suggestion`：修复建议，Agent 可据此自动尝试恢复
- `context`：附加上下文（可选，路径已脱敏）

## 退出码语义

| 退出码 | 含义 | Agent 行动建议 |
|--------|------|---------------|
| 0 | 成功 | 继续下一步 |
| 1 | 通用错误 | 查看 message 和 suggestion |
| 10 | WPS 未安装/未检测到 | 引导用户安装 WPS Office 2019+ |
| 11 | 会话管理失败 | 重试或重启 wps resident |
| 20 | 文件操作失败（路径、权限） | 检查文件权限和路径 |
| 21 | 文件不存在 | 确认文件路径，建议用户提供正确路径 |
| 30 | COM 调用失败 | 运行 wps doctor 诊断 |
| 40 | 不支持的格式 | 检查文件扩展名是否在支持列表中 |
| 50 | 参数校验失败 | 修正参数后重试 |
| 60 | 操作超时 | 增加超时时间或检查 WPS 是否卡死 |
| 61 | 批量操作部分失败 | 查看 per-step 结果获取具体失败项 |

## 批量命令输出格式

```json
{
  "success": false,
  "command": "batch",
  "data": {
    "steps": [
      {"index": 0, "success": true, "command": "writer.replace", "result": {"replaced": 3}},
      {"index": 1, "success": false, "command": "calc.cell-set", "error": {"type": "ValidationError", "message": "..."}},
      {"index": 2, "success": true, "command": "calc.cell-formula", "result": {"ref": "A1"}}
    ],
    "summary": {"total": 3, "succeeded": 2, "failed": 1}
  }
}
```

- 默认继续执行（continue-on-error），单条失败不影响后续
- 使用 `--stop-on-error` / `-s` 切换为遇错停止
- `summary` 提供快速统计，`steps` 提供逐条详情

## 错误信息脱敏

JSON 错误响应自动脱敏：
- 本地文件路径 → `<path>`
- 用户主目录 → `~`
- UNC 路径 → `<unc-path>`
- 相对路径 → `<rel-path>`

这是为了防止 AI Agent 上下文和日志中泄露文件系统结构。
