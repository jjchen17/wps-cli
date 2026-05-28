"""wps-cli 异常层次定义

退出码语义：
  0    成功
  1    通用错误
  10   WPS 未安装/未启动
  11   会话管理失败
  20   文件操作失败（路径、权限）
  21   文件未找到
  30   COM 调用失败
  40   不支持的格式
  50   参数校验失败
  60   操作超时
  61   部分批量操作失败
"""

from __future__ import annotations


class WpsCliError(Exception):
    """所有 wps-cli 异常的基类"""

    exit_code: int = 1

    def __init__(self, message: str, *, suggestion: str = "", context: dict | None = None):
        super().__init__(message)
        self.suggestion = suggestion
        self.context = context or {}


class WpsNotFoundError(WpsCliError):
    """WPS 未安装或未启动"""

    exit_code = 10

    def __init__(self, component: str = ""):
        name = f" {component}" if component else ""
        super().__init__(
            f"无法连接 WPS{name}。请确认 WPS Office 2019+ 已安装。",
            suggestion="运行 'wps doctor' 诊断环境",
            context={"component": component} if component else {},
        )


class SessionError(WpsCliError):
    """COM 会话管理异常"""

    exit_code = 11


class FileOperationError(WpsCliError):
    """文件操作失败"""

    exit_code = 20

    def __init__(self, path: str, reason: str = ""):
        msg = f"文件操作失败: {path}"
        if reason:
            msg += f" — {reason}"
        super().__init__(msg, context={"path": path, "reason": reason})


class FileNotFoundErrorCli(WpsCliError):
    """指定文件不存在"""

    exit_code = 21

    def __init__(self, path: str):
        super().__init__(
            f"文件不存在: {path}",
            suggestion="请检查文件路径是否正确",
            context={"path": path},
        )


class ComCallError(WpsCliError):
    """COM 调用失败"""

    exit_code = 30

    def __init__(self, operation: str, detail: str = ""):
        msg = f"WPS 操作失败: {operation}"
        if detail:
            msg += f" — {detail}"
        super().__init__(msg, context={"operation": operation, "detail": detail})


class FormatError(WpsCliError):
    """不支持的格式"""

    exit_code = 40

    def __init__(self, fmt: str, supported: str = ""):
        msg = f"不支持的格式: {fmt}"
        if supported:
            msg += f"。支持的格式: {supported}"
        super().__init__(
            msg,
            suggestion=f"支持的格式: {supported}" if supported else "",
            context={"format": fmt, "supported": supported},
        )


class ValidationError(WpsCliError):
    """参数校验失败"""

    exit_code = 50


class TimeoutError(WpsCliError):  # noqa: A001 - 故意覆盖以提供项目内一致语义
    """操作超时"""

    exit_code = 60


class PartialFailureError(WpsCliError):
    """批量操作部分失败"""

    exit_code = 61
