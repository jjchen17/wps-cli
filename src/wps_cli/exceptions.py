"""wps-cli 异常层次定义"""


class WpsCliError(Exception):
    """所有 wps-cli 异常的基类"""
    exit_code: int = 1


class WpsNotFoundError(WpsCliError):
    """WPS 未安装或未启动"""
    exit_code = 10

    def __init__(self, component: str = ""):
        name = f" {component}" if component else ""
        super().__init__(
            f"无法连接 WPS{name}。请确认 WPS Office 2019+ 已安装。\n"
            "运行 'wps doctor' 诊断环境。"
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
        super().__init__(msg)


class ComCallError(WpsCliError):
    """COM 调用失败"""
    exit_code = 30

    def __init__(self, operation: str, detail: str = ""):
        msg = f"WPS 操作失败: {operation}"
        if detail:
            msg += f" — {detail}"
        super().__init__(msg)


class FormatError(WpsCliError):
    """不支持的格式"""
    exit_code = 40

    def __init__(self, fmt: str, supported: str = ""):
        msg = f"不支持的格式: {fmt}"
        if supported:
            msg += f"。支持的格式: {supported}"
        super().__init__(msg)


class ValidationError(WpsCliError):
    """参数校验失败"""
    exit_code = 50
