"""平台兼容性检查"""

import sys


def check_windows() -> None:
    """检查是否在 Windows 上运行"""
    if sys.platform != "win32":
        raise RuntimeError(
            "wps-cli 仅支持 Windows 系统（需要 COM 自动化接口）\n"
            f"当前平台: {sys.platform}"
        )


def check_pywin32() -> None:
    """检查 pywin32 是否已安装"""
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        raise RuntimeError(
            "pywin32 未安装，请执行: pip install pywin32"
        ) from None


def check_environment() -> None:
    """完整环境检查"""
    check_windows()
    check_pywin32()
