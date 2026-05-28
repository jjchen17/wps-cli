"""COM 后端抽象基类

后端职责：
* 启动 / 关闭 WPS 应用进程
* 安全地打开文档（强制禁用宏自动执行）

具体子类只需实现进程级方法，文档级 ``open_*`` 方法保留默认实现。
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any


class ComBackend(ABC):
    """COM 后端抽象基类，隔离 win32com 依赖，方便测试和替换"""

    @abstractmethod
    def connect(self, app_type: str) -> object:
        """连接到 WPS 应用实例

        Args:
            app_type: 应用类型 (writer/calc/impress)
        Returns:
            COM 应用对象
        """

    @abstractmethod
    def disconnect(self, app: Any) -> None:
        """断开并关闭 WPS 应用实例"""

    @abstractmethod
    def is_alive(self, app: Any) -> bool:
        """检查应用实例是否存活"""

    @abstractmethod
    def get_version(self, app: Any) -> str:
        """获取 WPS 版本号"""

    def harden(self, app: Any) -> None:
        """对应用进程做安全加固：禁用宏自动执行、关闭警告弹窗等

        默认实现尝试设置 AutomationSecurity = msoAutomationSecurityForceDisable，
        失败时静默忽略（部分组件可能不支持该属性）。
        """
        try:
            from wps_cli.consts import MSO_AUTOMATION_SECURITY_FORCE_DISABLE

            app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        except Exception:
            pass
        try:
            app.DisplayAlerts = False
        except Exception:
            pass
