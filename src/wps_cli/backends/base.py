"""COM 后端抽象基类"""

from abc import ABC, abstractmethod


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
    def disconnect(self, app: object) -> None:
        """断开并关闭 WPS 应用实例"""

    @abstractmethod
    def is_alive(self, app: object) -> bool:
        """检查应用实例是否存活"""

    @abstractmethod
    def get_version(self, app: object) -> str:
        """获取 WPS 版本号"""
