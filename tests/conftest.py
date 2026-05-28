"""公共测试 fixtures"""

from typing import Any

import pytest


class MockComBackend:
    """模拟 COM 后端，用于测试"""

    def connect(self, app_type: str) -> object:
        return MockApp(app_type)

    def disconnect(self, app: Any) -> None:
        pass

    def is_alive(self, app: Any) -> bool:
        return True

    def get_version(self, app: Any) -> str:
        return "12.0.0-test"


class MockApp:
    """模拟 WPS 应用对象"""

    def __init__(self, app_type: str):
        self.Name = f"Mock {app_type}"
        self.Version = "12.0.0-test"
        self.Visible = False
        self._app_type = app_type

    def Quit(self):
        pass


@pytest.fixture
def mock_backend():
    """提供模拟 COM 后端"""
    return MockComBackend()


@pytest.fixture
def mock_app():
    """提供模拟 WPS 应用对象"""
    return MockApp("writer")
