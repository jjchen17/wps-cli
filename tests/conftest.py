"""公共测试 fixtures"""

from __future__ import annotations

from typing import Any

import pytest

from wps_cli.backends.base import ComBackend


class MockApp:
    """模拟 WPS 应用对象，记录所有 COM 调用便于断言"""

    def __init__(self, app_type: str):
        self.Name = f"Mock {app_type}"
        self.Version = "12.0.0-test"
        self.Visible = False
        self.AutomationSecurity = 0
        self.DisplayAlerts = True
        self._app_type = app_type
        self.calls: list[tuple[str, tuple, dict]] = []

    def Quit(self) -> None:
        self.calls.append(("Quit", (), {}))


class MockComBackend(ComBackend):
    """模拟 COM 后端，用于测试"""

    def __init__(self) -> None:
        self.last_app: MockApp | None = None

    def connect(self, app_type: str) -> object:
        app = MockApp(app_type)
        self.last_app = app
        self.harden(app)
        return app

    def disconnect(self, app: Any) -> None:
        try:
            app.Quit()
        except Exception:
            pass

    def is_alive(self, app: Any) -> bool:
        return True

    def get_version(self, app: Any) -> str:
        return "12.0.0-test"


@pytest.fixture
def mock_backend() -> MockComBackend:
    """提供模拟 COM 后端"""
    return MockComBackend()


@pytest.fixture
def mock_app() -> MockApp:
    """提供模拟 WPS 应用对象"""
    return MockApp("writer")
