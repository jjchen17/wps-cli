"""WPS COM 后端实现"""

from __future__ import annotations

from typing import Any

from wps_cli.backends.base import ComBackend
from wps_cli.exceptions import WpsNotFoundError


class WpsComBackend(ComBackend):
    """通过 win32com 驱动 WPS Office 桌面端"""

    APP_MAP: dict[str, str] = {
        "writer": "KWPS.Application",
        "calc": "KET.Application",
        "impress": "KWPP.Application",
    }

    def connect(self, app_type: str) -> object:
        prog_id = self.APP_MAP.get(app_type)
        if not prog_id:
            raise ValueError(f"不支持的应用类型: {app_type}，可选: {list(self.APP_MAP.keys())}")
        try:
            import win32com.client
        except ImportError as exc:
            raise WpsNotFoundError(app_type) from exc

        try:
            app = win32com.client.Dispatch(prog_id)
        except Exception as exc:
            raise WpsNotFoundError(app_type) from exc
        self.harden(app)
        return app

    def disconnect(self, app: Any) -> None:
        try:
            import pythoncom

            com_error = pythoncom.com_error
        except ImportError:
            com_error = Exception
        try:
            app.Quit()
        except com_error:
            pass

    def is_alive(self, app: Any) -> bool:
        try:
            import pythoncom

            com_error = pythoncom.com_error
        except ImportError:
            com_error = Exception
        try:
            _ = app.Name
            return True
        except com_error:
            return False
        except AttributeError:
            return False

    def get_version(self, app: Any) -> str:
        try:
            return str(app.Version)
        except AttributeError:
            return "unknown"
        except Exception:
            return "unknown"
