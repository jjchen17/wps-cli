"""WPS COM 后端实现"""

from typing import Any

import win32com.client

from wps_cli.backends.base import ComBackend


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
        return win32com.client.Dispatch(prog_id)

    def disconnect(self, app: Any) -> None:
        try:
            app.Quit()
        except Exception:
            pass  # 进程已退出时忽略

    def is_alive(self, app: Any) -> bool:
        try:
            _ = app.Name
            return True
        except Exception:
            return False

    def get_version(self, app: Any) -> str:
        try:
            return str(app.Version)
        except Exception:
            return "unknown"
