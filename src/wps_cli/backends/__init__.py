"""后端层 — 不同平台的 WPS 驱动实现"""

from wps_cli.backends.base import ComBackend
from wps_cli.backends.wps_com import WpsComBackend
from wps_cli.backends.wps_js import WpsJsBackend

__all__ = ["ComBackend", "WpsComBackend", "WpsJsBackend"]
