"""COM 后端层"""

from wps_cli.backends.base import ComBackend
from wps_cli.backends.wps_com import WpsComBackend

__all__ = ["ComBackend", "WpsComBackend"]
