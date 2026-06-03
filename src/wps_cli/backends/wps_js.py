"""WPS JS 后端 — 通过 bridge addon 驱动 WPS Office Linux 桌面端。

bridge 架构：
    Python CLI → BridgeClient (HTTP)
        → BridgeServer (WebSocket)
            → WPS JS Addon (WPS JS API)
                → WPS Application

仅负责「发命令、收结果」；所有 WPS 操作逻辑在 JS addon 中执行。
"""

from __future__ import annotations

from typing import Any

from wps_cli.backends.base import ComBackend
from wps_cli.bridge.client import BridgeClient, BridgeError
from wps_cli.exceptions import WpsNotFoundError


class WpsJsBackend(ComBackend):
    """通过 JS addon bridge 驱动 WPS Office Linux 桌面端。"""

    def __init__(self, port: int = 3889) -> None:
        self._client = BridgeClient(port=port)

    # ── ComBackend 接口 ──────────────────────────────────────

    def connect(self, app_type: str) -> BridgeClient:
        """验证 bridge 连通性并返回客户端句柄。

        ``app_type`` 仅用于校验：确认对应 WPS 组件已连上 bridge。
        """
        if app_type not in ("writer", "calc", "impress", "pdf"):
            raise ValueError(f"不支持的应用类型: {app_type}，可选: writer, calc, impress, pdf")
        try:
            self._client.ensure_server()
        except BridgeError as exc:
            raise WpsNotFoundError(app_type) from exc
        return self._client

    def disconnect(self, app: Any) -> None:
        """JS 后端无持久连接，此方法为空操作。"""

    def is_alive(self, app: Any) -> bool:
        """检查 bridge 可达性。"""
        try:
            self._client.ensure_server()
            return True
        except BridgeError:
            return False

    def get_version(self, app: Any) -> str:
        """获取 WPS 版本号。"""
        try:
            # 通过 writer app 获取版本（所有组件共享同一安装）
            data = self._client.send("writer.info", params={})
            return data.get("version", "unknown")
        except BridgeError:
            return "unknown"
