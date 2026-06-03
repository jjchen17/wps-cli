"""WPS CLI 桥接客户端 — 同步 HTTP client，供 backend 调用。

职责：
* 向 bridge server 发送 JSON 命令
* 等待并返回结果
* 自动管理 bridge daemon 的启动与停止
"""

from __future__ import annotations

import json
import subprocess
import sys
import time
import urllib.error
import urllib.request
import uuid
from pathlib import Path

DEFAULT_PORT = 3890
STARTUP_TIMEOUT = 5  # seconds to wait for daemon to start


class BridgeError(Exception):
    """桥接通信错误。"""


def _is_server_running(port: int = DEFAULT_PORT) -> bool:
    """Check if the bridge server is already listening."""
    try:
        req = urllib.request.Request(
            f"http://127.0.0.1:{port}/api",
            data=b"{}",
            headers={"Content-Type": "application/json"},
        )
        urllib.request.urlopen(req, timeout=1)
        return True
    except (urllib.error.URLError, OSError):
        return False


def _start_server(port: int = DEFAULT_PORT) -> None:
    """Start the bridge daemon as a subprocess."""
    server_path = Path(__file__).parent / "server.py"
    subprocess.Popen(
        [sys.executable, str(server_path), str(port)],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True,
    )

    deadline = time.time() + STARTUP_TIMEOUT
    while time.time() < deadline:
        if _is_server_running(port):
            return
        time.sleep(0.1)

    raise BridgeError(f"Bridge server failed to start on port {port}")


class BridgeClient:
    """同步客户端，与 bridge server 通信。"""

    def __init__(self, port: int = DEFAULT_PORT) -> None:
        self.port = port
        self._endpoint = f"http://127.0.0.1:{port}/api"

    def ensure_server(self) -> None:
        """确保 bridge daemon 在运行，否则启动。"""
        if not _is_server_running(self.port):
            _start_server(self.port)

    def send(self, method: str, params: dict | None = None, timeout: int = 60) -> dict:
        """发送命令并返回结果。

        Args:
            method: 命令名，格式 "{app}.{action}"，如 "writer.info"
            params: 命令参数
            timeout: 超时秒数

        Returns:
            result data dict (the "data" field of the response)

        Raises:
            BridgeError: 通信失败或 WPS 未连接
        """
        self.ensure_server()

        req_id = uuid.uuid4().hex[:12]
        payload = json.dumps(
            {"id": req_id, "method": method, "params": params or {}},
            ensure_ascii=False,
        ).encode("utf-8")

        req = urllib.request.Request(
            self._endpoint,
            data=payload,
            headers={"Content-Type": "application/json; charset=utf-8"},
            method="POST",
        )

        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                result: dict = json.loads(resp.read().decode("utf-8"))
        except urllib.error.URLError as exc:
            raise BridgeError(f"Cannot reach bridge server: {exc}") from exc
        except json.JSONDecodeError as exc:
            raise BridgeError(f"Invalid bridge server response: {exc}") from exc
        except OSError as exc:
            raise BridgeError(f"Bridge communication error: {exc}") from exc

        if not result.get("ok"):
            err = result.get("error", "Unknown error")
            raise BridgeError(err)

        return result.get("data")
