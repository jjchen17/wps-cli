"""WPS CLI 桥接服务器 — HTTP + WebSocket daemon.

职责：
* 接收 CLI 端 HTTP POST 命令，转发给 WPS JS addon
* 维持与 WPS addon 的 WebSocket 长连接
* 将 addon 执行结果回传给 CLI 端
"""

from __future__ import annotations

import asyncio
import json
import os
import sys
from pathlib import Path

import websockets
from websockets.asyncio.server import ServerConnection

ADDON_DIR = Path(__file__).parent / "addon"
DEFAULT_PORT = 3890
REQUEST_TIMEOUT = 60  # seconds

# MIME types for static files served to WPS
_MIME = {
    ".html": "text/html; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
}


class _BridgeState:
    """单例共享状态，存于模块级别以避免闭包绑定问题。"""

    def __init__(self, port: int = DEFAULT_PORT) -> None:
        self.port = port
        self._connections: dict[str, ServerConnection] = {}  # app_type -> ws
        self._pending: dict[str, asyncio.Future[dict]] = {}

    def register_ws(self, app_type: str, ws: ServerConnection) -> None:
        self._connections[app_type] = ws

    def unregister_ws(self, app_type: str, ws: ServerConnection) -> None:
        if self._connections.get(app_type) is ws:
            del self._connections[app_type]

    def get_ws(self, app_type: str) -> ServerConnection | None:
        return self._connections.get(app_type)

    def create_future(self, req_id: str) -> asyncio.Future[dict]:
        loop = asyncio.get_running_loop()
        fut: asyncio.Future[dict] = loop.create_future()
        self._pending[req_id] = fut
        return fut

    def resolve(self, req_id: str, result: dict) -> bool:
        fut = self._pending.pop(req_id, None)
        if fut and not fut.done():
            fut.set_result(result)
            return True
        return False


_state: _BridgeState | None = None


def _get_state() -> _BridgeState:
    global _state
    if _state is None:
        port = int(os.environ.get("WPS_BRIDGE_PORT", str(DEFAULT_PORT)))
        _state = _BridgeState(port)
    return _state


def _http_response(
    status: int, body: dict | str, extra_headers: dict[str, str] | None = None
) -> tuple[int, str, list[tuple[str, str]], bytes]:
    """Construct an HTTP response tuple for the websockets process_request hook."""
    if isinstance(body, dict):
        body_str = json.dumps(body, ensure_ascii=False)
        content_type = "application/json; charset=utf-8"
    else:
        body_str = body
        content_type = "text/html; charset=utf-8"

    headers: list[tuple[str, str]] = [("Content-Type", content_type)]
    if extra_headers:
        headers.extend(extra_headers.items())
    headers.append(("Access-Control-Allow-Origin", "*"))
    headers.append(("Access-Control-Allow-Headers", "Content-Type"))
    headers.append(("Access-Control-Allow-Methods", "POST, GET, OPTIONS"))

    return status, "OK" if status < 400 else "Error", headers, body_str.encode("utf-8")


# ── WebSocket handler ──────────────────────────────────────────


async def _ws_handler(ws: ServerConnection) -> None:
    """处理来自 WPS JS addon 的 WebSocket 连接。"""
    state = _get_state()
    app_type: str | None = None

    try:
        async for raw in ws:
            try:
                msg: dict = json.loads(raw)
            except json.JSONDecodeError:
                continue

            # 握手：识别是哪个 WPS 组件
            if msg.get("type") == "hello":
                app_type = msg.get("app", "writer")
                state.register_ws(app_type, ws)
                await ws.send(json.dumps({"type": "ok"}))
                continue

            # 响应：匹配到等待中的 CLI 请求
            rid = msg.get("id")
            if rid:
                state.resolve(rid, msg)

    except websockets.exceptions.ConnectionClosed:
        pass
    finally:
        if app_type:
            state.unregister_ws(app_type, ws)


# ── HTTP handler ───────────────────────────────────────────────


async def _serve_static(path: str) -> tuple[int, str, list[tuple[str, str]], bytes] | None:
    """Serve static addon files. Returns None if not a static request."""
    prefix = "/addon/"
    if not (path.startswith(prefix) and len(path) > len(prefix)):
        return None

    filename = path[len(prefix) :] or "index.html"
    # Security: prevent path traversal
    if ".." in filename or "/" in filename:
        return _http_response(403, "Forbidden")

    filepath = ADDON_DIR / filename
    if not filepath.is_file():
        return _http_response(404, "Not found")

    ext = filepath.suffix
    mime = _MIME.get(ext, "application/octet-stream")
    return (200, "OK", [("Content-Type", mime)], filepath.read_bytes())


async def _handle_api(body: bytes) -> tuple[int, str, list[tuple[str, str]], bytes]:
    """Handle POST /api — forward command to WPS addon and wait for result."""
    state = _get_state()

    try:
        cmd: dict = json.loads(body)
    except json.JSONDecodeError:
        return _http_response(400, {"error": "Invalid JSON"})

    req_id = cmd.get("id")
    if not req_id:
        return _http_response(400, {"error": "Missing request id"})

    method: str = cmd.get("method", "")
    app_type = method.split(".")[0] if "." in method else "writer"

    ws = state.get_ws(app_type)
    if not ws:
        return _http_response(
            503,
            {
                "id": req_id,
                "ok": False,
                "error": f"WPS {app_type} not connected. Please start WPS {app_type.title()} first.",
            },
        )

    try:
        fut = state.create_future(req_id)
        await ws.send(json.dumps(cmd, ensure_ascii=False))
        result = await asyncio.wait_for(fut, timeout=REQUEST_TIMEOUT)
    except asyncio.TimeoutError:
        state.resolve(req_id, {"id": req_id, "ok": False, "error": "Request timed out"})
        return _http_response(504, {"id": req_id, "ok": False, "error": "Request timed out"})
    except websockets.exceptions.ConnectionClosed:
        return _http_response(
            503,
            {
                "id": req_id,
                "ok": False,
                "error": f"WPS {app_type} connection lost. Please restart WPS.",
            },
        )

    return _http_response(200, result)


async def _http_handler(
    connection: ServerConnection, request: websockets.http11.Request
) -> websockets.http11.Response | None:
    """HTTP 请求分发：静态文件 | API | OPTIONS preflight | WebSocket 升级。"""
    path = request.path or "/"
    method = request.method or "GET"

    # CORS preflight
    if method == "OPTIONS":
        _, _, headers, body = _http_response(204, "")
        return websockets.http11.Response(204, "No Content", headers, body)

    # Static files
    static_resp = await _serve_static(path)
    if static_resp:
        status, reason, headers, body = static_resp
        return websockets.http11.Response(status, reason, headers, body)

    # API endpoint
    if path == "/api" and method == "POST":
        body = await request.read() if hasattr(request, "read") else b""
        status, reason, headers, body = await _handle_api(body)
        return websockets.http11.Response(status, reason, headers, body)

    # Fallback: check if it's a WebSocket upgrade
    if method == "GET" and request.headers.get("Upgrade", "").lower() == "websocket":
        return None  # upgrade to WebSocket

    status, reason, headers, body = _http_response(404, "Not found")
    return websockets.http11.Response(status, reason, headers, body)


# ── Server lifecycle ───────────────────────────────────────────


async def _serve(port: int) -> None:
    """Start the bridge server and run forever."""
    async with websockets.asyncio.server.serve(
        _ws_handler,
        host="127.0.0.1",
        port=port,
        process_request=_http_handler,
    ):
        await asyncio.get_running_loop().create_future()  # run forever


def run_server(port: int = DEFAULT_PORT) -> None:
    """同步入口：启动 bridge daemon（阻塞调用）。"""
    os.environ["WPS_BRIDGE_PORT"] = str(port)

    # Write PID file for daemon management
    pid_file = Path(os.environ.get("XDG_RUNTIME_DIR", "/tmp")) / "wps-bridge.pid"
    pid_file.write_text(str(os.getpid()))

    try:
        asyncio.run(_serve(port))
    except KeyboardInterrupt:
        pass
    finally:
        pid_file.unlink(missing_ok=True)


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_PORT
    run_server(port)
