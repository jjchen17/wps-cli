# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""驻留模式 CLI 命令"""

from __future__ import annotations

import json as json_mod
from urllib import request
from urllib.error import URLError

import typer

from wps_cli.cli.common import handle_error, success
from wps_cli.utils.path_utils import ensure_safe_input_path

app = typer.Typer(help="驻留模式 -- 保持 COM 进程存活以加速连续操作")

DEFAULT_PORT = 9123
DEFAULT_HOST = "127.0.0.1"


def _api_url(host: str, port: int, endpoint: str) -> str:
    return f"http://{host}:{port}{endpoint}"


def _post(endpoint: str, payload: dict, host: str = DEFAULT_HOST, port: int = DEFAULT_PORT) -> dict:
    """向驻留服务发送 POST 请求"""
    req = request.Request(
        _api_url(host, port, endpoint),
        data=json_mod.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with request.urlopen(req, timeout=30) as resp:
            return json_mod.loads(resp.read().decode("utf-8"))
    except URLError as e:
        raise ConnectionError(f"无法连接驻留服务 ({host}:{port}): {e.reason}") from e


def _get(endpoint: str, host: str = DEFAULT_HOST, port: int = DEFAULT_PORT) -> dict:
    """向驻留服务发送 GET 请求"""
    try:
        with request.urlopen(_api_url(host, port, endpoint), timeout=10) as resp:
            return json_mod.loads(resp.read().decode("utf-8"))
    except URLError as e:
        raise ConnectionError(f"无法连接驻留服务 ({host}:{port}): {e.reason}") from e


@app.command()
def start(
    port: int = typer.Option(DEFAULT_PORT, "--port", "-p", help="HTTP 服务端口"),
    host: str = typer.Option(DEFAULT_HOST, "--host", "-h", help="监听地址"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """启动驻留进程（阻塞当前终端）"""
    cmd = "resident.start"
    try:
        from wps_cli.services.resident_daemon import ResidentDaemon

        daemon = ResidentDaemon(port=port, host=host)
        success(
            {"status": "starting", "host": host, "port": port},
            command=cmd,
            json_mode=json_output,
        )
        daemon.start()
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def stop(
    port: int = typer.Option(DEFAULT_PORT, "--port", "-p", help="HTTP 服务端口"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """停止驻留进程"""
    cmd = "resident.stop"
    try:
        result = _post("/shutdown", {}, port=port)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command("sessions")
def list_sessions(
    port: int = typer.Option(DEFAULT_PORT, "--port", "-p", help="HTTP 服务端口"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """列出当前活跃会话"""
    cmd = "resident.sessions"
    try:
        result = _get("/sessions", port=port)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def open(
    file: str = typer.Argument(..., help="文档路径"),
    type_: str = typer.Option("writer", "--type", "-t", help="文档类型: writer/calc/impress"),
    port: int = typer.Option(DEFAULT_PORT, "--port", "-p", help="HTTP 服务端口"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """在驻留进程中打开文档"""
    cmd = "resident.open"
    try:
        path = ensure_safe_input_path(file)
        result = _post("/open", {"path": str(path), "type": type_}, port=port)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def close(
    session_id: str = typer.Argument(..., help="会话 ID"),
    port: int = typer.Option(DEFAULT_PORT, "--port", "-p", help="HTTP 服务端口"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """关闭文档但不退出 COM 进程"""
    cmd = "resident.close"
    try:
        result = _post("/close", {"session_id": session_id}, port=port)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def status(
    port: int = typer.Option(DEFAULT_PORT, "--port", "-p", help="HTTP 服务端口"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """检查驻留服务状态"""
    cmd = "resident.status"
    try:
        result = _get("/status", port=port)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)
