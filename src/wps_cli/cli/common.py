"""CLI 公共工具：输出格式化、错误处理

设计目标
========

* 所有命令的 JSON 输出遵循统一外层结构::

      { "success": bool, "command": str, "data": {...} }      # 成功
      { "success": bool, "command": str, "error": {...} }     # 失败

* 错误对象的字段::

      { "type": str, "message": str, "code": int,
        "suggestion": str, "context": {...} }

* 退出码遵循 ``wps_cli.exceptions`` 中各异常类的 ``exit_code`` 字段。
* 错误消息中的本地路径会被脱敏。
"""

from __future__ import annotations

import json
import logging
from collections.abc import Callable
from typing import Any, TypeVar

import typer
from rich.console import Console
from rich.table import Table

from wps_cli.exceptions import WpsCliError
from wps_cli.utils.path_utils import redact_path

T = TypeVar("T")

console = Console(stderr=True)
logger = logging.getLogger("wps_cli")


def _emit_json(data: Any) -> None:
    typer.echo(json.dumps(data, ensure_ascii=False, indent=2, default=str))


def output_table(title: str, headers: list[str], rows: list[list[Any]]) -> None:
    """Rich 表格输出到 stderr"""
    table = Table(title=title, show_lines=True)
    for h in headers:
        table.add_column(h, style="cyan")
    for row in rows:
        table.add_row(*[str(c) for c in row])
    console.print(table)


def output_tsv(headers: list[str], rows: list[list[Any]]) -> None:
    """TSV 格式输出到 stdout（管道友好）"""
    typer.echo("\t".join(headers))
    for row in rows:
        typer.echo("\t".join(str(c) for c in row))


def success(
    data: Any = None,
    *,
    command: str = "",
    json_mode: bool = False,
    title: str = "",
    headers: list[str] | None = None,
) -> None:
    """统一成功响应输出

    JSON 模式下输出 ``{success, command, data}``；终端模式下：
    若 ``data`` 是 list 且提供了 ``headers``，渲染为 Rich 表格；否则直接输出。
    """
    if json_mode:
        payload = {"success": True, "command": command, "data": data}
        _emit_json(payload)
        return

    if data is None:
        typer.echo("✓ 完成")
        return

    if headers and isinstance(data, list):
        output_table(title or command, headers, data)
        return

    if isinstance(data, dict):
        for k, v in data.items():
            typer.echo(f"{k}: {v}")
    elif isinstance(data, (list, tuple)):
        for item in data:
            typer.echo(str(item))
    else:
        typer.echo(str(data))


def handle_error(error: Exception, *, command: str = "", json_mode: bool = False) -> None:
    """统一错误处理

    根据异常类型确定 exit code、suggestion、context。所有路径会被脱敏。
    """
    if isinstance(error, WpsCliError):
        exit_code = error.exit_code
        err_type = type(error).__name__
        suggestion = error.suggestion
        context = error.context
    else:
        exit_code = 1
        err_type = type(error).__name__
        suggestion = ""
        context = {}

    safe_message = redact_path(str(error))
    safe_context = {k: redact_path(str(v)) for k, v in context.items()}

    logger.debug("命令 %s 异常: %s", command, error, exc_info=True)

    if json_mode:
        payload = {
            "success": False,
            "command": command,
            "error": {
                "type": err_type,
                "message": safe_message,
                "code": exit_code,
                "suggestion": suggestion,
                "context": safe_context,
            },
        }
        _emit_json(payload)
    else:
        typer.echo(f"错误 [{err_type}]: {safe_message}", err=True)
        if suggestion:
            typer.echo(f"建议: {suggestion}", err=True)
    raise typer.Exit(code=exit_code)


def make_get_service(service_cls: type[T]) -> Callable[[], T]:
    """创建惰性单例的 _get_service 工厂函数

    每个 CLI 模块调用一次，返回该模块的 _get_service 函数。内部延迟导入
    WpsComBackend，避免循环依赖。

    注意：当前实现假设单线程使用（CLI 场景）。如需在多线程环境复用，需要
    自行加锁或为每个线程创建独立 SessionManager。
    """
    _state: dict[str, Any] = {"manager": None, "service": None}

    def _get_service() -> T:
        if _state["service"] is None:
            from wps_cli.backends.wps_com import WpsComBackend
            from wps_cli.services.session_manager import SessionManager

            _state["manager"] = SessionManager(backend=WpsComBackend())
            _state["service"] = service_cls(manager=_state["manager"])
        return _state["service"]

    return _get_service
