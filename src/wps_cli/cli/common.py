"""CLI 公共工具：输出格式化、错误处理"""

import json
from pathlib import Path
from typing import Any, Callable, TypeVar

import typer
from rich.console import Console
from rich.table import Table

T = TypeVar("T")

console = Console(stderr=True)


def output_json(data: Any) -> None:
    """JSON 格式输出到 stdout"""
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


def handle_error(error: Exception, json_mode: bool = False) -> None:
    """统一错误处理"""
    if json_mode:
        output_json({
            "success": False,
            "error": {"type": type(error).__name__, "message": str(error)},
        })
    else:
        typer.echo(f"错误: {error}", err=True)
    raise typer.Exit(code=1)


def do_output(data: Any, json_mode: bool = False, title: str = "", headers: list[str] | None = None) -> None:
    """根据格式输出数据"""
    if json_mode:
        output_json(data)
    elif headers and isinstance(data, list):
        output_table(title, headers, data)
    else:
        typer.echo(str(data))


def make_get_service(service_cls: type[T]) -> Callable[[], T]:
    """创建惰性单例的 _get_service 工厂函数

    每个 CLI 模块调用一次，返回该模块的 _get_service 函数。
    内部延迟导入 WpsComBackend，避免循环依赖。
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
