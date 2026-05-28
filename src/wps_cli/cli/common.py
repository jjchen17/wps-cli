"""CLI 公共工具：输出格式化、错误处理"""

import json
from typing import Any

import typer
from rich.console import Console
from rich.table import Table

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
