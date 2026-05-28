"""导出 CLI 命令"""

from __future__ import annotations

import typer

from wps_cli.cli.common import handle_error, make_get_service, success
from wps_cli.consts import (
    CALC_INPUT_EXTENSIONS,
    IMPRESS_INPUT_EXTENSIONS,
    WRITER_INPUT_EXTENSIONS,
)
from wps_cli.services.export_service import ExportService
from wps_cli.utils.path_utils import ensure_safe_input_path, ensure_safe_output_path

app = typer.Typer(help="格式转换与导出")

_get_service = make_get_service(ExportService)

_OFFICE_INPUT_EXTENSIONS = (
    WRITER_INPUT_EXTENSIONS | CALC_INPUT_EXTENSIONS | IMPRESS_INPUT_EXTENSIONS
)


@app.command()
def convert(
    file: str = typer.Argument(..., help="源文件路径"),
    target_format: str = typer.Argument(
        ...,
        metavar="FORMAT",
        help="目标格式: docx/pdf/html/txt/csv/xlsx/pptx",
    ),
    output: str = typer.Option("", "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """格式转换"""
    cmd = "export.convert"
    try:
        path = ensure_safe_input_path(file, allowed_extensions=_OFFICE_INPUT_EXTENSIONS)
        out_path = ensure_safe_output_path(output) if output else None
        result = _get_service().convert(path, target_format, out_path)
        success(
            {"path": str(result), "format": target_format},
            command=cmd,
            json_mode=json_output,
        )
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def batch(
    glob_pattern: str = typer.Argument(
        ..., help="文件模式（相对路径），如 *.docx 或 docs/**/*.docx"
    ),
    target_format: str = typer.Option(
        ...,
        "--to",
        "-t",
        help="目标格式",
    ),
    output_dir: str = typer.Option(..., "--output-dir", "-d", help="输出目录"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """批量格式转换

    glob 模式不允许使用绝对路径或 UNC 路径，匹配结果必须位于当前工作目录之下。
    """
    cmd = "export.batch"
    try:
        out_dir = ensure_safe_output_path(output_dir)
        results = _get_service().batch_convert(glob_pattern, target_format, out_dir)
        success_n = sum(1 for r in results if r["status"] == "ok")
        failed_n = sum(1 for r in results if r["status"] == "failed")
        success(
            {
                "total": len(results),
                "success_count": success_n,
                "failed_count": failed_n,
                "results": results,
            },
            command=cmd,
            json_mode=json_output,
        )
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)
