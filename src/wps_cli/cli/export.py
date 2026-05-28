"""导出 CLI 命令"""

from pathlib import Path

import typer

from wps_cli.cli.common import do_output, handle_error, make_get_service
from wps_cli.services.export_service import ExportService

app = typer.Typer(help="格式转换与导出")

_get_service = make_get_service(ExportService)


@app.command()
def convert(
    file: str = typer.Argument(..., help="源文件路径"),
    format: str = typer.Argument(..., help="目标格式: docx/pdf/html/txt/csv/xlsx/pptx"),
    output: str = typer.Option("", "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """格式转换"""
    try:
        result = _get_service().convert(Path(file), format, Path(output) if output else None)
        do_output({"success": True, "path": str(result)}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def batch(
    glob_pattern: str = typer.Argument(..., help="文件模式，如 *.docx"),
    format: str = typer.Option(..., "--to", "-t", help="目标格式"),
    output_dir: str = typer.Option(..., "--output-dir", "-d", help="输出目录"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """批量格式转换"""
    try:
        results = _get_service().batch_convert(glob_pattern, format, Path(output_dir))
        success = sum(1 for r in results if r["status"] == "ok")
        failed = sum(1 for r in results if r["status"] == "failed")
        do_output({"success": failed == 0, "total": len(results), "success_count": success, "failed_count": failed, "results": results}, json_output)
    except Exception as e:
        handle_error(e, json_output)
