"""PDF CLI 命令"""

from __future__ import annotations

import typer

from wps_cli.cli.common import handle_error, make_get_service, success
from wps_cli.services.pdf_service import PdfService
from wps_cli.utils.path_utils import ensure_safe_input_path, ensure_safe_output_path

app = typer.Typer(help="PDF 文档操作")

_get_service = make_get_service(PdfService)


@app.command()
def info(
    file: str = typer.Argument(..., help="PDF 文件路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """输出 PDF 元信息"""
    cmd = "pdf.info"
    try:
        path = ensure_safe_input_path(file)
        result = _get_service().info(path)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def merge(
    files: list[str] = typer.Argument(..., help="要合并的 PDF 文件"),
    output: str = typer.Option(..., "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """合并多个 PDF"""
    cmd = "pdf.merge"
    try:
        paths = [ensure_safe_input_path(f) for f in files]
        out_path = ensure_safe_output_path(output)
        result = _get_service().merge(paths, out_path)
        success(
            {"path": str(result), "merged": len(paths)},
            command=cmd,
            json_mode=json_output,
        )
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def extract_pages(
    file: str = typer.Argument(..., help="PDF 文件路径"),
    pages: str = typer.Argument(..., help="页码范围，如 1-3,5,7-9"),
    output: str = typer.Option(..., "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """提取指定页面"""
    cmd = "pdf.extract_pages"
    try:
        path = ensure_safe_input_path(file)
        out_path = ensure_safe_output_path(output)
        result = _get_service().extract_pages(path, pages, out_path)
        success(
            {"path": str(result), "pages": pages},
            command=cmd,
            json_mode=json_output,
        )
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def split(
    file: str = typer.Argument(..., help="PDF 文件路径"),
    every: int = typer.Option(..., "--every", "-e", help="每 N 页拆分"),
    output_dir: str = typer.Option(..., "--output-dir", "-d", help="输出目录"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """按每 N 页拆分"""
    cmd = "pdf.split"
    try:
        path = ensure_safe_input_path(file)
        out_dir = ensure_safe_output_path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        results = _get_service().split(path, every, out_dir)
        success(
            {"parts": len(results), "files": [str(r) for r in results]},
            command=cmd,
            json_mode=json_output,
        )
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def watermark(
    file: str = typer.Argument(..., help="PDF 文件路径"),
    text: str = typer.Argument(..., help="水印文字"),
    output: str = typer.Option(..., "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """添加文字水印"""
    cmd = "pdf.watermark"
    try:
        path = ensure_safe_input_path(file)
        out_path = ensure_safe_output_path(output)
        result = _get_service().watermark(path, text, out_path)
        success({"path": str(result)}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)
