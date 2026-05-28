"""PDF CLI 命令"""

import typer

from wps_cli.cli.common import do_output, handle_error
from wps_cli.services.pdf_service import PdfService
from wps_cli.services.session_manager import SessionManager

app = typer.Typer(help="PDF 文档操作")

_manager: SessionManager | None = None
_service: PdfService | None = None


def _get_service() -> PdfService:
    global _manager, _service
    if _service is None:
        from wps_cli.backends.wps_com import WpsComBackend
        _manager = SessionManager(backend=WpsComBackend())
        _service = PdfService(manager=_manager)
    return _service


@app.command()
def info(
    file: str = typer.Argument(..., help="PDF 文件路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """输出 PDF 元信息"""
    try:
        from pathlib import Path
        result = _get_service().info(Path(file))
        do_output(result, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def merge(
    files: list[str] = typer.Argument(..., help="要合并的 PDF 文件"),
    output: str = typer.Option(..., "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """合并多个 PDF"""
    try:
        from pathlib import Path
        result = _get_service().merge([Path(f) for f in files], Path(output))
        do_output({"success": True, "path": str(result), "merged": len(files)}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def extract_pages(
    file: str = typer.Argument(..., help="PDF 文件路径"),
    pages: str = typer.Argument(..., help="页码范围，如 1-3,5,7-9"),
    output: str = typer.Option(..., "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """提取指定页面"""
    try:
        from pathlib import Path
        result = _get_service().extract_pages(Path(file), pages, Path(output))
        do_output({"success": True, "path": str(result), "pages": pages}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def split(
    file: str = typer.Argument(..., help="PDF 文件路径"),
    every: int = typer.Option(..., "--every", "-e", help="每 N 页拆分"),
    output_dir: str = typer.Option(..., "--output-dir", "-d", help="输出目录"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """按每 N 页拆分"""
    try:
        from pathlib import Path
        results = _get_service().split(Path(file), every, Path(output_dir))
        do_output({"success": True, "parts": len(results), "files": [str(r) for r in results]}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def watermark(
    file: str = typer.Argument(..., help="PDF 文件路径"),
    text: str = typer.Argument(..., help="水印文字"),
    output: str = typer.Option(..., "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """添加文字水印"""
    try:
        from pathlib import Path
        result = _get_service().watermark(Path(file), text, Path(output))
        do_output({"success": True, "path": str(result)}, json_output)
    except Exception as e:
        handle_error(e, json_output)
