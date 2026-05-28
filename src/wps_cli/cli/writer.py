"""Writer CLI 命令"""

import typer

from wps_cli.cli.common import do_output, handle_error
from wps_cli.services.session_manager import SessionManager
from wps_cli.services.writer_service import WriterService

app = typer.Typer(help="Word 文档操作")

_manager: SessionManager | None = None
_service: WriterService | None = None


def _get_service() -> WriterService:
    global _manager, _service
    if _service is None:
        from wps_cli.backends.wps_com import WpsComBackend
        _manager = SessionManager(backend=WpsComBackend())
        _service = WriterService(manager=_manager)
    return _service


@app.command()
def new(
    output: str = typer.Option("", "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """新建空白 Word 文档"""
    try:
        from pathlib import Path
        result = _get_service().new(Path(output) if output else None)
        do_output({"success": True, "path": str(result)}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def info(
    file: str = typer.Argument(..., help="文档路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """输出文档元信息"""
    try:
        from pathlib import Path
        result = _get_service().info(Path(file))
        do_output(result, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def replace(
    file: str = typer.Argument(..., help="文档路径"),
    old: str = typer.Argument(..., help="查找文本"),
    new_text: str = typer.Argument(..., help="替换文本"),
    regex: bool = typer.Option(False, "--regex", help="正则模式"),
    case: bool = typer.Option(False, "--case", help="区分大小写"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """查找替换文本"""
    try:
        from pathlib import Path
        svc = _get_service()
        session = svc.open(Path(file))
        try:
            count = svc.text_replace(session.app, old, new_text, regex, case)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        do_output({"success": True, "replaced": count}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def count(
    file: str = typer.Argument(..., help="文档路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """统计字数、段落、页数"""
    try:
        from pathlib import Path
        svc = _get_service()
        session = svc.open(Path(file))
        result = svc.text_count(session.app)
        svc.manager.stop(session.session_id)
        do_output(result, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def table_insert(
    file: str = typer.Argument(..., help="文档路径"),
    rows: int = typer.Option(..., "--rows", "-r", help="行数"),
    cols: int = typer.Option(..., "--cols", "-c", help="列数"),
    data: str = typer.Option("", "--data", "-d", help="JSON 数据"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """插入表格"""
    try:
        import json as json_mod
        from pathlib import Path
        svc = _get_service()
        session = svc.open(Path(file))
        parsed = json_mod.loads(data) if data else None
        idx = svc.table_insert(session.app, rows, cols, parsed)
        svc.save(session.app)
        svc.manager.stop(session.session_id)
        do_output({"success": True, "table_index": idx}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def table_get(
    file: str = typer.Argument(..., help="文档路径"),
    index: int = typer.Option(1, "--index", "-i", help="表格序号"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """读取表格数据"""
    try:
        from pathlib import Path
        svc = _get_service()
        session = svc.open(Path(file), readonly=True)
        result = svc.table_get(session.app, index)
        svc.manager.stop(session.session_id)
        do_output({"success": True, "data": result}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def image_insert(
    file: str = typer.Argument(..., help="文档路径"),
    image: str = typer.Option(..., "--image", "-i", help="图片路径"),
    width: float = typer.Option(0, "--width", "-w", help="宽度（磅）"),
    height: float = typer.Option(0, "--height", "-h", help="高度（磅）"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """插入图片"""
    try:
        from pathlib import Path
        svc = _get_service()
        session = svc.open(Path(file))
        svc.image_insert(session.app, Path(image), width if width else None, height if height else None)
        svc.save(session.app)
        svc.manager.stop(session.session_id)
        do_output({"success": True}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def page_setup(
    file: str = typer.Argument(..., help="文档路径"),
    width: float = typer.Option(210, "--width", help="纸张宽度 mm"),
    height: float = typer.Option(297, "--height", help="纸张高度 mm"),
    margin_top: float = typer.Option(25.4, "--margin-top", help="上边距 mm"),
    margin_bottom: float = typer.Option(25.4, "--margin-bottom", help="下边距 mm"),
    margin_left: float = typer.Option(31.75, "--margin-left", help="左边距 mm"),
    margin_right: float = typer.Option(31.75, "--margin-right", help="右边距 mm"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """设置页面布局"""
    try:
        from pathlib import Path
        svc = _get_service()
        session = svc.open(Path(file))
        svc.page_setup(session.app, width, height, margin_top, margin_bottom, margin_left, margin_right)
        svc.save(session.app)
        svc.manager.stop(session.session_id)
        do_output({"success": True}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def export_pdf(
    file: str = typer.Argument(..., help="文档路径"),
    output: str = typer.Option("", "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """导出为 PDF"""
    try:
        from pathlib import Path
        svc = _get_service()
        session = svc.open(Path(file))
        out_path = Path(output) if output else Path(file).with_suffix(".pdf")
        svc.export_pdf(session.app, out_path)
        svc.manager.stop(session.session_id)
        do_output({"success": True, "path": str(out_path)}, json_output)
    except Exception as e:
        handle_error(e, json_output)
