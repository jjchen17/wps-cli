"""Writer CLI 命令"""

from pathlib import Path

import typer

from wps_cli.cli.common import do_output, handle_error, make_get_service
from wps_cli.services.style_engine import StyleEngine
from wps_cli.services.writer_service import WriterService

app = typer.Typer(help="Word 文档操作")

_get_service = make_get_service(WriterService)


@app.command()
def new(
    output: str = typer.Option("", "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """新建空白 Word 文档"""
    try:
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
        result = _get_service().info(Path(file))
        do_output(result, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def replace(
    file: str = typer.Argument(..., help="文档路径"),
    old: str = typer.Argument(..., help="查找文本"),
    new_text: str = typer.Argument(..., help="替换文本"),
    wildcard: bool = typer.Option(False, "--wildcard", "-w", help="通配符模式（* ? [abc]）"),
    case: bool = typer.Option(False, "--case", help="区分大小写"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """查找替换文本"""
    try:
        svc = _get_service()
        session = svc.open(Path(file))
        try:
            count = svc.text_replace(session.app, old, new_text, wildcard, case)
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
        svc = _get_service()
        session = svc.open(Path(file))
        try:
            result = svc.text_count(session.app)
        finally:
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
        svc = _get_service()
        session = svc.open(Path(file))
        try:
            parsed = json_mod.loads(data) if data else None
            idx = svc.table_insert(session.app, rows, cols, parsed)
            svc.save(session.app)
        finally:
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
        svc = _get_service()
        session = svc.open(Path(file), readonly=True)
        try:
            result = svc.table_get(session.app, index)
        finally:
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
        svc = _get_service()
        session = svc.open(Path(file))
        try:
            svc.image_insert(session.app, Path(image), width if width != 0 else None, height if height != 0 else None)
            svc.save(session.app)
        finally:
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
        svc = _get_service()
        session = svc.open(Path(file))
        try:
            svc.page_setup(session.app, width, height, margin_top, margin_bottom, margin_left, margin_right)
            svc.save(session.app)
        finally:
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
        svc = _get_service()
        session = svc.open(Path(file))
        try:
            out_path = Path(output) if output else Path(file).with_suffix(".pdf")
            svc.export_pdf(session.app, out_path)
        finally:
            svc.manager.stop(session.session_id)
        do_output({"success": True, "path": str(out_path)}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command("style-apply")
def style_apply(
    file: str = typer.Argument(..., help="文档路径"),
    preset: str = typer.Argument(..., help="样式预设名称"),
    list_presets: bool = typer.Option(False, "--list", "-l", help="列出所有可用预设"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """应用预设样式到文档当前选区"""
    engine = StyleEngine()
    if list_presets:
        do_output({"presets": engine.list_presets()}, json_output)
        return
    try:
        svc = _get_service()
        session = svc.open(Path(file))
        try:
            engine.apply_preset(session.app, preset)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        do_output({"success": True, "preset": preset}, json_output)
    except Exception as e:
        handle_error(e, json_output)
