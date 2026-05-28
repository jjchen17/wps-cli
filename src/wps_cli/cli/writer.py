"""Writer CLI 命令"""

from __future__ import annotations

import json as json_mod
import re

import typer

from wps_cli.cli.common import handle_error, make_get_service, success
from wps_cli.consts import MAX_REPLACE_TEXT_LEN, WRITER_INPUT_EXTENSIONS
from wps_cli.exceptions import ValidationError
from wps_cli.services.style_engine import StyleEngine
from wps_cli.services.writer_service import WriterService
from wps_cli.utils.path_utils import ensure_safe_input_path, ensure_safe_output_path

app = typer.Typer(help="Word 文档操作")

_get_service = make_get_service(WriterService)


def _safe_writer_input(file: str):
    return ensure_safe_input_path(file, allowed_extensions=WRITER_INPUT_EXTENSIONS)


@app.command()
def new(
    output: str = typer.Option("", "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """新建空白 Word 文档"""
    cmd = "writer.new"
    try:
        out_path = ensure_safe_output_path(output) if output else None
        result = _get_service().new(out_path)
        success({"path": str(result)}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def info(
    file: str = typer.Argument(..., help="文档路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """输出文档元信息"""
    cmd = "writer.info"
    try:
        path = _safe_writer_input(file)
        result = _get_service().info(path)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


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
    cmd = "writer.replace"
    try:
        path = _safe_writer_input(file)
        if len(old) > MAX_REPLACE_TEXT_LEN or len(new_text) > MAX_REPLACE_TEXT_LEN:
            raise ValidationError(f"查找/替换文本长度不能超过 {MAX_REPLACE_TEXT_LEN} 字符")
        if not old:
            raise ValidationError("查找文本不能为空")
        if wildcard and re.search(r"\\[1-9]", new_text):
            raise ValidationError(
                "通配符替换模式中不允许使用反向引用 (\\1-\\9)，可能导致内容指数级膨胀"
            )
        svc = _get_service()
        session = svc.open_document(path)
        try:
            count = svc.text_replace(session.app, old, new_text, wildcard, case)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"replaced": count}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def count(
    file: str = typer.Argument(..., help="文档路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """统计字数、段落、页数"""
    cmd = "writer.count"
    try:
        path = _safe_writer_input(file)
        svc = _get_service()
        session = svc.open_document(path, readonly=True)
        try:
            result = svc.text_count(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def table_insert(
    file: str = typer.Argument(..., help="文档路径"),
    rows: int = typer.Option(..., "--rows", "-r", help="行数"),
    cols: int = typer.Option(..., "--cols", "-c", help="列数"),
    data: str = typer.Option("", "--data", "-d", help="JSON 数据，二维数组"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """插入表格"""
    cmd = "writer.table_insert"
    try:
        path = _safe_writer_input(file)
        if rows <= 0 or cols <= 0:
            raise ValidationError("rows / cols 必须为正整数")
        try:
            parsed = json_mod.loads(data) if data else None
        except json_mod.JSONDecodeError as exc:
            raise ValidationError(f"--data 必须是合法 JSON: {exc}") from exc
        svc = _get_service()
        session = svc.open_document(path)
        try:
            idx = svc.table_insert(session.app, rows, cols, parsed)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"table_index": idx}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def table_get(
    file: str = typer.Argument(..., help="文档路径"),
    index: int = typer.Option(1, "--index", "-i", help="表格序号"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """读取表格数据"""
    cmd = "writer.table_get"
    try:
        path = _safe_writer_input(file)
        svc = _get_service()
        session = svc.open_document(path, readonly=True)
        try:
            result = svc.table_get(session.app, index)
        finally:
            svc.manager.stop(session.session_id)
        success({"data": result}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def image_insert(
    file: str = typer.Argument(..., help="文档路径"),
    image: str = typer.Option(..., "--image", "-i", help="图片路径"),
    width: float = typer.Option(0, "--width", "-w", help="宽度（磅）"),
    height: float = typer.Option(0, "--height", "-h", help="高度（磅）"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """插入图片"""
    cmd = "writer.image_insert"
    try:
        path = _safe_writer_input(file)
        image_path = ensure_safe_input_path(image)
        svc = _get_service()
        session = svc.open_document(path)
        try:
            svc.image_insert(
                session.app,
                image_path,
                width if width != 0 else None,
                height if height != 0 else None,
            )
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"image": str(image_path)}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


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
    cmd = "writer.page_setup"
    try:
        path = _safe_writer_input(file)
        svc = _get_service()
        session = svc.open_document(path)
        try:
            svc.page_setup(
                session.app, width, height, margin_top, margin_bottom, margin_left, margin_right
            )
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success(None, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def export_pdf(
    file: str = typer.Argument(..., help="文档路径"),
    output: str = typer.Option("", "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """导出为 PDF"""
    cmd = "writer.export_pdf"
    try:
        path = _safe_writer_input(file)
        out_path = ensure_safe_output_path(output) if output else path.with_suffix(".pdf")
        svc = _get_service()
        session = svc.open_document(path, readonly=True)
        try:
            svc.export_pdf(session.app, out_path)
        finally:
            svc.manager.stop(session.session_id)
        success({"path": str(out_path)}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command("style-apply")
def style_apply(
    file: str = typer.Argument(..., help="文档路径"),
    preset: str = typer.Argument(..., help="样式预设名称"),
    list_presets: bool = typer.Option(False, "--list", "-l", help="列出所有可用预设"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """应用预设样式到文档当前选区"""
    cmd = "writer.style_apply"
    engine = StyleEngine()
    if list_presets:
        success({"presets": engine.list_presets()}, command=cmd, json_mode=json_output)
        return
    try:
        path = _safe_writer_input(file)
        svc = _get_service()
        session = svc.open_document(path)
        try:
            engine.apply_preset(session.app, preset)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"preset": preset}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)
