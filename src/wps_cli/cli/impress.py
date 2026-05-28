"""Impress CLI 命令"""

from __future__ import annotations

from pathlib import Path

import typer

from wps_cli.cli.common import handle_error, make_get_service, success
from wps_cli.services.impress_service import ImpressService
from wps_cli.utils.path_utils import ensure_safe_input_path, ensure_safe_output_path

app = typer.Typer(help="PPT 演示文稿操作")

_get_service = make_get_service(ImpressService)


def _open_pres(svc: ImpressService, path: Path, readonly: bool = False):
    session = svc.manager.start("impress")
    try:
        session.app.Presentations.Open(str(path), ReadOnly=readonly)
        return session
    except Exception:
        svc.manager.stop(session.session_id)
        raise


@app.command()
def new(
    output: str = typer.Option("", "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """新建空白演示文稿"""
    cmd = "impress.new"
    try:
        out_path = ensure_safe_output_path(output) if output else None
        result = _get_service().new(out_path)
        success({"path": str(result)}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def info(
    file: str = typer.Argument(..., help="文件路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """输出演示文稿元信息"""
    cmd = "impress.info"
    try:
        path = ensure_safe_input_path(file)
        result = _get_service().info(path)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def slide_list(
    file: str = typer.Argument(..., help="文件路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """列出所有幻灯片"""
    cmd = "impress.slide_list"
    try:
        path = ensure_safe_input_path(file)
        svc = _get_service()
        session = _open_pres(svc, path, readonly=True)
        try:
            result = svc.slide_list(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success(result, command=cmd, json_mode=json_output, headers=["index", "title", "layout"])
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def slide_add(
    file: str = typer.Argument(..., help="文件路径"),
    layout: int = typer.Option(1, "--layout", "-l", help="版式编号"),
    at: int = typer.Option(0, "--at", "-a", help="插入位置"),
    title: str = typer.Option("", "--title", "-t", help="标题文本"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """新增幻灯片"""
    cmd = "impress.slide_add"
    try:
        path = ensure_safe_input_path(file)
        svc = _get_service()
        session = _open_pres(svc, path)
        try:
            idx = svc.slide_add(session.app, layout, at if at else None, title)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"index": idx}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def slide_delete(
    file: str = typer.Argument(..., help="文件路径"),
    index: int = typer.Argument(..., help="幻灯片编号"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """删除幻灯片"""
    cmd = "impress.slide_delete"
    try:
        path = ensure_safe_input_path(file)
        svc = _get_service()
        session = _open_pres(svc, path)
        try:
            svc.slide_delete(session.app, index)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"deleted": index}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def text_set(
    file: str = typer.Argument(..., help="文件路径"),
    slide: int = typer.Option(..., "--slide", "-s", help="幻灯片编号"),
    placeholder: str = typer.Option(
        "title", "--placeholder", "-p", help="占位符: title/body/subtitle"
    ),
    text: str = typer.Option(..., "--text", "-t", help="文本内容"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """设置幻灯片文本"""
    cmd = "impress.text_set"
    try:
        path = ensure_safe_input_path(file)
        svc = _get_service()
        session = _open_pres(svc, path)
        try:
            svc.text_set(session.app, slide, placeholder, text)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"slide": slide, "placeholder": placeholder}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def text_get(
    file: str = typer.Argument(..., help="文件路径"),
    slide: int = typer.Option(..., "--slide", "-s", help="幻灯片编号"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """提取幻灯片文本"""
    cmd = "impress.text_get"
    try:
        path = ensure_safe_input_path(file)
        svc = _get_service()
        session = _open_pres(svc, path, readonly=True)
        try:
            result = svc.text_get(session.app, slide)
        finally:
            svc.manager.stop(session.session_id)
        success({"slide": slide, "text": result}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def image_insert(
    file: str = typer.Argument(..., help="文件路径"),
    slide: int = typer.Option(..., "--slide", "-s", help="幻灯片编号"),
    image: str = typer.Option(..., "--image", "-i", help="图片路径"),
    left: float = typer.Option(100, "--left", help="左边距"),
    top: float = typer.Option(100, "--top", help="上边距"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """插入图片"""
    cmd = "impress.image_insert"
    try:
        path = ensure_safe_input_path(file)
        image_path = ensure_safe_input_path(image)
        svc = _get_service()
        session = _open_pres(svc, path)
        try:
            svc.image_insert(session.app, slide, image_path, left, top)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"slide": slide, "image": str(image_path)}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def export_pdf(
    file: str = typer.Argument(..., help="文件路径"),
    output: str = typer.Option("", "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """导出为 PDF"""
    cmd = "impress.export_pdf"
    try:
        path = ensure_safe_input_path(file)
        out_path = ensure_safe_output_path(output) if output else path.with_suffix(".pdf")
        svc = _get_service()
        session = _open_pres(svc, path, readonly=True)
        try:
            svc.export_pdf(session.app, out_path)
        finally:
            svc.manager.stop(session.session_id)
        success({"path": str(out_path)}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)
