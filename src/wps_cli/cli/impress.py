"""Impress CLI 命令"""

from pathlib import Path

import typer

from wps_cli.cli.common import do_output, handle_error, make_get_service
from wps_cli.services.impress_service import ImpressService

app = typer.Typer(help="PPT 演示文稿操作")

_get_service = make_get_service(ImpressService)


@app.command()
def new(
    output: str = typer.Option("", "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """新建空白演示文稿"""
    try:
        result = _get_service().new(Path(output) if output else None)
        do_output({"success": True, "path": str(result)}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def info(
    file: str = typer.Argument(..., help="文件路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """输出演示文稿元信息"""
    try:
        result = _get_service().info(Path(file))
        do_output(result, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def slide_list(
    file: str = typer.Argument(..., help="文件路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """列出所有幻灯片"""
    try:
        svc = _get_service()
        session = svc.manager.start("impress")
        try:
            session.app.Presentations.Open(str(Path(file)))
            result = svc.slide_list(session.app)
        finally:
            svc.manager.stop(session.session_id)
        do_output(result, json_output, headers=["index", "title", "layout"])
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def slide_add(
    file: str = typer.Argument(..., help="文件路径"),
    layout: int = typer.Option(1, "--layout", "-l", help="版式编号"),
    at: int = typer.Option(0, "--at", "-a", help="插入位置"),
    title: str = typer.Option("", "--title", "-t", help="标题文本"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """新增幻灯片"""
    try:
        svc = _get_service()
        session = svc.manager.start("impress")
        try:
            session.app.Presentations.Open(str(Path(file)))
            idx = svc.slide_add(session.app, layout, at if at else None, title)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        do_output({"success": True, "index": idx}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def slide_delete(
    file: str = typer.Argument(..., help="文件路径"),
    index: int = typer.Argument(..., help="幻灯片编号"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """删除幻灯片"""
    try:
        svc = _get_service()
        session = svc.manager.start("impress")
        try:
            session.app.Presentations.Open(str(Path(file)))
            svc.slide_delete(session.app, index)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        do_output({"success": True}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def text_set(
    file: str = typer.Argument(..., help="文件路径"),
    slide: int = typer.Option(..., "--slide", "-s", help="幻灯片编号"),
    placeholder: str = typer.Option("title", "--placeholder", "-p", help="占位符: title/body/subtitle"),
    text: str = typer.Option(..., "--text", "-t", help="文本内容"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """设置幻灯片文本"""
    try:
        svc = _get_service()
        session = svc.manager.start("impress")
        try:
            session.app.Presentations.Open(str(Path(file)))
            svc.text_set(session.app, slide, placeholder, text)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        do_output({"success": True}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def text_get(
    file: str = typer.Argument(..., help="文件路径"),
    slide: int = typer.Option(..., "--slide", "-s", help="幻灯片编号"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """提取幻灯片文本"""
    try:
        svc = _get_service()
        session = svc.manager.start("impress")
        try:
            session.app.Presentations.Open(str(Path(file)))
            result = svc.text_get(session.app, slide)
        finally:
            svc.manager.stop(session.session_id)
        do_output({"slide": slide, "text": result}, json_output)
    except Exception as e:
        handle_error(e, json_output)


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
    try:
        svc = _get_service()
        session = svc.manager.start("impress")
        try:
            session.app.Presentations.Open(str(Path(file)))
            svc.image_insert(session.app, slide, Path(image), left, top)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        do_output({"success": True}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def export_pdf(
    file: str = typer.Argument(..., help="文件路径"),
    output: str = typer.Option("", "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """导出为 PDF"""
    try:
        svc = _get_service()
        session = svc.manager.start("impress")
        try:
            session.app.Presentations.Open(str(Path(file)))
            out_path = Path(output) if output else Path(file).with_suffix(".pdf")
            svc.export_pdf(session.app, out_path)
        finally:
            svc.manager.stop(session.session_id)
        do_output({"success": True, "path": str(out_path)}, json_output)
    except Exception as e:
        handle_error(e, json_output)
