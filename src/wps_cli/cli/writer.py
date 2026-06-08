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


@app.command()
def merge(
    file: str = typer.Argument(..., help="模板文档路径"),
    data: str = typer.Option(..., "--data", "-d", help='JSON 数据，如 \'{"name":"张三"}\''),
    output: str = typer.Option("", "--output", "-o", help="输出路径（默认覆盖原文件）"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """模板合并：将 {{key}} 占位符替换为实际值

    设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

    Examples:
        wps writer merge template.docx --data '{"name":"张三","date":"2026-06-07"}' -o output.docx
    """
    cmd = "writer.merge"
    try:
        path = _safe_writer_input(file)
        try:
            parsed = json_mod.loads(data)
        except json_mod.JSONDecodeError as exc:
            raise ValidationError(f"--data 必须是合法 JSON: {exc}") from exc
        if not isinstance(parsed, dict):
            raise ValidationError("--data 必须是 JSON 对象（键值对）")
        if not all(isinstance(v, str) for v in parsed.values()):
            raise ValidationError("--data 中所有值必须是字符串类型")
        svc = _get_service()
        session = svc.open_document(path)
        try:
            result = svc.template_fill(session.app, parsed)
            out_path = ensure_safe_output_path(output) if output else path
            svc.save(session.app, out_path)
        finally:
            svc.manager.stop(session.session_id)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


# ── Validate 验证命令 ──
# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)


@app.command()
def validate(
    file: str = typer.Argument(..., help="文档路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """验证 Word 文档完整性

    设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

    检查项: 拼写错误、超链接有效性、字段状态、嵌入对象、字体可用性、文档结构
    """
    cmd = "writer.validate"
    try:
        path = _safe_writer_input(file)
        from wps_cli.services.validate_service import ValidateService

        svc = ValidateService(manager=_get_service().manager)
        result = svc.validate_writer(path)
        success(
            {
                "passed": result.passed,
                "file": result.file,
                "checks": result.checks,
                "issues_count": result.issues_count,
                "errors_count": result.errors_count,
                "warnings_count": result.warnings_count,
            },
            command=cmd,
            json_mode=json_output,
        )
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


# ── Refresh 刷新命令 ──
# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)


@app.command()
def refresh(
    file: str = typer.Argument(..., help="文档路径"),
    field: str = typer.Option(
        "all", "--field", "-f", help="字段类型: all/toc/page"
    ),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """刷新文档字段

    设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

    --field all   刷新所有字段和目录
    --field toc   仅刷新目录
    --field page  仅刷新 PAGE 等页码字段
    """
    cmd = "writer.refresh"
    try:
        path = _safe_writer_input(file)
        if field not in ("all", "toc", "page"):
            raise ValidationError(
                f"不支持的字段类型: {field}",
                suggestion="可选: all, toc, page",
            )
        svc = _get_service()
        session = svc.open_document(path)
        try:
            result = svc.refresh_fields(session.app, field if field != "all" else None)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


# ── 表单域与内容控件命令 ──
# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)


@app.command("formfield-list")
def formfield_list(
    file: str = typer.Argument(..., help="文档路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """列出所有表单域（旧式 FormFields）

    设计参考: iOfficeAI/OfficeCLI (Apache 2.0)
    """
    cmd = "writer.formfield_list"
    try:
        path = _safe_writer_input(file)
        svc = _get_service()
        session = svc.open_document(path, readonly=True)
        try:
            result = svc.formfield_list(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success(result, command=cmd, json_mode=json_output, headers=["index", "name", "type", "result"])
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command("formfield-get")
def formfield_get(
    file: str = typer.Argument(..., help="文档路径"),
    index: int = typer.Option(..., "--index", "-i", help="表单域序号（从1开始）"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """获取指定表单域信息

    设计参考: iOfficeAI/OfficeCLI (Apache 2.0)
    """
    cmd = "writer.formfield_get"
    try:
        path = _safe_writer_input(file)
        svc = _get_service()
        session = svc.open_document(path, readonly=True)
        try:
            result = svc.formfield_get(session.app, index)
        finally:
            svc.manager.stop(session.session_id)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command("formfield-set")
def formfield_set(
    file: str = typer.Argument(..., help="文档路径"),
    index: int = typer.Option(..., "--index", "-i", help="表单域序号（从1开始）"),
    value: str = typer.Option(..., "--value", "-v", help="新值"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """设置表单域值

    设计参考: iOfficeAI/OfficeCLI (Apache 2.0)
    """
    cmd = "writer.formfield_set"
    try:
        path = _safe_writer_input(file)
        svc = _get_service()
        session = svc.open_document(path)
        try:
            svc.formfield_set(session.app, index, value)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"index": index, "value": value}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command("contentcontrol-list")
def contentcontrol_list(
    file: str = typer.Argument(..., help="文档路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """列出所有内容控件（ContentControls）

    设计参考: iOfficeAI/OfficeCLI (Apache 2.0)
    """
    cmd = "writer.contentcontrol_list"
    try:
        path = _safe_writer_input(file)
        svc = _get_service()
        session = svc.open_document(path, readonly=True)
        try:
            result = svc.content_control_list(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success(result, command=cmd, json_mode=json_output, headers=["index", "title", "tag", "type", "text"])
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command("contentcontrol-set")
def contentcontrol_set(
    file: str = typer.Argument(..., help="文档路径"),
    index: int = typer.Option(..., "--index", "-i", help="内容控件序号（从1开始）"),
    text: str = typer.Option(..., "--text", "-t", help="新文本"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """设置内容控件文本

    设计参考: iOfficeAI/OfficeCLI (Apache 2.0)
    """
    cmd = "writer.contentcontrol_set"
    try:
        path = _safe_writer_input(file)
        svc = _get_service()
        session = svc.open_document(path)
        try:
            svc.content_control_set(session.app, index, text)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"index": index, "text": text}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


# ── 语义视图与路径定位（Phase 4）──
# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)


@app.command("view")
def view(
    file: str = typer.Argument(..., help="文档路径"),
    view_type: str = typer.Argument("summary", help="视图类型: summary/issues/outline/annotated/stats"),
    type_filter: str = typer.Option("", "--type", "-t", help="过滤问题子类型（仅对 issues 视图有效）"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """文档语义视图（参考 OfficeCLI L1 Read）

    设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

    支持五种视图:
      summary   — 文档结构摘要（标题/表格/图片）
      issues    — 文档诊断（字体/布局/图片/样式问题），可用 --type 过滤子类型
      outline   — 纯标题大纲
      annotated — 带路径标注的文档内容
      stats     — 纯数字统计
    """
    cmd = f"writer.view_{view_type}"
    try:
        path = _safe_writer_input(file)
        svc = _get_service()
        session = svc.open_document(path, readonly=True)
        try:
            if view_type == "summary":
                result = svc.summarize(session.app)
            elif view_type == "issues":
                result = svc.diagnose(session.app)
                if type_filter:
                    result = [r for r in result if r.get("subtype", "") == type_filter]
            elif view_type == "outline":
                result = svc.summarize(session.app)["headings"]
            elif view_type == "annotated":
                result = svc.annotate(session.app)
            elif view_type == "stats":
                result = svc.get_stats(session.app)
            else:
                raise ValidationError(
                    f"不支持的视图类型: {view_type}",
                    suggestion="可选: summary, issues, outline, annotated, stats",
                )
        finally:
            svc.manager.stop(session.session_id)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def get(
    file: str = typer.Argument(..., help="文档路径"),
    path: str = typer.Argument(..., help="元素路径，如 /section[1]/paragraph[3]"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """通过路径获取文档元素内容

    设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

    Examples:
        wps writer get doc.docx "/section[1]/paragraph[3]"
        wps writer get doc.docx "/section[1]/table[2]"
        wps writer get doc.docx "/body"
    """
    cmd = "writer.get"
    try:
        file_path = _safe_writer_input(file)
        svc = _get_service()
        session = svc.open_document(file_path, readonly=True)
        try:
            from wps_cli.services.path_resolver import PathResolver

            resolver = PathResolver()
            obj = resolver.resolve(session.app, "writer", path)
            content = str(obj.Text if hasattr(obj, "Text") else obj)
            result = {"path": path, "content": content}
        finally:
            svc.manager.stop(session.session_id)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)
