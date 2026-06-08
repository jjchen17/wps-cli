# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""Batch 批量命令执行

在一个 COM 会话中顺序执行多条命令，默认 continue-on-error。
支持驻留模式转发，自动复用已打开文件。

命令格式::

    [
      {
        "command": "writer.replace",
        "params": {"file": "a.docx", "old": "foo", "new": "bar"}
      },
      {
        "command": "calc.cell-set",
        "params": {"file": "b.xlsx", "ref": "A1", "value": 100, "sheet": "Sheet1"}
      }
    ]
"""

from __future__ import annotations

import json as json_mod
import logging
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, ClassVar

import typer

from wps_cli.cli.common import handle_error, success
from wps_cli.exceptions import ValidationError
from wps_cli.utils.path_utils import ensure_safe_input_path, ensure_safe_output_path

app = typer.Typer(help="批量命令执行", invoke_without_command=True)
logger = logging.getLogger("wps_cli.batch")


# ── 应用类型检测 ──

_WRITER_EXTS = {".doc", ".docx", ".wps", ".rtf", ".txt", ".html"}
_CALC_EXTS = {".xls", ".xlsx", ".xlsm", ".et", ".csv"}
_IMPRESS_EXTS = {".ppt", ".pptx", ".pps", ".dps"}


def _detect_app_type(path: str | Path) -> str:
    """根据文件扩展名判断应用类型"""
    ext = Path(path).suffix.lower()
    if ext in _WRITER_EXTS:
        return "writer"
    if ext in _CALC_EXTS:
        return "calc"
    if ext in _IMPRESS_EXTS:
        return "impress"
    raise ValidationError(f"无法识别文件类型: {ext}")


# ── 驻留模式检测 ──

_RESIDENT_PORT = 9123
_RESIDENT_HOST = "127.0.0.1"


def _is_resident_running() -> bool:
    """检测驻留进程是否在运行"""
    import socket
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(0.5)
        result = sock.connect_ex((_RESIDENT_HOST, _RESIDENT_PORT))
        sock.close()
        return result == 0
    except Exception:
        return False


def _forward_to_resident(commands: list[dict], stop_on_error: bool = False) -> dict:
    """将 batch 请求转发到驻留进程"""
    import http.client
    import json as json_mod

    body = json_mod.dumps({
        "commands": commands,
        "stop_on_error": stop_on_error,
    }).encode("utf-8")

    conn = http.client.HTTPConnection(_RESIDENT_HOST, _RESIDENT_PORT, timeout=120)
    try:
        conn.request("POST", "/batch", body=body)
        resp = conn.getresponse()
        data = json_mod.loads(resp.read().decode("utf-8"))
        return data
    finally:
        conn.close()


# ── Batch 执行器 ──


@dataclass
class BatchExecutor:
    """批量命令执行器

    管理 COM 会话生命周期，缓存已打开文件避免重复 open/save。
    """

    _open_sessions: dict[str, tuple[str, Any, str]] = field(default_factory=dict)
    # path -> (app_type, app, session_id)

    _manager: Any = None
    _writer: Any = None
    _calc: Any = None
    _impress: Any = None
    _pdf: Any = None
    _export: Any = None

    def _ensure_services(self) -> None:
        """延迟初始化服务（避免导入时触发 COM）"""
        if self._manager is not None:
            return
        from wps_cli.backends.wps_com import WpsComBackend
        from wps_cli.services.calc_service import CalcService
        from wps_cli.services.export_service import ExportService
        from wps_cli.services.impress_service import ImpressService
        from wps_cli.services.pdf_service import PdfService
        from wps_cli.services.session_manager import SessionManager
        from wps_cli.services.writer_service import WriterService

        self._manager = SessionManager(backend=WpsComBackend())
        self._writer = WriterService(manager=self._manager)
        self._calc = CalcService(manager=self._manager)
        self._impress = ImpressService(manager=self._manager)
        self._pdf = PdfService(manager=self._manager)
        self._export = ExportService(manager=self._manager)

    def _open_file(self, path: str, app_type: str) -> Any:
        """打开文件，如果已打开则复用会话"""
        abs_path = str(ensure_safe_input_path(path))

        if abs_path in self._open_sessions:
            return self._open_sessions[abs_path][1]

        session = self._manager.start(app_type)
        try:
            if app_type == "writer":
                session.app.Documents.Open(
                    abs_path,
                    ConfirmConversions=False,
                    ReadOnly=False,
                    AddToRecentFiles=False,
                )
            elif app_type == "calc":
                session.app.Workbooks.Open(
                    abs_path,
                    UpdateLinks=0,
                    ReadOnly=False,
                )
            elif app_type == "impress":
                session.app.Presentations.Open(abs_path, ReadOnly=False)
            else:
                raise ValidationError(f"不支持的应用类型: {app_type}")
        except Exception:
            self._manager.stop(session.session_id)
            raise

        self._open_sessions[abs_path] = (app_type, session.app, session.session_id)
        return session.app

    def _get_app_for_file(self, file_path: str, app_type: str | None = None) -> Any:
        """获取文件对应的 COM app 对象（自动打开/复用）"""
        if app_type is None:
            app_type = _detect_app_type(file_path)
        return self._open_file(file_path, app_type)

    def _save_and_close_all(self) -> None:
        """保存并关闭所有打开的文件"""
        for abs_path, (app_type, app, session_id) in list(self._open_sessions.items()):
            try:
                if app_type == "writer":
                    doc = app.ActiveDocument
                    if doc:
                        doc.Save()
                        doc.Close()
                elif app_type == "calc":
                    wb = app.ActiveWorkbook
                    if wb:
                        wb.Save()
                        wb.Close()
                elif app_type == "impress":
                    pres = app.ActivePresentation
                    if pres:
                        pres.Save()
                        pres.Close()
            except Exception:
                logger.debug("保存/关闭文件时出错: %s", abs_path, exc_info=True)
            try:
                self._manager.stop(session_id)
            except Exception:
                pass
        self._open_sessions.clear()

    def _cleanup(self) -> None:
        """清理所有资源"""
        try:
            self._save_and_close_all()
        except Exception:
            pass
        try:
            if self._manager:
                self._manager.stop_all()
        except Exception:
            pass

    # ── 命令处理器映射 ──

    COMMAND_HANDLER_MAP: ClassVar[dict[str, str]] = {
        # Writer
        "writer.info": "_h_writer_info",
        "writer.replace": "_h_writer_replace",
        "writer.count": "_h_writer_count",
        "writer.table-get": "_h_writer_table_get",
        "writer.table-insert": "_h_writer_table_insert",
        "writer.image-insert": "_h_writer_image_insert",
        "writer.page-setup": "_h_writer_page_setup",
        "writer.style-apply": "_h_writer_style_apply",
        "writer.export-pdf": "_h_writer_export_pdf",
        "writer.new": "_h_writer_new",
        # Calc
        "calc.info": "_h_calc_info",
        "calc.cell-get": "_h_calc_cell_get",
        "calc.cell-set": "_h_calc_cell_set",
        "calc.cell-range": "_h_calc_range_get",
        "calc.cell-formula": "_h_calc_cell_formula",
        "calc.chart-create": "_h_calc_chart_create",
        "calc.sort": "_h_calc_sort",
        "calc.export-csv": "_h_calc_export_csv",
        "calc.sheet-list": "_h_calc_sheet_list",
        "calc.new": "_h_calc_new",
        # Impress
        "impress.info": "_h_impress_info",
        "impress.slide-list": "_h_impress_slide_list",
        "impress.slide-add": "_h_impress_slide_add",
        "impress.slide-delete": "_h_impress_slide_delete",
        "impress.text-get": "_h_impress_text_get",
        "impress.text-set": "_h_impress_text_set",
        "impress.image-insert": "_h_impress_image_insert",
        "impress.export-pdf": "_h_impress_export_pdf",
        "impress.new": "_h_impress_new",
        # PDF
        "pdf.info": "_h_pdf_info",
        "pdf.merge": "_h_pdf_merge",
        "pdf.split": "_h_pdf_split",
        "pdf.watermark": "_h_pdf_watermark",
        "pdf.extract-pages": "_h_pdf_extract_pages",
        # Export
        "export.convert": "_h_export_convert",
        "export.batch": "_h_export_batch",
    }

    def _resolve_handler(self, command: str):
        """将命令字符串解析为处理器方法"""
        handler_name = self.COMMAND_HANDLER_MAP.get(command)
        if handler_name is None:
            raise ValidationError(
                f"未知命令: {command}，支持的命令: {', '.join(sorted(self.COMMAND_HANDLER_MAP))}"
            )
        return getattr(self, handler_name)

    # ── Writer 处理器 ──

    def _h_writer_info(self, params: dict) -> dict:
        file = ensure_safe_input_path(params["file"])
        return self._writer.info(file)

    def _h_writer_replace(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "writer")
        count = self._writer.text_replace(
            app,
            params["old"],
            params["new"],
            params.get("wildcard", False),
            params.get("case", False),
        )
        return {"replaced": count}

    def _h_writer_count(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "writer")
        return self._writer.text_count(app)

    def _h_writer_table_get(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "writer")
        data = self._writer.table_get(app, params.get("index", 1))
        return {"data": data}

    def _h_writer_table_insert(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "writer")
        idx = self._writer.table_insert(
            app, params["rows"], params["cols"], params.get("data")
        )
        return {"table_index": idx}

    def _h_writer_image_insert(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "writer")
        image = ensure_safe_input_path(params["image"])
        self._writer.image_insert(
            app, image,
            width=params.get("width"),
            height=params.get("height"),
        )
        return {"status": "ok"}

    def _h_writer_page_setup(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "writer")
        kwargs = {
            k: v for k, v in params.items()
            if k in ("width_mm", "height_mm", "margin_top", "margin_bottom",
                      "margin_left", "margin_right", "orientation")
        }
        self._writer.page_setup(app, **kwargs)
        return {"status": "ok"}

    def _h_writer_style_apply(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "writer")
        from wps_cli.services.style_engine import StyleEngine
        engine = StyleEngine()
        engine.apply_preset(app, params.get("preset", ""))
        return {"preset": params.get("preset", "")}

    def _h_writer_export_pdf(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "writer")
        output = ensure_safe_output_path(params["output"])
        result = self._writer.export_pdf(app, output)
        return {"output": str(result)}

    def _h_writer_new(self, params: dict) -> dict:
        output = ensure_safe_output_path(params["output"]) if "output" in params else None
        result = self._writer.new(output)
        return {"path": str(result)}

    # ── Calc 处理器 ──

    def _h_calc_info(self, params: dict) -> dict:
        file = ensure_safe_input_path(params["file"])
        return self._calc.info(file)

    def _h_calc_cell_get(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "calc")
        value = self._calc.cell_get(app, params["ref"], params.get("sheet"))
        return {"value": value}

    def _h_calc_cell_set(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "calc")
        self._calc.cell_set(app, params["ref"], params["value"], params.get("sheet"))
        return {"status": "ok"}

    def _h_calc_range_get(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "calc")
        data = self._calc.range_get(app, params["ref"], params.get("sheet"))
        return {"data": data}

    def _h_calc_cell_formula(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "calc")
        self._calc.cell_formula(app, params["ref"], params["formula"], params.get("sheet"))
        return {"status": "ok"}

    def _h_calc_chart_create(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "calc")
        idx = self._calc.chart_create(
            app,
            params["data_range"],
            params.get("chart_type", "bar"),
            params.get("title", ""),
            params.get("sheet"),
        )
        return {"chart_index": idx}

    def _h_calc_sort(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "calc")
        self._calc.sort(
            app,
            params.get("range_ref", ""),
            params["by_col"],
            params.get("order", "asc"),
            params.get("sheet"),
        )
        return {"status": "ok"}

    def _h_calc_export_csv(self, params: dict) -> dict:
        file = ensure_safe_input_path(params["file"])
        output = ensure_safe_output_path(params["output"])
        result = self._export.convert(file, "csv", output)
        return {"output": str(result)}

    def _h_calc_sheet_list(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "calc")
        sheets = self._calc.sheet_list(app)
        return {"sheets": sheets}

    def _h_calc_new(self, params: dict) -> dict:
        output = ensure_safe_output_path(params["output"]) if "output" in params else None
        result = self._calc.new(output)
        return {"path": str(result)}

    # ── Impress 处理器 ──

    def _h_impress_info(self, params: dict) -> dict:
        file = ensure_safe_input_path(params["file"])
        return self._impress.info(file)

    def _h_impress_slide_list(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "impress")
        slides = self._impress.slide_list(app)
        return {"slides": slides}

    def _h_impress_slide_add(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "impress")
        idx = self._impress.slide_add(
            app,
            params.get("layout", 1),
            params.get("at"),
            params.get("title", ""),
        )
        return {"slide_index": idx}

    def _h_impress_slide_delete(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "impress")
        self._impress.slide_delete(app, params["index"])
        return {"status": "ok"}

    def _h_impress_text_get(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "impress")
        text = self._impress.text_get(app, params["slide_idx"])
        return {"text": text}

    def _h_impress_text_set(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "impress")
        self._impress.text_set(
            app,
            params["slide_idx"],
            params.get("placeholder", "title"),
            params.get("text", ""),
        )
        return {"status": "ok"}

    def _h_impress_image_insert(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "impress")
        image = ensure_safe_input_path(params["image"])
        self._impress.image_insert(
            app,
            params["slide_idx"],
            image,
            width=params.get("width"),
            height=params.get("height"),
        )
        return {"status": "ok"}

    def _h_impress_export_pdf(self, params: dict) -> dict:
        app = self._get_app_for_file(params["file"], "impress")
        output = ensure_safe_output_path(params["output"])
        result = self._impress.export_pdf(app, output)
        return {"output": str(result)}

    def _h_impress_new(self, params: dict) -> dict:
        output = ensure_safe_output_path(params["output"]) if "output" in params else None
        result = self._impress.new(output)
        return {"path": str(result)}

    # ── PDF 处理器 ──

    def _h_pdf_info(self, params: dict) -> dict:
        file = ensure_safe_input_path(params["file"])
        return self._pdf.info(file)

    def _h_pdf_merge(self, params: dict) -> dict:
        files = [ensure_safe_input_path(f) for f in params["files"]]
        output = ensure_safe_output_path(params["output"])
        result = self._pdf.merge(files, output)
        return {"output": str(result)}

    def _h_pdf_split(self, params: dict) -> dict:
        file = ensure_safe_input_path(params["file"])
        output_dir = Path(params["output_dir"])
        output_dir.mkdir(parents=True, exist_ok=True)
        result = self._pdf.split(file, params["every"], output_dir)
        return {"files": [str(p) for p in result]}

    def _h_pdf_watermark(self, params: dict) -> dict:
        file = ensure_safe_input_path(params["file"])
        output = ensure_safe_output_path(params["output"])
        result = self._pdf.watermark(file, params["text"], output)
        return {"output": str(result)}

    def _h_pdf_extract_pages(self, params: dict) -> dict:
        file = ensure_safe_input_path(params["file"])
        output = ensure_safe_output_path(params["output"])
        result = self._pdf.extract_pages(file, params["pages"], output)
        return {"output": str(result)}

    # ── Export 处理器 ──

    def _h_export_convert(self, params: dict) -> dict:
        input_path = ensure_safe_input_path(params["file"])
        output_path = ensure_safe_output_path(params["output"]) if "output" in params else None
        result = self._export.convert(input_path, params["format"], output_path)
        return {"output": str(result)}

    def _h_export_batch(self, params: dict) -> dict:
        from wps_cli.utils.path_utils import ensure_safe_glob
        files = ensure_safe_glob(params["pattern"])
        output_dir = ensure_safe_output_path(params["output_dir"])
        results = self._export.batch_convert(files, params["format"], output_dir)
        return {"converted": len(results), "files": [str(p) for p in results]}

    # ── 主执行入口 ──

    def execute_batch(
        self, commands: list[dict], *, stop_on_error: bool = False
    ) -> dict:
        """执行批量命令序列

        Args:
            commands: 命令列表，每项为 {"command": "calc.cell-set", "params": {...}}
            stop_on_error: True 时遇到错误立即停止，False 时继续执行后续命令

        Returns:
            {"steps": [...], "summary": {"total": N, "succeeded": N, "failed": N}}
        """
        self._ensure_services()
        steps: list[dict] = []
        succeeded = 0
        failed = 0

        try:
            for i, cmd in enumerate(commands):
                command_name = cmd.get("command", "")
                params = cmd.get("params", {}) or {}

                if not command_name:
                    step = {
                        "index": i,
                        "success": False,
                        "command": "",
                        "error": {"type": "ValidationError", "message": "缺少 command 字段"},
                    }
                    steps.append(step)
                    failed += 1
                    if stop_on_error:
                        break
                    continue

                try:
                    handler = self._resolve_handler(command_name)
                    result = handler(params)
                    steps.append({
                        "index": i,
                        "success": True,
                        "command": command_name,
                        "result": result,
                    })
                    succeeded += 1
                except typer.Exit:
                    raise
                except Exception as exc:
                    err_type = type(exc).__name__
                    err_msg = str(exc)
                    from wps_cli.utils.path_utils import redact_path
                    safe_msg = redact_path(err_msg)

                    step = {
                        "index": i,
                        "success": False,
                        "command": command_name,
                        "error": {"type": err_type, "message": safe_msg},
                    }
                    # 如果是 WpsCliError，提取额外上下文
                    from wps_cli.exceptions import WpsCliError
                    if isinstance(exc, WpsCliError):
                        step["error"]["code"] = exc.exit_code
                        if exc.suggestion:
                            step["error"]["suggestion"] = exc.suggestion

                    steps.append(step)
                    failed += 1
                    logger.debug("命令 %d (%s) 执行失败: %s", i, command_name, safe_msg)

                    if stop_on_error:
                        break
        finally:
            self._cleanup()

        return {
            "steps": steps,
            "summary": {
                "total": len(commands),
                "succeeded": succeeded,
                "failed": failed,
            },
        }


# ── CLI 命令 ──


@app.callback()
def batch(
    commands: str | None = typer.Option(
        None, "--commands", "-c",
        help="JSON 命令数组字符串",
    ),
    input_file: str | None = typer.Option(
        None, "--input", "-i",
        help="包含 JSON 命令数组的文件路径",
    ),
    json_input: bool = typer.Option(
        False, "--json", "-j",
        help="从 stdin 读取 JSON 命令数组",
    ),
    stop_on_error: bool = typer.Option(
        False, "--stop-on-error", "-s",
        help="遇到第一个错误时停止（默认 continue-on-error）",
    ),
) -> None:
    """批量执行多条命令

    在一个 COM 会话中顺序执行多条命令，默认 continue-on-error。
    支持三种输入方式：

    1. ``--commands`` 直接传 JSON 字符串::

        wps batch --commands '[{"command":"calc.cell-set","params":{"file":"a.xlsx","ref":"A1","value":100}}]'

    2. ``--input`` 从文件读取::

        wps batch --input commands.json

    3. ``--json`` 从 stdin 读取（管道友好）::

        echo '[{...}]' | wps batch --json

    命令支持列表：writer.* / calc.* / impress.* / pdf.* / export.*
    """
    try:
        # 解析输入
        raw: str | None = None
        source = ""
        if commands:
            raw = commands
            source = "--commands"
        elif input_file:
            raw = Path(input_file).read_text(encoding="utf-8")
            source = f"--input {input_file}"
        elif json_input:
            raw = sys.stdin.read()
            source = "--json (stdin)"
        else:
            typer.echo(
                "错误: 请提供 --commands、--input 或 --json 参数。\n"
                "示例: wps batch --commands '[{\"command\":\"calc.cell-set\",\"params\":{...}}]'",
                err=True,
            )
            raise typer.Exit(1)

        if not raw or not raw.strip():
            typer.echo("错误: 输入为空", err=True)
            raise typer.Exit(1)

        try:
            parsed = json_mod.loads(raw.strip())
        except json_mod.JSONDecodeError as e:
            typer.echo(f"错误: JSON 解析失败 ({source}): {e}", err=True)
            raise typer.Exit(1) from e

        if not isinstance(parsed, list):
            typer.echo("错误: 输入必须是 JSON 数组", err=True)
            raise typer.Exit(1)

        if not parsed:
            success(
                {
                    "steps": [],
                    "summary": {"total": 0, "succeeded": 0, "failed": 0},
                },
                command="batch",
                json_mode=True,
            )
            typer.echo("命令数组为空，无操作执行。")
            return

        # 检查是否有驻留进程可以转发
        if _is_resident_running():
            logger.info("检测到驻留进程，转发 batch 请求...")
            try:
                result = _forward_to_resident(parsed, stop_on_error)
                success(result, command="batch", json_mode=True)
                _print_batch_summary(result)
                return
            except Exception as e:
                logger.warning("驻留转发失败，回退到本地执行: %s", e)

        # 本地执行
        executor = BatchExecutor()
        result = executor.execute_batch(parsed, stop_on_error=stop_on_error)

        success(result, command="batch", json_mode=True)
        _print_batch_summary(result)

    except typer.Exit:
        raise
    except Exception as e:
        handle_error(e, command="batch", json_mode=True)


def _print_batch_summary(result: dict) -> None:
    """打印 batch 执行摘要"""
    summary = result.get("summary", {})
    total = summary.get("total", 0)
    succeeded = summary.get("succeeded", 0)
    failed = summary.get("failed", 0)
    typer.echo(f"\n共 {total} 条命令: {succeeded} 成功, {failed} 失败")

    if failed:
        for step in result.get("steps", []):
            if not step.get("success"):
                err = step.get("error", {})
                typer.echo(
                    f"  [#{step['index']}] {step['command']}: "
                    f"{err.get('type', 'Unknown')}: {err.get('message', '')}",
                    err=True,
                )
