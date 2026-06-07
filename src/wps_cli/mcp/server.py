# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""MCP (Model Context Protocol) 服务器

将 wps-cli 的文档操作能力通过 MCP 协议暴露给 AI Agent（Claude Code、Cursor 等）。
协议实现：JSON-RPC 2.0 over stdio，零外部依赖。
"""

from __future__ import annotations

import json
import sys
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from wps_cli.backends.wps_com import WpsComBackend
from wps_cli.exceptions import WpsCliError
from wps_cli.services.calc_service import CalcService
from wps_cli.services.export_service import ExportService
from wps_cli.services.impress_service import ImpressService
from wps_cli.services.pdf_service import PdfService
from wps_cli.services.session_manager import SessionManager
from wps_cli.services.writer_service import WriterService

# ── JSON-RPC 2.0 常量 ─────────────────────────────────────────────

JSONRPC_VERSION = "2.0"
MCP_VERSION = "2024-11-05"

# JSON-RPC 错误码
PARSE_ERROR = -32700
INVALID_REQUEST = -32600
METHOD_NOT_FOUND = -32601
INVALID_PARAMS = -32602
INTERNAL_ERROR = -32603
TOOL_NOT_FOUND = -32000
TOOL_CALL_ERROR = -32001


# ── Tool 定义 ─────────────────────────────────────────────────────


@dataclass
class ToolDef:
    """MCP tool 定义"""
    name: str
    description: str
    inputSchema: dict  # noqa: N815 — MCP 协议要求 camelCase


def _make_tools() -> list[ToolDef]:
    """构建所有 MCP tool 定义"""
    return [
        # ── Writer tools ──
        ToolDef(
            name="writer_info",
            description="获取 Word 文档元信息（页数、字数、字符数、段落数、作者、创建/修改时间）",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "Word 文档路径 (.doc/.docx/.wps/.rtf/.txt/.html)"},
                },
                "required": ["file"],
            },
        ),
        ToolDef(
            name="writer_replace",
            description="在 Word 文档中查找替换文本，返回替换次数",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "Word 文档路径"},
                    "old_text": {"type": "string", "description": "要查找的文本"},
                    "new_text": {"type": "string", "description": "替换为的文本"},
                    "wildcard": {"type": "boolean", "description": "是否启用通配符模式（* ? [abc]），默认 false"},
                    "case_sensitive": {"type": "boolean", "description": "是否区分大小写，默认 false"},
                },
                "required": ["file", "old_text", "new_text"],
            },
        ),
        ToolDef(
            name="writer_count",
            description="统计 Word 文档的字数、字符数、段落数、页数",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "Word 文档路径"},
                },
                "required": ["file"],
            },
        ),
        ToolDef(
            name="writer_table_get",
            description="读取 Word 文档中的指定表格，返回二维数组",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "Word 文档路径"},
                    "index": {"type": "integer", "description": "表格序号（从 1 开始），默认 1"},
                },
                "required": ["file"],
            },
        ),
        ToolDef(
            name="writer_table_insert",
            description="在 Word 文档中插入表格，可附带 JSON 二维数组数据",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "Word 文档路径"},
                    "rows": {"type": "integer", "description": "行数"},
                    "cols": {"type": "integer", "description": "列数"},
                    "data_json": {"type": "string", "description": "表格数据的 JSON 二维数组，如 '[[\"A\",\"B\"],[\"C\",\"D\"]]'"},
                },
                "required": ["file", "rows", "cols"],
            },
        ),
        ToolDef(
            name="writer_export_pdf",
            description="将 Word 文档导出为 PDF",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "Word 文档路径"},
                    "output": {"type": "string", "description": "输出 PDF 路径（可选，默认同目录同名 .pdf）"},
                },
                "required": ["file"],
            },
        ),
        ToolDef(
            name="writer_image_insert",
            description="在 Word 文档当前光标位置插入图片",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "Word 文档路径"},
                    "image": {"type": "string", "description": "图片文件路径"},
                    "width": {"type": "number", "description": "图片宽度（磅），可选"},
                    "height": {"type": "number", "description": "图片高度（磅），可选"},
                },
                "required": ["file", "image"],
            },
        ),
        ToolDef(
            name="writer_page_setup",
            description="设置 Word 文档的页面布局（纸张大小、页边距）",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "Word 文档路径"},
                    "width_mm": {"type": "number", "description": "纸张宽度 mm，默认 210"},
                    "height_mm": {"type": "number", "description": "纸张高度 mm，默认 297"},
                    "margin_top": {"type": "number", "description": "上边距 mm，默认 25.4"},
                    "margin_bottom": {"type": "number", "description": "下边距 mm，默认 25.4"},
                    "margin_left": {"type": "number", "description": "左边距 mm，默认 31.75"},
                    "margin_right": {"type": "number", "description": "右边距 mm，默认 31.75"},
                },
                "required": ["file"],
            },
        ),

        # ── Calc tools ──
        ToolDef(
            name="calc_info",
            description="获取 Excel 工作簿元信息（路径、工作表数量、工作表名称列表、作者）",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "Excel 文件路径 (.xls/.xlsx/.xlsm/.et/.csv)"},
                },
                "required": ["file"],
            },
        ),
        ToolDef(
            name="calc_cell_get",
            description="读取 Excel 单元格的值",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "Excel 文件路径"},
                    "ref": {"type": "string", "description": "单元格引用，如 A1 或 B3"},
                    "sheet": {"type": "string", "description": "工作表名称（可选，默认活动工作表）"},
                },
                "required": ["file", "ref"],
            },
        ),
        ToolDef(
            name="calc_cell_set",
            description="设置 Excel 单元格的值（不能以 = 开头，公式请用 calc_cell_formula）",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "Excel 文件路径"},
                    "ref": {"type": "string", "description": "单元格引用，如 A1"},
                    "value": {"type": "string", "description": "要设置的值"},
                    "sheet": {"type": "string", "description": "工作表名称（可选）"},
                },
                "required": ["file", "ref", "value"],
            },
        ),
        ToolDef(
            name="calc_range_get",
            description="读取 Excel 单元格区域的值，返回二维数组",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "Excel 文件路径"},
                    "ref": {"type": "string", "description": "区域引用，如 A1:D10"},
                    "sheet": {"type": "string", "description": "工作表名称（可选）"},
                },
                "required": ["file", "ref"],
            },
        ),
        ToolDef(
            name="calc_cell_formula",
            description="在 Excel 单元格中设置公式（已内置安全防护，禁止 SHELL/DDE/HYPERLINK 等危险函数）",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "Excel 文件路径"},
                    "ref": {"type": "string", "description": "单元格引用，如 A1"},
                    "formula": {"type": "string", "description": "公式，如 =SUM(A1:A10)"},
                    "sheet": {"type": "string", "description": "工作表名称（可选）"},
                },
                "required": ["file", "ref", "formula"],
            },
        ),
        ToolDef(
            name="calc_chart_create",
            description="在 Excel 中创建图表（支持 bar/line/pie/scatter/area）",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "Excel 文件路径"},
                    "data_range": {"type": "string", "description": "数据区域，如 A1:C10"},
                    "chart_type": {"type": "string", "description": "图表类型: bar/line/pie/scatter/area，默认 bar"},
                    "title": {"type": "string", "description": "图表标题（可选）"},
                    "sheet": {"type": "string", "description": "工作表名称（可选）"},
                },
                "required": ["file", "data_range"],
            },
        ),
        ToolDef(
            name="calc_sheet_list",
            description="列出 Excel 工作簿中的所有工作表",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "Excel 文件路径"},
                },
                "required": ["file"],
            },
        ),

        # ── Impress tools ──
        ToolDef(
            name="impress_info",
            description="获取 PPT 演示文稿元信息（路径、幻灯片数量、标题、作者）",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "PPT 文件路径 (.ppt/.pptx/.dps)"},
                },
                "required": ["file"],
            },
        ),
        ToolDef(
            name="impress_slide_list",
            description="列出 PPT 中所有幻灯片的索引、标题、版式信息",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "PPT 文件路径"},
                },
                "required": ["file"],
            },
        ),
        ToolDef(
            name="impress_text_get",
            description="提取 PPT 某张幻灯片中的所有文本内容",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "PPT 文件路径"},
                    "slide_idx": {"type": "integer", "description": "幻灯片索引（从 1 开始）"},
                },
                "required": ["file", "slide_idx"],
            },
        ),
        ToolDef(
            name="impress_text_set",
            description="设置 PPT 某张幻灯片中指定占位符的文本（title/body/subtitle）",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "PPT 文件路径"},
                    "slide_idx": {"type": "integer", "description": "幻灯片索引（从 1 开始）"},
                    "placeholder": {"type": "string", "description": "占位符类型: title / body / subtitle，默认 title"},
                    "text": {"type": "string", "description": "要设置的文本内容"},
                },
                "required": ["file", "slide_idx", "text"],
            },
        ),
        ToolDef(
            name="impress_export_pdf",
            description="将 PPT 演示文稿导出为 PDF",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "PPT 文件路径"},
                    "output": {"type": "string", "description": "输出 PDF 路径（可选）"},
                },
                "required": ["file"],
            },
        ),

        # ── PDF tools ──
        ToolDef(
            name="pdf_info",
            description="获取 PDF 文件元信息（路径、文件大小、修改时间）",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "PDF 文件路径"},
                },
                "required": ["file"],
            },
        ),
        ToolDef(
            name="pdf_merge",
            description="合并多个 PDF 文件为一个",
            inputSchema={
                "type": "object",
                "properties": {
                    "files_json": {"type": "string", "description": "要合并的 PDF 文件路径 JSON 数组，如 '[\"a.pdf\",\"b.pdf\"]'"},
                    "output": {"type": "string", "description": "输出路径"},
                },
                "required": ["files_json", "output"],
            },
        ),
        ToolDef(
            name="pdf_extract_pages",
            description="从 PDF 中提取指定页面，页码范围格式如 '1-3,5,7-9'",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "PDF 文件路径"},
                    "pages": {"type": "string", "description": "页码范围，如 1-3,5,7-9"},
                    "output": {"type": "string", "description": "输出路径"},
                },
                "required": ["file", "pages", "output"],
            },
        ),
        ToolDef(
            name="pdf_watermark",
            description="给 PDF 添加文字水印",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "PDF 文件路径"},
                    "text": {"type": "string", "description": "水印文字（最多 100 字符）"},
                    "output": {"type": "string", "description": "输出路径"},
                },
                "required": ["file", "text", "output"],
            },
        ),
        ToolDef(
            name="pdf_split",
            description="按每 N 页拆分 PDF",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "PDF 文件路径"},
                    "every": {"type": "integer", "description": "每 N 页拆分为一个文件"},
                    "output_dir": {"type": "string", "description": "输出目录"},
                },
                "required": ["file", "every", "output_dir"],
            },
        ),

        # ── Export tools ──
        ToolDef(
            name="export_convert",
            description="通用格式转换（支持 Word/Excel/PPT 之间的格式转换）",
            inputSchema={
                "type": "object",
                "properties": {
                    "file": {"type": "string", "description": "源文件路径"},
                    "output_format": {"type": "string", "description": "目标格式: docx/pdf/html/txt/csv/xlsx/pptx"},
                    "output": {"type": "string", "description": "输出路径（可选）"},
                },
                "required": ["file", "output_format"],
            },
        ),
    ]


# ── JSON-RPC 消息结构 ─────────────────────────────────────────────


@dataclass
class JsonRpcRequest:
    jsonrpc: str = JSONRPC_VERSION
    id: int | str = 0
    method: str = ""
    params: dict = field(default_factory=dict)


@dataclass
class JsonRpcResponse:
    id: int | str
    result: Any = None
    jsonrpc: str = JSONRPC_VERSION


@dataclass
class JsonRpcError:
    id: int | str
    code: int
    message: str
    data: Any = None
    jsonrpc: str = JSONRPC_VERSION


# ── Tool Handler 类型 ─────────────────────────────────────────────

ToolHandler = Callable[[dict], dict]


# ── MCP 服务器主类 ────────────────────────────────────────────────


class WpsMcpServer:
    """MCP stdio 服务器

    将 wps-cli 的所有文档操作能力通过 JSON-RPC 2.0 over stdio 暴露给 AI Agent。

    用法::

        python -m wps_cli.mcp.server
        # 或
        wps mcp serve
    """

    def __init__(self):
        self._manager = SessionManager(backend=WpsComBackend())
        self._writer = WriterService(manager=self._manager)
        self._calc = CalcService(manager=self._manager)
        self._impress = ImpressService(manager=self._manager)
        self._pdf = PdfService(manager=self._manager)
        self._export = ExportService(manager=self._manager)
        self._tools: list[ToolDef] = _make_tools()
        self._handlers: dict[str, ToolHandler] = {
            "writer_info": self._handle_writer_info,
            "writer_replace": self._handle_writer_replace,
            "writer_count": self._handle_writer_count,
            "writer_table_get": self._handle_writer_table_get,
            "writer_table_insert": self._handle_writer_table_insert,
            "writer_export_pdf": self._handle_writer_export_pdf,
            "writer_image_insert": self._handle_writer_image_insert,
            "writer_page_setup": self._handle_writer_page_setup,
            "calc_info": self._handle_calc_info,
            "calc_cell_get": self._handle_calc_cell_get,
            "calc_cell_set": self._handle_calc_cell_set,
            "calc_range_get": self._handle_calc_range_get,
            "calc_cell_formula": self._handle_calc_cell_formula,
            "calc_chart_create": self._handle_calc_chart_create,
            "calc_sheet_list": self._handle_calc_sheet_list,
            "impress_info": self._handle_impress_info,
            "impress_slide_list": self._handle_impress_slide_list,
            "impress_text_get": self._handle_impress_text_get,
            "impress_text_set": self._handle_impress_text_set,
            "impress_export_pdf": self._handle_impress_export_pdf,
            "pdf_info": self._handle_pdf_info,
            "pdf_merge": self._handle_pdf_merge,
            "pdf_extract_pages": self._handle_pdf_extract_pages,
            "pdf_watermark": self._handle_pdf_watermark,
            "pdf_split": self._handle_pdf_split,
            "export_convert": self._handle_export_convert,
        }

    def list_tools(self) -> list[dict]:
        """返回所有可用工具定义"""
        return [
            {"name": t.name, "description": t.description, "inputSchema": t.inputSchema}
            for t in self._tools
        ]

    def call_tool(self, name: str, arguments: dict) -> dict:
        """调用指定工具并返回结果"""
        handler = self._handlers.get(name)
        if handler is None:
            raise KeyError(f"未知工具: {name}")

        try:
            result = handler(arguments)
            return {"content": [{"type": "text", "text": json.dumps(result, ensure_ascii=False, default=str)}]}
        except WpsCliError as e:
            return {
                "content": [{"type": "text", "text": json.dumps({
                    "error": type(e).__name__,
                    "message": str(e),
                    "suggestion": e.suggestion,
                }, ensure_ascii=False)}],
                "isError": True,
            }
        except Exception as e:
            return {
                "content": [{"type": "text", "text": json.dumps({
                    "error": type(e).__name__,
                    "message": str(e),
                }, ensure_ascii=False)}],
                "isError": True,
            }

    # ── Writer handlers ──

    def _handle_writer_info(self, args: dict) -> dict:
        path = Path(args["file"])
        return self._writer.info(path)

    def _handle_writer_replace(self, args: dict) -> dict:
        path = Path(args["file"])
        old = args["old_text"]
        new = args["new_text"]
        wildcard = args.get("wildcard", False)
        case = args.get("case_sensitive", False)
        svc = self._writer
        session = svc.open_document(path)
        try:
            count = svc.text_replace(session.app, old, new, wildcard, case)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        return {"replaced": count}

    def _handle_writer_count(self, args: dict) -> dict:
        path = Path(args["file"])
        svc = self._writer
        session = svc.open_document(path, readonly=True)
        try:
            return svc.text_count(session.app)
        finally:
            svc.manager.stop(session.session_id)

    def _handle_writer_table_get(self, args: dict) -> dict:
        path = Path(args["file"])
        index = args.get("index", 1)
        svc = self._writer
        session = svc.open_document(path, readonly=True)
        try:
            data = svc.table_get(session.app, index)
            return {"table_index": index, "data": data}
        finally:
            svc.manager.stop(session.session_id)

    def _handle_writer_table_insert(self, args: dict) -> dict:
        path = Path(args["file"])
        rows = args["rows"]
        cols = args["cols"]
        data_raw = args.get("data_json")
        data = json.loads(data_raw) if data_raw else None
        svc = self._writer
        session = svc.open_document(path)
        try:
            idx = svc.table_insert(session.app, rows, cols, data)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        return {"table_index": idx}

    def _handle_writer_export_pdf(self, args: dict) -> dict:
        path = Path(args["file"])
        output = Path(args["output"]) if args.get("output") else path.with_suffix(".pdf")
        svc = self._writer
        session = svc.open_document(path, readonly=True)
        try:
            svc.export_pdf(session.app, output)
        finally:
            svc.manager.stop(session.session_id)
        return {"path": str(output)}

    def _handle_writer_image_insert(self, args: dict) -> dict:
        path = Path(args["file"])
        image = Path(args["image"])
        width = args.get("width")
        height = args.get("height")
        svc = self._writer
        session = svc.open_document(path)
        try:
            svc.image_insert(session.app, image, width, height)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        return {"image": str(image)}

    def _handle_writer_page_setup(self, args: dict) -> dict:
        path = Path(args["file"])
        svc = self._writer
        session = svc.open_document(path)
        try:
            svc.page_setup(
                session.app,
                width_mm=args.get("width_mm"),
                height_mm=args.get("height_mm"),
                margin_top=args.get("margin_top"),
                margin_bottom=args.get("margin_bottom"),
                margin_left=args.get("margin_left"),
                margin_right=args.get("margin_right"),
            )
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        return {"status": "ok"}

    # ── Calc handlers ──

    def _handle_calc_info(self, args: dict) -> dict:
        path = Path(args["file"])
        return self._calc.info(path)

    def _handle_calc_cell_get(self, args: dict) -> dict:
        path = Path(args["file"])
        ref = args["ref"]
        sheet = args.get("sheet")
        svc = self._calc
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path, readonly=True)
            value = svc.cell_get(session.app, ref, sheet or None)
        finally:
            svc.manager.stop(session.session_id)
        return {"ref": ref, "value": value}

    def _handle_calc_cell_set(self, args: dict) -> dict:
        path = Path(args["file"])
        ref = args["ref"]
        value = args["value"]
        sheet = args.get("sheet")
        svc = self._calc
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path)
            svc.cell_set(session.app, ref, value, sheet or None)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        return {"ref": ref, "value": value}

    def _handle_calc_range_get(self, args: dict) -> dict:
        path = Path(args["file"])
        ref = args["ref"]
        sheet = args.get("sheet")
        svc = self._calc
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path, readonly=True)
            values = svc.range_get(session.app, ref, sheet or None)
        finally:
            svc.manager.stop(session.session_id)
        return {
            "range": ref,
            "values": values,
            "rows": len(values),
            "cols": len(values[0]) if values else 0,
        }

    def _handle_calc_cell_formula(self, args: dict) -> dict:
        path = Path(args["file"])
        ref = args["ref"]
        formula = args["formula"]
        sheet = args.get("sheet")
        svc = self._calc
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path)
            svc.cell_formula(session.app, ref, formula, sheet or None)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        return {"ref": ref, "formula": formula}

    def _handle_calc_chart_create(self, args: dict) -> dict:
        path = Path(args["file"])
        data_range = args["data_range"]
        chart_type = args.get("chart_type", "bar")
        title = args.get("title", "")
        sheet = args.get("sheet")
        svc = self._calc
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path)
            idx = svc.chart_create(session.app, data_range, chart_type, title, sheet or None)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        return {"chart_index": idx, "chart_type": chart_type, "title": title}

    def _handle_calc_sheet_list(self, args: dict) -> dict:
        path = Path(args["file"])
        svc = self._calc
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path, readonly=True)
            sheets = svc.sheet_list(session.app)
        finally:
            svc.manager.stop(session.session_id)
        return {"sheets": sheets}

    # ── Impress handlers ──

    def _handle_impress_info(self, args: dict) -> dict:
        path = Path(args["file"])
        return self._impress.info(path)

    def _handle_impress_slide_list(self, args: dict) -> dict:
        path = Path(args["file"])
        svc = self._impress
        session = self._open_impress(svc, path, readonly=True)
        try:
            slides = svc.slide_list(session.app)
        finally:
            svc.manager.stop(session.session_id)
        return {"slides": slides}

    def _handle_impress_text_get(self, args: dict) -> dict:
        path = Path(args["file"])
        slide_idx = args["slide_idx"]
        svc = self._impress
        session = self._open_impress(svc, path, readonly=True)
        try:
            text = svc.text_get(session.app, slide_idx)
        finally:
            svc.manager.stop(session.session_id)
        return {"slide": slide_idx, "text": text}

    def _handle_impress_text_set(self, args: dict) -> dict:
        path = Path(args["file"])
        slide_idx = args["slide_idx"]
        placeholder = args.get("placeholder", "title")
        text = args["text"]
        svc = self._impress
        session = self._open_impress(svc, path)
        try:
            svc.text_set(session.app, slide_idx, placeholder, text)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        return {"slide": slide_idx, "placeholder": placeholder}

    def _handle_impress_export_pdf(self, args: dict) -> dict:
        path = Path(args["file"])
        output = Path(args["output"]) if args.get("output") else path.with_suffix(".pdf")
        svc = self._impress
        session = self._open_impress(svc, path, readonly=True)
        try:
            svc.export_pdf(session.app, output)
        finally:
            svc.manager.stop(session.session_id)
        return {"path": str(output)}

    # ── PDF handlers ──

    def _handle_pdf_info(self, args: dict) -> dict:
        path = Path(args["file"])
        return self._pdf.info(path)

    def _handle_pdf_merge(self, args: dict) -> dict:
        files_json = args["files_json"]
        inputs = [Path(p) for p in json.loads(files_json)]
        output = Path(args["output"])
        result = self._pdf.merge(inputs, output)
        return {"path": str(result), "merged": len(inputs)}

    def _handle_pdf_extract_pages(self, args: dict) -> dict:
        path = Path(args["file"])
        pages = args["pages"]
        output = Path(args["output"])
        result = self._pdf.extract_pages(path, pages, output)
        return {"path": str(result), "pages": pages}

    def _handle_pdf_watermark(self, args: dict) -> dict:
        path = Path(args["file"])
        text = args["text"]
        output = Path(args["output"])
        result = self._pdf.watermark(path, text, output)
        return {"path": str(result)}

    def _handle_pdf_split(self, args: dict) -> dict:
        path = Path(args["file"])
        every = args["every"]
        output_dir = Path(args["output_dir"])
        results = self._pdf.split(path, every, output_dir)
        return {"parts": len(results), "files": [str(r) for r in results]}

    # ── Export handlers ──

    def _handle_export_convert(self, args: dict) -> dict:
        path = Path(args["file"])
        output_format = args["output_format"]
        output = Path(args["output"]) if args.get("output") else None
        result = self._export.convert(path, output_format, output)
        return {"path": str(result), "format": output_format}

    # ── Impress 助手 ──

    @staticmethod
    def _open_impress(svc: ImpressService, path: Path, readonly: bool = False):
        """打开 Impress 演示文稿并返回会话"""
        session = svc.manager.start("impress")
        try:
            session.app.Presentations.Open(str(path), ReadOnly=readonly)
            return session
        except Exception:
            svc.manager.stop(session.session_id)
            raise

    # ── JSON-RPC 请求处理 ──

    def _handle_request(self, raw: dict) -> dict:
        """处理单个 JSON-RPC 请求"""
        req_id = raw.get("id", 0)
        method = raw.get("method", "")

        try:
            if method == "initialize":
                return self._make_response(req_id, {
                    "protocolVersion": MCP_VERSION,
                    "capabilities": {"tools": {}},
                    "serverInfo": {
                        "name": "wps-cli",
                        "version": "0.1.0",
                    },
                })

            elif method == "tools/list":
                return self._make_response(req_id, {"tools": self.list_tools()})

            elif method == "tools/call":
                params = raw.get("params", {})
                tool_name = params.get("name", "")
                arguments = params.get("arguments", {})
                try:
                    result = self.call_tool(tool_name, arguments)
                    return self._make_response(req_id, result)
                except KeyError:
                    return self._make_error(req_id, TOOL_NOT_FOUND, f"未知工具: {tool_name}")

            elif method == "notifications/initialized":
                # 客户端初始化完成通知，无需响应
                return {}

            else:
                return self._make_error(req_id, METHOD_NOT_FOUND, f"未知方法: {method}")

        except Exception as e:
            return self._make_error(req_id, INTERNAL_ERROR, f"内部错误: {e}")

    @staticmethod
    def _make_response(req_id: int | str, result: Any) -> dict:
        return {"jsonrpc": JSONRPC_VERSION, "id": req_id, "result": result}

    @staticmethod
    def _make_error(req_id: int | str, code: int, message: str, data: Any = None) -> dict:
        error = {"code": code, "message": message}
        if data is not None:
            error["data"] = data
        return {"jsonrpc": JSONRPC_VERSION, "id": req_id, "error": error}

    # ── stdio 主循环 ──

    def run(self) -> None:
        """stdio 主循环：读取 JSON-RPC 请求 → 处理 → 输出响应

        每一行是一个完整的 JSON-RPC 请求/通知。
        """
        try:
            for line in sys.stdin:
                line = line.strip()
                if not line:
                    continue
                try:
                    request = json.loads(line)
                except json.JSONDecodeError as e:
                    err_response = self._make_error(0, PARSE_ERROR, f"JSON 解析错误: {e}")
                    self._write_response(err_response)
                    continue

                response = self._handle_request(request)

                # 通知（如 notifications/initialized）不需要响应
                if response:
                    self._write_response(response)
        except KeyboardInterrupt:
            pass
        finally:
            self._cleanup()

    @staticmethod
    def _write_response(response: dict) -> None:
        """将 JSON-RPC 响应写入 stdout"""
        sys.stdout.write(json.dumps(response, ensure_ascii=False, default=str) + "\n")
        sys.stdout.flush()

    def _cleanup(self) -> None:
        """清理所有 COM 会话"""
        try:
            self._manager.stop_all()
        except Exception:
            pass

    def __del__(self) -> None:
        self._cleanup()


def main() -> None:
    """entry point for console_scripts"""
    server = WpsMcpServer()
    server.run()


if __name__ == "__main__":
    main()
