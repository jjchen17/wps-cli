"""Calc CLI 命令"""

from __future__ import annotations

import typer

from wps_cli.cli.common import handle_error, make_get_service, success
from wps_cli.services.calc_service import CalcService
from wps_cli.utils.path_utils import ensure_safe_input_path, ensure_safe_output_path

app = typer.Typer(help="Excel 电子表格操作")

_get_service = make_get_service(CalcService)


@app.command()
def new(
    output: str = typer.Option("", "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """新建空白 Excel 工作簿"""
    cmd = "calc.new"
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
    """输出工作簿元信息"""
    cmd = "calc.info"
    try:
        path = ensure_safe_input_path(file)
        result = _get_service().info(path)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def sheet_list(
    file: str = typer.Argument(..., help="文件路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """列出所有工作表"""
    cmd = "calc.sheet_list"
    try:
        path = ensure_safe_input_path(file)
        svc = _get_service()
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path, readonly=True)
            result = svc.sheet_list(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success(result, command=cmd, json_mode=json_output, headers=["index", "name"])
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def cell_get(
    file: str = typer.Argument(..., help="文件路径"),
    ref: str = typer.Argument(..., help="单元格引用，如 A1 或 B3"),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """读取单元格值"""
    cmd = "calc.cell_get"
    try:
        path = ensure_safe_input_path(file)
        svc = _get_service()
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path, readonly=True)
            value = svc.cell_get(session.app, ref, sheet or None)
        finally:
            svc.manager.stop(session.session_id)
        success({"ref": ref, "value": value}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def cell_set(
    file: str = typer.Argument(..., help="文件路径"),
    ref: str = typer.Argument(..., help="单元格引用"),
    value: str = typer.Argument(..., help="值（不能以 = 开头，公式请用 cell-formula）"),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """写入单元格"""
    cmd = "calc.cell_set"
    try:
        path = ensure_safe_input_path(file)
        svc = _get_service()
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path)
            svc.cell_set(session.app, ref, value, sheet or None)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"ref": ref, "value": value}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def cell_range(
    file: str = typer.Argument(..., help="文件路径"),
    ref: str = typer.Argument(..., help="区域引用，如 A1:D10"),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """读取单元格区域"""
    cmd = "calc.cell_range"
    try:
        path = ensure_safe_input_path(file)
        svc = _get_service()
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path, readonly=True)
            result = svc.range_get(session.app, ref, sheet or None)
        finally:
            svc.manager.stop(session.session_id)
        success(
            {
                "range": ref,
                "values": result,
                "rows": len(result),
                "cols": len(result[0]) if result else 0,
            },
            command=cmd,
            json_mode=json_output,
        )
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def cell_formula(
    file: str = typer.Argument(..., help="文件路径"),
    ref: str = typer.Argument(..., help="单元格引用"),
    formula: str = typer.Argument(..., help="公式，如 =SUM(A1:A10)"),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """设置公式（拒绝 SHELL/DDE/HYPERLINK 等危险函数）"""
    cmd = "calc.cell_formula"
    try:
        path = ensure_safe_input_path(file)
        svc = _get_service()
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path)
            svc.cell_formula(session.app, ref, formula, sheet or None)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"ref": ref, "formula": formula}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def chart_create(
    file: str = typer.Argument(..., help="文件路径"),
    data: str = typer.Option(..., "--data", "-d", help="数据区域，如 A1:C10"),
    chart_type: str = typer.Option(
        "bar",
        "--type",
        "-t",
        help="图表类型: bar/line/pie/scatter/area",
    ),
    title: str = typer.Option("", "--title", help="图表标题"),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """创建图表"""
    cmd = "calc.chart_create"
    try:
        path = ensure_safe_input_path(file)
        svc = _get_service()
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path)
            idx = svc.chart_create(session.app, data, chart_type, title, sheet or None)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"chart_index": idx}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def sort(
    file: str = typer.Argument(..., help="文件路径"),
    by: str = typer.Option(..., "--by", "-b", help="排序列"),
    order: str = typer.Option("asc", "--order", help="排序方向: asc/desc"),
    range_ref: str = typer.Option("", "--range", "-r", help="数据区域"),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """排序"""
    cmd = "calc.sort"
    try:
        path = ensure_safe_input_path(file)
        svc = _get_service()
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path)
            svc.sort(session.app, range_ref or "A1:Z1000", by, order, sheet or None)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"sorted_by": by, "order": order}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def export_csv(
    file: str = typer.Argument(..., help="文件路径"),
    output: str = typer.Option("", "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """导出为 CSV"""
    cmd = "calc.export_csv"
    try:
        path = ensure_safe_input_path(file)
        out_path = ensure_safe_output_path(output) if output else None
        from wps_cli.services.export_service import ExportService

        export_svc = ExportService(manager=_get_service().manager)
        result = export_svc.convert(path, "csv", out_path)
        success({"path": str(result)}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)
