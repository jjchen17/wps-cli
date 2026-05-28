"""Calc CLI 命令"""

import typer

from wps_cli.cli.common import do_output, handle_error
from wps_cli.services.calc_service import CalcService
from wps_cli.services.session_manager import SessionManager

app = typer.Typer(help="Excel 电子表格操作")

_manager: SessionManager | None = None
_service: CalcService | None = None


def _get_service() -> CalcService:
    global _manager, _service
    if _service is None:
        from wps_cli.backends.wps_com import WpsComBackend
        _manager = SessionManager(backend=WpsComBackend())
        _service = CalcService(manager=_manager)
    return _service


@app.command()
def new(
    output: str = typer.Option("", "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """新建空白 Excel 工作簿"""
    try:
        from pathlib import Path
        result = _get_service().new(Path(output) if output else None)
        do_output({"success": True, "path": str(result)}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def info(
    file: str = typer.Argument(..., help="文件路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """输出工作簿元信息"""
    try:
        from pathlib import Path
        result = _get_service().info(Path(file))
        do_output(result, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def sheet_list(
    file: str = typer.Argument(..., help="文件路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """列出所有工作表"""
    try:
        from pathlib import Path
        svc = _get_service()
        session = svc.manager.start("calc")
        session.app.Workbooks.Open(str(Path(file)))
        result = svc.sheet_list(session.app)
        svc.manager.stop(session.session_id)
        do_output(result, json_output, headers=["index", "name"])
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def cell_get(
    file: str = typer.Argument(..., help="文件路径"),
    ref: str = typer.Argument(..., help="单元格引用，如 A1 或 B3"),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """读取单元格值"""
    try:
        from pathlib import Path
        svc = _get_service()
        session = svc.manager.start("calc")
        session.app.Workbooks.Open(str(Path(file)))
        value = svc.cell_get(session.app, ref, sheet or None)
        svc.manager.stop(session.session_id)
        do_output({"ref": ref, "value": value}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def cell_set(
    file: str = typer.Argument(..., help="文件路径"),
    ref: str = typer.Argument(..., help="单元格引用"),
    value: str = typer.Argument(..., help="值"),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """写入单元格"""
    try:
        from pathlib import Path
        svc = _get_service()
        session = svc.manager.start("calc")
        session.app.Workbooks.Open(str(Path(file)))
        svc.cell_set(session.app, ref, value, sheet or None)
        svc.save(session.app)
        svc.manager.stop(session.session_id)
        do_output({"success": True, "ref": ref, "value": value}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def cell_range(
    file: str = typer.Argument(..., help="文件路径"),
    ref: str = typer.Argument(..., help="区域引用，如 A1:D10"),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """读取单元格区域"""
    try:
        from pathlib import Path
        svc = _get_service()
        session = svc.manager.start("calc")
        session.app.Workbooks.Open(str(Path(file)))
        result = svc.range_get(session.app, ref, sheet or None)
        svc.manager.stop(session.session_id)
        do_output({"range": ref, "values": result, "rows": len(result), "cols": len(result[0]) if result else 0}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def cell_formula(
    file: str = typer.Argument(..., help="文件路径"),
    ref: str = typer.Argument(..., help="单元格引用"),
    formula: str = typer.Argument(..., help="公式，如 =SUM(A1:A10)"),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """设置公式"""
    try:
        from pathlib import Path
        svc = _get_service()
        session = svc.manager.start("calc")
        session.app.Workbooks.Open(str(Path(file)))
        svc.cell_formula(session.app, ref, formula, sheet or None)
        svc.save(session.app)
        svc.manager.stop(session.session_id)
        do_output({"success": True, "ref": ref, "formula": formula}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def chart_create(
    file: str = typer.Argument(..., help="文件路径"),
    data: str = typer.Option(..., "--data", "-d", help="数据区域，如 A1:C10"),
    type: str = typer.Option("bar", "--type", "-t", help="图表类型: bar/line/pie/scatter/area"),
    title: str = typer.Option("", "--title", help="图表标题"),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """创建图表"""
    try:
        from pathlib import Path
        svc = _get_service()
        session = svc.manager.start("calc")
        session.app.Workbooks.Open(str(Path(file)))
        idx = svc.chart_create(session.app, data, type, title, sheet or None)
        svc.save(session.app)
        svc.manager.stop(session.session_id)
        do_output({"success": True, "chart_index": idx}, json_output)
    except Exception as e:
        handle_error(e, json_output)


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
    try:
        from pathlib import Path
        svc = _get_service()
        session = svc.manager.start("calc")
        session.app.Workbooks.Open(str(Path(file)))
        svc.sort(session.app, range_ref or "A1:Z1000", by, order, sheet or None)
        svc.save(session.app)
        svc.manager.stop(session.session_id)
        do_output({"success": True}, json_output)
    except Exception as e:
        handle_error(e, json_output)


@app.command()
def export_csv(
    file: str = typer.Argument(..., help="文件路径"),
    output: str = typer.Option("", "--output", "-o", help="输出路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """导出为 CSV"""
    try:
        from pathlib import Path
        from wps_cli.services.export_service import ExportService
        export_svc = ExportService(manager=_get_service().manager)
        result = export_svc.convert(Path(file), "csv", Path(output) if output else None)
        do_output({"success": True, "path": str(result)}, json_output)
    except Exception as e:
        handle_error(e, json_output)
