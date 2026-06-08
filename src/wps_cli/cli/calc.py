"""Calc CLI 命令"""

from __future__ import annotations

import typer

from wps_cli.cli.common import handle_error, make_get_service, success
from wps_cli.consts import CALC_INPUT_EXTENSIONS
from wps_cli.services.calc_service import CalcService
from wps_cli.utils.path_utils import ensure_safe_input_path, ensure_safe_output_path

app = typer.Typer(help="Excel 电子表格操作")

_get_service = make_get_service(CalcService)


def _safe_calc_input(file: str):
    return ensure_safe_input_path(file, allowed_extensions=CALC_INPUT_EXTENSIONS)


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
        path = _safe_calc_input(file)
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
        path = _safe_calc_input(file)
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
        path = _safe_calc_input(file)
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
        path = _safe_calc_input(file)
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
        path = _safe_calc_input(file)
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
        path = _safe_calc_input(file)
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
        path = _safe_calc_input(file)
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
        path = _safe_calc_input(file)
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
        path = _safe_calc_input(file)
        out_path = ensure_safe_output_path(output) if output else None
        from wps_cli.services.export_service import ExportService

        export_svc = ExportService(manager=_get_service().manager)
        result = export_svc.convert(path, "csv", out_path)
        success({"path": str(result)}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


# ── Validate 验证命令 ──
# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)


@app.command()
def validate(
    file: str = typer.Argument(..., help="文件路径"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """验证 Excel 工作簿完整性

    设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

    检查项: 公式错误、命名区域引用、外部链接、工作表结构
    """
    cmd = "calc.validate"
    try:
        path = _safe_calc_input(file)
        from wps_cli.services.validate_service import ValidateService

        svc = ValidateService(manager=_get_service().manager)
        result = svc.validate_calc(path)
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
    file: str = typer.Argument(..., help="文件路径"),
    field: str = typer.Option(
        "all", "--field", "-f", help="刷新类型: all/pivot"
    ),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """刷新工作簿数据和公式

    设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

    --field all    刷新所有（公式重算 + 外部数据 + 透视表）
    --field pivot  仅刷新透视表缓存
    """
    cmd = "calc.refresh"
    try:
        path = _safe_calc_input(file)
        if field not in ("all", "pivot"):
            from wps_cli.exceptions import ValidationError

            raise ValidationError(
                f"不支持的刷新类型: {field}",
                suggestion="可选: all, pivot",
            )
        svc = _get_service()
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path)
            result = svc.refresh(session.app, field if field != "all" else None)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


# ── 语义视图与路径定位（Phase 4）──
# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)


@app.command("view")
def view(
    file: str = typer.Argument(..., help="文件路径"),
    view_type: str = typer.Argument("summary", help="视图类型: summary/issues/sheets/annotated/stats"),
    type_filter: str = typer.Option("", "--type", "-t", help="过滤问题子类型（仅对 issues 视图有效）"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """工作簿语义视图（参考 OfficeCLI L1 Read）

    设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

    支持五种视图:
      summary   — 工作簿结构摘要（工作表/图表/命名区域）
      issues    — 文档诊断（公式错误/隐藏行列），可用 --type 过滤子类型
      sheets    — 工作表列表概览
      annotated — 带路径标注的单元格内容
      stats     — 纯数字统计
    """
    cmd = f"calc.view_{view_type}"
    try:
        path = _safe_calc_input(file)
        svc = _get_service()
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path, readonly=True)
            if view_type == "summary":
                result = svc.summarize(session.app)
            elif view_type == "issues":
                result = svc.diagnose(session.app)
                if type_filter:
                    result = [r for r in result if r.get("subtype", "") == type_filter]
            elif view_type == "sheets":
                result = svc.sheet_list(session.app)
            elif view_type == "annotated":
                result = svc.annotate(session.app)
            elif view_type == "stats":
                result = svc.get_stats(session.app)
            else:
                from wps_cli.exceptions import ValidationError

                raise ValidationError(
                    f"不支持的视图类型: {view_type}",
                    suggestion="可选: summary, issues, sheets, annotated, stats",
                )
        finally:
            svc.manager.stop(session.session_id)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def get(
    file: str = typer.Argument(..., help="文件路径"),
    path: str = typer.Argument(..., help="元素路径，如 /sheet[\"Sheet1\"]/cell[\"A1\"]"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """通过路径获取工作簿元素值

    设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

    Examples:
        wps calc get data.xlsx '/sheet["Sheet1"]/cell["A1"]'
        wps calc get data.xlsx '/sheet["Sheet1"]/range["A1:B10"]'
        wps calc get data.xlsx '$Sheet1:A1'
    """
    cmd = "calc.get"
    try:
        file_path = _safe_calc_input(file)
        svc = _get_service()
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, file_path, readonly=True)
            from wps_cli.services.path_resolver import PathResolver

            resolver = PathResolver()
            obj = resolver.resolve(session.app, "calc", path)
            value = obj.Value if hasattr(obj, "Value") else str(obj)
            result = {"path": path, "value": value}
        finally:
            svc.manager.stop(session.session_id)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


# ── 条件格式 (Phase 5) ──
# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)


@app.command("conditional-format-add")
def conditional_format_add(
    file: str = typer.Argument(..., help="文件路径"),
    range_ref: str = typer.Option(..., "--range", "-r", help="应用区域，如 A1:A10"),
    cf_type: str = typer.Option(
        "cellvalue", "--type", "-t",
        help="条件类型: cellvalue/formulabased/databar/colorscale/iconset/toprank/aboveaverage",
    ),
    operator: str = typer.Option(
        "greaterthan", "--op", "-p",
        help="运算符: greaterthan/lessthan/between/equal/notequal/contains/notcontains",
    ),
    formula1: str = typer.Option("", "--formula1", "-f1", help="条件值/公式1"),
    formula2: str = typer.Option("", "--formula2", "-f2", help="条件值2（between 时使用）"),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """添加条件格式"""
    cmd = "calc.conditional_format_add"
    try:
        path = _safe_calc_input(file)
        svc = _get_service()
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path)
            idx = svc.conditional_format_add(
                session.app, range_ref, cf_type, operator, formula1, formula2, sheet or None
            )
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"index": idx, "range": range_ref}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command("conditional-format-list")
def conditional_format_list(
    file: str = typer.Argument(..., help="文件路径"),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """列出条件格式"""
    cmd = "calc.conditional_format_list"
    try:
        path = _safe_calc_input(file)
        svc = _get_service()
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path, readonly=True)
            result = svc.conditional_format_list(session.app, sheet or None)
        finally:
            svc.manager.stop(session.session_id)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command("conditional-format-delete")
def conditional_format_delete(
    file: str = typer.Argument(..., help="文件路径"),
    index: int = typer.Option(..., "--index", "-i", help="条件格式序号（0=全部删除）"),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """删除条件格式"""
    cmd = "calc.conditional_format_delete"
    try:
        path = _safe_calc_input(file)
        svc = _get_service()
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path)
            svc.conditional_format_delete(session.app, index, sheet or None)
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"deleted_index": index}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


# ── 数据验证 (Phase 5) ──
# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)


@app.command("data-validation-add")
def data_validation_add(
    file: str = typer.Argument(..., help="文件路径"),
    range_ref: str = typer.Option(..., "--range", "-r", help="应用区域，如 B1:B10"),
    validation_type: str = typer.Option(
        "list", "--type", "-t",
        help="验证类型: whole/decimal/list/date/time/textlength/custom",
    ),
    formula1: str = typer.Option("", "--formula1", "-f1", help="验证公式/值1"),
    formula2: str = typer.Option("", "--formula2", "-f2", help="验证公式/值2"),
    alert_style: str = typer.Option(
        "stop", "--alert", "-a",
        help="警告样式: stop/warning/information",
    ),
    alert_message: str = typer.Option("", "--message", "-m", help="错误提示消息"),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """添加数据验证（下拉列表等）"""
    cmd = "calc.data_validation_add"
    try:
        path = _safe_calc_input(file)
        svc = _get_service()
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path)
            svc.data_validation_add(
                session.app, range_ref, validation_type,
                formula1, formula2, alert_style, alert_message,
                sheet or None,
            )
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"range": range_ref, "type": validation_type}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command("data-validation-list")
def data_validation_list(
    file: str = typer.Argument(..., help="文件路径"),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """列出数据验证"""
    cmd = "calc.data_validation_list"
    try:
        path = _safe_calc_input(file)
        svc = _get_service()
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path, readonly=True)
            result = svc.data_validation_list(session.app, sheet or None)
        finally:
            svc.manager.stop(session.session_id)
        success(result, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


# ── 迷你图 (Phase 5) ──
# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)


@app.command("sparkline-add")
def sparkline_add(
    file: str = typer.Argument(..., help="文件路径"),
    range_ref: str = typer.Option(..., "--range", "-r", help="放置位置，如 F1:F10"),
    source_data: str = typer.Option(..., "--source", "-d", help="数据源区域，如 A1:E10"),
    spark_type: str = typer.Option(
        "line", "--type", "-t",
        help="迷你图类型: line/column/stacked100",
    ),
    sheet: str = typer.Option("", "--sheet", "-s", help="工作表名"),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """添加迷你图"""
    cmd = "calc.sparkline_add"
    try:
        path = _safe_calc_input(file)
        svc = _get_service()
        session = svc.manager.start("calc")
        try:
            svc._open_workbook(session.app, path)
            idx = svc.sparkline_add(
                session.app, range_ref, spark_type, source_data, sheet or None
            )
            svc.save(session.app)
        finally:
            svc.manager.stop(session.session_id)
        success({"sparkline_index": idx, "range": range_ref}, command=cmd, json_mode=json_output)
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)
