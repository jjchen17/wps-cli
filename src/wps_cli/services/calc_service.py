"""Calc 电子表格操作业务逻辑"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from wps_cli.consts import (
    DANGEROUS_FORMULA_TOKENS,
    WD_DO_NOT_SAVE_CHANGES,
    XL_AREA,
    XL_ASCENDING,
    XL_COLUMN_CLUSTERED,
    XL_DESCENDING,
    XL_FILTER_EQUAL,
    XL_FILTER_GREATER,
    XL_FILTER_GREATER_EQUAL,
    XL_FILTER_LESS,
    XL_FILTER_LESS_EQUAL,
    XL_FILTER_NO_OP,
    XL_FILTER_NOT_EQUAL,
    XL_LINE,
    XL_PIE,
    XL_XY_SCATTER,
    XL_YES,
)
from wps_cli.exceptions import ValidationError
from wps_cli.services.session_manager import SessionManager


def _check_formula_safe(formula: str) -> None:
    """阻止已知的危险公式函数

    在 Excel/WPS Calc 中, ``=SHELL()`` / ``=DDE()`` / ``=HYPERLINK()`` 等
    公式可以触发命令执行或外联请求。AI Agent 场景下攻击者可以通过 prompt
    诱导生成恶意公式，因此必须阻断。
    """
    if not isinstance(formula, str):
        return
    upper = formula.upper().replace(" ", "")
    if not upper.startswith("="):
        raise ValidationError(
            f"公式必须以 '=' 开头: {formula!r}",
        )
    for token in DANGEROUS_FORMULA_TOKENS:
        if token in upper:
            raise ValidationError(
                f"公式中包含禁止的函数 {token.rstrip('(')}",
            )


def _coerce_cell_value(value: object) -> object:
    """将 CLI 字符串值规范化为单元格安全值

    若用户传入 ``"=...""``，必须走 ``cell_formula`` 而非 ``cell_set``，否则
    会被 WPS 解析为公式（公式注入二次路径）。这里直接拒绝。
    """
    if isinstance(value, str) and value.lstrip().startswith("="):
        raise ValidationError(
            "cell_set 不接受以 '=' 开头的值（疑似公式）",
        )
    return value


@dataclass
class CalcService:
    """Excel 表格操作"""

    manager: SessionManager

    def _ws(self, app: Any, sheet: str | None = None) -> Any:
        """统一获取工作表：指定名称则按名取，否则取活动工作表"""
        return app.ActiveWorkbook.Sheets(sheet) if sheet else app.ActiveSheet

    @staticmethod
    def _open_workbook(app: Any, path: Path | str, *, readonly: bool = False) -> Any:
        """统一的 Workbook 打开入口，强制只读/读写并禁用外部链接更新

        ``UpdateLinks=0`` 防止打开时自动加载外部数据源；只读模式由调用方按需开启。
        """
        return app.Workbooks.Open(
            str(path),
            UpdateLinks=0,
            ReadOnly=readonly,
        )

    # ── 文档生命周期 ──

    def new(self, output: Path | None = None) -> Path:
        with self.manager.session("calc") as app:
            wb = app.Workbooks.Add()
            if output:
                wb.SaveAs(str(output))
            path = wb.FullName
            wb.Close(WD_DO_NOT_SAVE_CHANGES)
        return Path(path)

    def info(self, path: Path) -> dict:
        with self.manager.session("calc") as app:
            wb = self._open_workbook(app, path, readonly=True)
            result = {
                "path": str(Path(wb.FullName)),
                "sheets": wb.Sheets.Count,
                "sheet_names": [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)],
                "author": wb.BuiltinDocumentProperties("Author").Value,
            }
            wb.Close(WD_DO_NOT_SAVE_CHANGES)
        return result

    # ── 工作表管理 ──

    def sheet_list(self, app: Any) -> list[dict]:
        wb = app.ActiveWorkbook
        return [{"index": i, "name": wb.Sheets(i).Name} for i in range(1, wb.Sheets.Count + 1)]

    def sheet_add(self, app: Any, name: str | None = None) -> str:
        wb = app.ActiveWorkbook
        ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        if name:
            ws.Name = name
        return ws.Name

    def sheet_delete(self, app: Any, name: str) -> None:
        wb = app.ActiveWorkbook
        wb.Sheets(name).Delete()

    def sheet_rename(self, app: Any, old: str, new: str) -> None:
        wb = app.ActiveWorkbook
        wb.Sheets(old).Name = new

    # ── 单元格操作 ──

    def cell_get(self, app: Any, ref: str, sheet: str | None = None) -> object:
        ws = self._ws(app, sheet)
        return ws.Range(ref).Value

    def cell_set(self, app: Any, ref: str, value: object, sheet: str | None = None) -> None:
        safe_value = _coerce_cell_value(value)
        ws = self._ws(app, sheet)
        ws.Range(ref).Value = safe_value

    def cell_formula(self, app: Any, ref: str, formula: str, sheet: str | None = None) -> None:
        _check_formula_safe(formula)
        ws = self._ws(app, sheet)
        ws.Range(ref).Formula = formula

    def range_get(self, app: Any, ref: str, sheet: str | None = None) -> list[list]:
        ws = self._ws(app, sheet)
        rng = ws.Range(ref)
        values = rng.Value
        if values is None:
            return []
        if not isinstance(values, (list, tuple)):
            return [[values]]
        return [list(row) for row in values]

    def range_set(self, app: Any, ref: str, data: list[list], sheet: str | None = None) -> None:
        ws = self._ws(app, sheet)
        ws.Range(ref).Value = data

    def cell_clear(self, app: Any, ref: str, sheet: str | None = None) -> None:
        ws = self._ws(app, sheet)
        ws.Range(ref).ClearContents()

    # ── 行列操作 ──

    def row_insert(self, app: Any, at: int, count: int = 1, sheet: str | None = None) -> None:
        ws = self._ws(app, sheet)
        for _ in range(count):
            ws.Rows(at).Insert()

    def row_delete(self, app: Any, at: int, count: int = 1, sheet: str | None = None) -> None:
        ws = self._ws(app, sheet)
        for _ in range(count):
            ws.Rows(at).Delete()

    def col_insert(self, app: Any, at: int, count: int = 1, sheet: str | None = None) -> None:
        ws = self._ws(app, sheet)
        for _ in range(count):
            ws.Columns(at).Insert()

    def col_delete(self, app: Any, at: int, count: int = 1, sheet: str | None = None) -> None:
        ws = self._ws(app, sheet)
        for _ in range(count):
            ws.Columns(at).Delete()

    # ── 排序与筛选 ──

    def sort(
        self, app: Any, range_ref: str, by_col: str, order: str = "asc", sheet: str | None = None
    ) -> None:
        ws = self._ws(app, sheet)
        rng = ws.Range(range_ref)
        sort_key = ws.Range(f"{by_col}1")
        xl_order = XL_ASCENDING if order == "asc" else XL_DESCENDING
        rng.Sort(Key1=sort_key, Order1=xl_order, Header=XL_YES)

    def auto_filter(
        self, app: Any, range_ref: str, col: int, op: str, value: object, sheet: str | None = None
    ) -> None:
        ws = self._ws(app, sheet)
        rng = ws.Range(range_ref)
        xl_op = {
            "=": XL_FILTER_EQUAL,
            "<>": XL_FILTER_NOT_EQUAL,
            ">": XL_FILTER_GREATER,
            "<": XL_FILTER_LESS,
            ">=": XL_FILTER_GREATER_EQUAL,
            "<=": XL_FILTER_LESS_EQUAL,
        }.get(op, XL_FILTER_EQUAL)
        criteria = f"*{value}*" if op == "contains" else value
        rng.AutoFilter(
            Field=col, Criteria1=criteria, Operator=xl_op if op != "contains" else XL_FILTER_NO_OP
        )

    # ── 图表 ──

    def chart_create(
        self,
        app: Any,
        data_range: str,
        chart_type: str = "bar",
        title: str = "",
        sheet: str | None = None,
    ) -> int:
        ws = self._ws(app, sheet)
        data = ws.Range(data_range)
        chart_obj = ws.ChartObjects().Add(300, 50, 400, 250)
        chart = chart_obj.Chart
        chart.SetSourceData(data)

        type_map = {
            "bar": XL_COLUMN_CLUSTERED,
            "line": XL_LINE,
            "pie": XL_PIE,
            "scatter": XL_XY_SCATTER,
            "area": XL_AREA,
        }
        chart.ChartType = type_map.get(chart_type, XL_COLUMN_CLUSTERED)
        if title:
            chart.HasTitle = True
            chart.ChartTitle.Text = title
        return chart_obj.Index

    def chart_list(self, app: Any, sheet: str | None = None) -> list[dict]:
        ws = self._ws(app, sheet)
        charts = []
        for i in range(1, ws.ChartObjects().Count + 1):
            co = ws.ChartObjects(i)
            charts.append(
                {
                    "index": i,
                    "left": co.Left,
                    "top": co.Top,
                    "width": co.Width,
                    "height": co.Height,
                }
            )
        return charts

    # ── 保存 ──

    def save(self, app: Any, path: Path | None = None) -> Path:
        wb = app.ActiveWorkbook
        if path:
            wb.SaveAs(str(path))
        else:
            wb.Save()
        return Path(wb.FullName)
