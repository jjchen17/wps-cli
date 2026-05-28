"""Calc 电子表格操作业务逻辑"""

from dataclasses import dataclass
from pathlib import Path

from wps_cli.services.session_manager import SessionManager


@dataclass
class CalcService:
    """Excel 表格操作"""

    manager: SessionManager

    # ── 文档生命周期 ──

    def new(self, output: Path | None = None) -> Path:
        with self.manager.session("calc") as app:
            wb = app.Workbooks.Add()
            if output:
                wb.SaveAs(str(output))
            path = wb.FullName
            wb.Close(0)
        return Path(path)

    def info(self, path: Path) -> dict:
        with self.manager.session("calc") as app:
            wb = app.Workbooks.Open(str(path))
            result = {
                "path": str(Path(wb.FullName)),
                "sheets": wb.Sheets.Count,
                "sheet_names": [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)],
                "author": wb.BuiltinDocumentProperties("Author").Value,
            }
            wb.Close(0)
        return result

    # ── 工作表管理 ──

    def sheet_list(self, app: object) -> list[dict]:
        wb = app.ActiveWorkbook
        return [
            {"index": i, "name": wb.Sheets(i).Name}
            for i in range(1, wb.Sheets.Count + 1)
        ]

    def sheet_add(self, app: object, name: str | None = None) -> str:
        wb = app.ActiveWorkbook
        ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        if name:
            ws.Name = name
        return ws.Name

    def sheet_delete(self, app: object, name: str) -> None:
        wb = app.ActiveWorkbook
        wb.Sheets(name).Delete()

    def sheet_rename(self, app: object, old: str, new: str) -> None:
        wb = app.ActiveWorkbook
        wb.Sheets(old).Name = new

    # ── 单元格操作 ──

    def cell_get(self, app: object, ref: str, sheet: str | None = None) -> object:
        ws = app.ActiveWorkbook.Sheets(sheet) if sheet else app.ActiveSheet
        return ws.Range(ref).Value

    def cell_set(self, app: object, ref: str, value: object, sheet: str | None = None) -> None:
        ws = app.ActiveWorkbook.Sheets(sheet) if sheet else app.ActiveSheet
        ws.Range(ref).Value = value

    def cell_formula(self, app: object, ref: str, formula: str, sheet: str | None = None) -> None:
        ws = app.ActiveWorkbook.Sheets(sheet) if sheet else app.ActiveSheet
        ws.Range(ref).Formula = formula

    def range_get(
        self, app: object, ref: str, sheet: str | None = None
    ) -> list[list]:
        ws = app.ActiveWorkbook.Sheets(sheet) if sheet else app.ActiveSheet
        rng = ws.Range(ref)
        values = rng.Value
        if values is None:
            return []
        if not isinstance(values, (list, tuple)):
            return [[values]]
        return [list(row) for row in values]

    def range_set(
        self, app: object, ref: str, data: list[list], sheet: str | None = None
    ) -> None:
        ws = app.ActiveWorkbook.Sheets(sheet) if sheet else app.ActiveSheet
        ws.Range(ref).Value = data

    def cell_clear(self, app: object, ref: str, sheet: str | None = None) -> None:
        ws = app.ActiveWorkbook.Sheets(sheet) if sheet else app.ActiveSheet
        ws.Range(ref).ClearContents()

    # ── 行列操作 ──

    def row_insert(self, app: object, at: int, count: int = 1, sheet: str | None = None) -> None:
        ws = app.ActiveWorkbook.Sheets(sheet) if sheet else app.ActiveSheet
        for _ in range(count):
            ws.Rows(at).Insert()

    def row_delete(self, app: object, at: int, count: int = 1, sheet: str | None = None) -> None:
        ws = app.ActiveWorkbook.Sheets(sheet) if sheet else app.ActiveSheet
        for _ in range(count):
            ws.Rows(at).Delete()

    def col_insert(self, app: object, at: int, count: int = 1, sheet: str | None = None) -> None:
        ws = app.ActiveWorkbook.Sheets(sheet) if sheet else app.ActiveSheet
        for _ in range(count):
            ws.Columns(at).Insert()

    def col_delete(self, app: object, at: int, count: int = 1, sheet: str | None = None) -> None:
        ws = app.ActiveWorkbook.Sheets(sheet) if sheet else app.ActiveSheet
        for _ in range(count):
            ws.Columns(at).Delete()

    # ── 排序与筛选 ──

    def sort(
        self, app: object, range_ref: str, by_col: str, order: str = "asc", sheet: str | None = None
    ) -> None:
        ws = app.ActiveWorkbook.Sheets(sheet) if sheet else app.ActiveSheet
        rng = ws.Range(range_ref)
        sort_key = ws.Range(f"{by_col}1")
        xl_order = 1 if order == "asc" else 2  # xlAscending / xlDescending
        rng.Sort(Key1=sort_key, Order1=xl_order, Header=1)  # xlYes

    def auto_filter(
        self, app: object, range_ref: str, col: int, op: str, value: object, sheet: str | None = None
    ) -> None:
        ws = app.ActiveWorkbook.Sheets(sheet) if sheet else app.ActiveSheet
        rng = ws.Range(range_ref)
        xl_op = {"=": 1, "<>": 2, ">": 5, "<": 6, ">=": 7, "<=": 8, "contains": 0}.get(op, 1)
        rng.AutoFilter(Field=col, Criteria1=value, Operator=xl_op)

    # ── 图表 ──

    def chart_create(
        self,
        app: object,
        data_range: str,
        chart_type: str = "bar",
        title: str = "",
        sheet: str | None = None,
    ) -> int:
        ws = app.ActiveWorkbook.Sheets(sheet) if sheet else app.ActiveSheet
        data = ws.Range(data_range)
        chart_obj = ws.ChartObjects().Add(300, 50, 400, 250)
        chart = chart_obj.Chart
        chart.SetSourceData(data)

        type_map = {
            "bar": 51,  # xlColumnClustered
            "line": 4,  # xlLine
            "pie": 5,  # xlPie
            "scatter": -4169,  # xlXYScatter
            "area": 76,  # xlArea
        }
        chart.ChartType = type_map.get(chart_type, 51)
        if title:
            chart.HasTitle = True
            chart.ChartTitle.Text = title
        return chart_obj.Index

    def chart_list(self, app: object, sheet: str | None = None) -> list[dict]:
        ws = app.ActiveWorkbook.Sheets(sheet) if sheet else app.ActiveSheet
        charts = []
        for i in range(1, ws.ChartObjects().Count + 1):
            co = ws.ChartObjects(i)
            charts.append({
                "index": i,
                "left": co.Left,
                "top": co.Top,
                "width": co.Width,
                "height": co.Height,
            })
        return charts

    # ── 保存 ──

    def save(self, app: object, path: Path | None = None) -> Path:
        wb = app.ActiveWorkbook
        if path:
            wb.SaveAs(str(path))
        else:
            wb.Save()
        return Path(wb.FullName)
