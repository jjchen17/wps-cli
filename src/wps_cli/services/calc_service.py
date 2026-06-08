"""Calc 电子表格操作业务逻辑

设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from wps_cli.consts import (
    DANGEROUS_FORMULA_TOKENS,
    WD_DO_NOT_SAVE_CHANGES,
    XL_AREA,
    XL_ASCENDING,
    XL_CF_ABOVE_AVERAGE,
    XL_CF_CELL_VALUE,
    XL_CF_COLOR_SCALE,
    XL_CF_DATA_BAR,
    XL_CF_EXPRESSION,
    XL_CF_ICON_SET,
    XL_CF_OP_BETWEEN,
    XL_CF_OP_EQUAL,
    XL_CF_OP_GREATER,
    XL_CF_OP_LESS,
    XL_CF_OP_NOT_EQUAL,
    XL_CF_TEXT_STRING,
    XL_CF_TOP_10,
    XL_COLUMN_CLUSTERED,
    XL_CONTAINS,
    XL_DESCENDING,
    XL_DOES_NOT_CONTAIN,
    XL_DV_ALERT_INFO,
    XL_DV_ALERT_STOP,
    XL_DV_ALERT_WARNING,
    XL_DV_CUSTOM,
    XL_DV_DATE,
    XL_DV_DECIMAL,
    XL_DV_LIST,
    XL_DV_TEXT_LENGTH,
    XL_DV_TIME,
    XL_DV_WHOLE,
    XL_FILTER_EQUAL,
    XL_FILTER_GREATER,
    XL_FILTER_GREATER_EQUAL,
    XL_FILTER_LESS,
    XL_FILTER_LESS_EQUAL,
    XL_FILTER_NO_OP,
    XL_FILTER_NOT_EQUAL,
    XL_LINE,
    XL_PIE,
    XL_SPARK_COLUMN,
    XL_SPARK_COLUMN_STACKED100,
    XL_SPARK_LINE,
    XL_XY_SCATTER,
    XL_YES,
)
from wps_cli.exceptions import ValidationError
from wps_cli.services.session_manager import SessionManager


def _col_letter(col: int) -> str:
    """将列号转换为字母 (1→A, 2→B, ..., 27→AA)"""
    result = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def _check_formula_safe(formula: str) -> None:
    """阻止已知的危险公式函数

    在 Excel/WPS Calc 中, ``=SHELL()`` / ``=DDE()`` / ``=HYPERLINK()`` 等
    公式可以触发命令执行或外联请求。AI Agent 场景下攻击者可以通过 prompt
    诱导生成恶意公式，因此必须阻断。

    注意：移除所有 Unicode 空白字符（含换行符、制表符、全角空格等），
    防止 ``=SHELL\\n("cmd")`` 这种换行绕过攻击。
    """
    import re

    if not isinstance(formula, str):
        return
    upper = re.sub(r"\s+", "", formula.upper())
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

    # ── 条件格式 (Conditional Formatting) ──
    # 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

    def conditional_format_add(
        self,
        app: Any,
        range_ref: str,
        cf_type: str = "cellvalue",
        operator: str = "greaterthan",
        formula1: str = "",
        formula2: str = "",
        sheet: str | None = None,
    ) -> int:
        """添加条件格式

        Args:
            cf_type: cellvalue | formulabased | databar | colorscale |
                     iconset | toprank | aboveaverage
            operator: greaterthan | lessthan | between | equal |
                      notequal | contains | notcontains
        """
        ws = self._ws(app, sheet)
        rng = ws.Range(range_ref)

        cf_type_map: dict[str, int] = {
            "cellvalue": XL_CF_CELL_VALUE,
            "formulabased": XL_CF_EXPRESSION,
            "databar": XL_CF_DATA_BAR,
            "colorscale": XL_CF_COLOR_SCALE,
            "iconset": XL_CF_ICON_SET,
            "toprank": XL_CF_TOP_10,
            "aboveaverage": XL_CF_ABOVE_AVERAGE,
        }
        xl_op_map: dict[str, int] = {
            "greaterthan": XL_CF_OP_GREATER,
            "lessthan": XL_CF_OP_LESS,
            "between": XL_CF_OP_BETWEEN,
            "equal": XL_CF_OP_EQUAL,
            "notequal": XL_CF_OP_NOT_EQUAL,
        }

        xl_type = cf_type_map.get(cf_type, XL_CF_CELL_VALUE)

        # 文本类条件（contains / notcontains）使用 xlTextString 类型
        if operator in ("contains", "notcontains"):
            text_op = XL_CONTAINS if operator == "contains" else XL_DOES_NOT_CONTAIN
            try:
                cf = rng.FormatConditions.Add(
                    Type=XL_CF_TEXT_STRING,
                    TextOperator=text_op,
                    String=formula1,
                )
            except Exception:
                # 回退：使用公式方式
                first_cell = range_ref.split(":")[0] if ":" in range_ref else range_ref
                if operator == "contains":
                    formula = f'=ISNUMBER(SEARCH("{formula1}",{first_cell}))'
                else:
                    formula = f'=NOT(ISNUMBER(SEARCH("{formula1}",{first_cell})))'
                cf = rng.FormatConditions.Add(
                    Type=XL_CF_EXPRESSION,
                    Formula1=formula,
                )
        elif cf_type == "formulabased":
            cf = rng.FormatConditions.Add(
                Type=XL_CF_EXPRESSION,
                Formula1=formula1,
            )
        elif cf_type in ("databar", "colorscale", "iconset"):
            cf = rng.FormatConditions.Add(Type=xl_type)
        else:
            xl_op = xl_op_map.get(operator, XL_CF_OP_GREATER)
            cf = rng.FormatConditions.Add(
                Type=xl_type,
                Operator=xl_op,
                Formula1=formula1,
                Formula2=formula2,
            )

        try:
            cf.SetFirstPriority()
        except Exception:
            pass
        try:
            cf.StopIfTrue = True
        except Exception:
            pass
        try:
            cf.Interior.Color = 0x00FFFF  # 黄色底色
        except Exception:
            pass
        try:
            cf.Font.Color = 0x0000FF  # 红色字体
        except Exception:
            pass

        return int(cf.Index)

    def conditional_format_list(
        self, app: Any, sheet: str | None = None
    ) -> list[dict]:
        """列出条件格式"""
        ws = self._ws(app, sheet)
        results: list[dict] = []
        try:
            count = ws.Cells.FormatConditions.Count
        except Exception:
            return results

        for i in range(1, count + 1):
            try:
                fc = ws.Cells.FormatConditions(i)
                results.append({
                    "index": i,
                    "type": str(fc.Type),
                    "operator": str(getattr(fc, "Operator", "")),
                    "formula1": str(getattr(fc, "Formula1", "")),
                    "formula2": str(getattr(fc, "Formula2", "")),
                    "applies_to": str(getattr(fc, "AppliesTo", {}).Address if hasattr(fc, "AppliesTo") else ""),
                })
            except Exception:
                pass
        return results

    def conditional_format_delete(
        self, app: Any, index: int, sheet: str | None = None
    ) -> None:
        """删除条件格式

        Args:
            index: 条件格式序号（从1开始），传0则删除全部
        """
        ws = self._ws(app, sheet)
        try:
            if index == 0:
                ws.Cells.FormatConditions.Delete()
            else:
                ws.Cells.FormatConditions(index).Delete()
        except Exception as e:
            raise RuntimeError(f"删除条件格式失败 (index={index}): {e}") from e

    def conditional_format_set_priority(
        self, app: Any, index: int, priority: int, sheet: str | None = None
    ) -> None:
        """设置条件格式优先级"""
        ws = self._ws(app, sheet)
        try:
            fc = ws.Cells.FormatConditions(index)
            fc.SetFirstPriority()
            for _ in range(priority - 1):
                fc.SetLastPriority()
        except Exception as e:
            raise RuntimeError(
                f"设置条件格式优先级失败 (index={index}, priority={priority}): {e}"
            ) from e

    # ── 数据验证 (Data Validation) ──
    # 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

    def data_validation_add(
        self,
        app: Any,
        range_ref: str,
        validation_type: str = "list",
        formula1: str = "",
        formula2: str = "",
        alert_style: str = "stop",
        alert_message: str = "",
        sheet: str | None = None,
    ) -> None:
        """添加数据验证

        Args:
            validation_type: whole | decimal | list | date | time |
                             textlength | custom
            alert_style: stop | warning | information
        """
        ws = self._ws(app, sheet)
        rng = ws.Range(range_ref)

        dv_type_map: dict[str, int] = {
            "whole": XL_DV_WHOLE,
            "decimal": XL_DV_DECIMAL,
            "list": XL_DV_LIST,
            "date": XL_DV_DATE,
            "time": XL_DV_TIME,
            "textlength": XL_DV_TEXT_LENGTH,
            "custom": XL_DV_CUSTOM,
        }
        alert_map: dict[str, int] = {
            "stop": XL_DV_ALERT_STOP,
            "warning": XL_DV_ALERT_WARNING,
            "information": XL_DV_ALERT_INFO,
        }

        xl_type = dv_type_map.get(validation_type, XL_DV_LIST)
        xl_alert = alert_map.get(alert_style, XL_DV_ALERT_STOP)

        try:
            validation = rng.Validation
            validation.Add(
                Type=xl_type,
                AlertStyle=xl_alert,
                Operator=XL_CF_OP_BETWEEN,
                Formula1=formula1,
                Formula2=formula2,
            )
            if validation_type == "list":
                try:
                    validation.InCellDropdown = True
                except Exception:
                    pass
            if alert_message:
                try:
                    validation.ErrorTitle = "输入错误"
                    validation.ErrorMessage = alert_message
                except Exception:
                    pass
        except Exception as e:
            raise RuntimeError(
                f"添加数据验证失败 (range={range_ref}, type={validation_type}): {e}"
            ) from e

    def data_validation_list(
        self, app: Any, sheet: str | None = None
    ) -> list[dict]:
        """列出数据验证"""
        ws = self._ws(app, sheet)
        results: list[dict] = []
        try:
            count = ws.UsedRange.Validation.Count if ws.UsedRange else 0
        except Exception:
            return results

        for i in range(1, count + 1):
            try:
                dv = ws.UsedRange.Validation(i)  # type: ignore[call-arg]
                results.append({
                    "index": i,
                    "type": str(dv.Type),
                    "formula1": str(getattr(dv, "Formula1", "")),
                    "formula2": str(getattr(dv, "Formula2", "")),
                })
            except Exception:
                pass
        return results

    def data_validation_delete(
        self, app: Any, range_ref: str = "", sheet: str | None = None
    ) -> None:
        """删除数据验证

        Args:
            range_ref: 要删除验证的区域。为空则删除整个工作表的验证。
        """
        ws = self._ws(app, sheet)
        try:
            if range_ref:
                ws.Range(range_ref).Validation.Delete()
            else:
                ws.Cells.Validation.Delete()
        except Exception as e:
            raise RuntimeError(
                f"删除数据验证失败 (range={range_ref}): {e}"
            ) from e

    # ── 迷你图 (Sparklines) ──
    # 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

    def sparkline_add(
        self,
        app: Any,
        range_ref: str,
        spark_type: str = "line",
        source_data: str = "",
        sheet: str | None = None,
    ) -> int:
        """添加迷你图组

        Args:
            range_ref: 迷你图放置位置（如 F1:F10）
            spark_type: line | column | stacked100
            source_data: 数据源区域（如 A1:E10）
        """
        ws = self._ws(app, sheet)

        spark_type_map: dict[str, int] = {
            "line": XL_SPARK_LINE,
            "column": XL_SPARK_COLUMN,
            "stacked100": XL_SPARK_COLUMN_STACKED100,
        }
        xl_type = spark_type_map.get(spark_type, XL_SPARK_LINE)

        try:
            target = ws.Range(range_ref)
            sg = target.SparklineGroups.Add(Type=xl_type, SourceData=source_data)
            return int(sg.Index if hasattr(sg, "Index") else 1)
        except AttributeError as exc:
            raise RuntimeError(
                "该版本的 WPS 不支持迷你图 (Sparklines) 功能，"
                "请使用 WPS 2019 或更高版本"
            ) from exc
        except Exception as e:
            raise RuntimeError(
                f"添加迷你图失败 (range={range_ref}): {e}"
            ) from e

    # ── 保存 ──

    def save(self, app: Any, path: Path | None = None) -> Path:
        wb = app.ActiveWorkbook
        if path:
            wb.SaveAs(str(path))
        else:
            wb.Save()
        return Path(wb.FullName)

    # ── Refresh 刷新 ──
    # 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

    def refresh(self, app: Any, field_type: str | None = None) -> dict:
        """刷新工作簿数据

        Args:
            field_type: 刷新类型过滤 (pivot/all), None 或 "all" 表示全部
        """
        wb = app.ActiveWorkbook
        result: dict = {"field_type": field_type or "all", "actions": []}

        if field_type is None or field_type == "all":
            try:
                wb.RefreshAll()
                result["actions"].append(
                    {"action": "refresh_all", "status": "ok"}
                )
            except Exception as e:
                result["actions"].append(
                    {"action": "refresh_all", "status": "error", "message": str(e)}
                )

            try:
                for i in range(1, wb.PivotCaches.Count + 1):
                    try:
                        wb.PivotCaches(i).Refresh()
                        result["actions"].append(
                            {"action": f"pivot_cache_{i}_refresh", "status": "ok"}
                        )
                    except Exception as e:
                        result["actions"].append(
                            {
                                "action": f"pivot_cache_{i}_refresh",
                                "status": "error",
                                "message": str(e),
                            }
                        )
            except Exception as e:
                result["actions"].append(
                    {"action": "pivot_refresh", "status": "error", "message": str(e)}
                )

        elif field_type == "pivot":
            try:
                count = wb.PivotCaches.Count
                for i in range(1, count + 1):
                    try:
                        wb.PivotCaches(i).Refresh()
                        result["actions"].append(
                            {"action": f"pivot_cache_{i}_refresh", "status": "ok"}
                        )
                    except Exception as e:
                        result["actions"].append(
                            {
                                "action": f"pivot_cache_{i}_refresh",
                                "status": "error",
                                "message": str(e),
                            }
                        )
                if count == 0:
                    result["actions"].append(
                        {"action": "pivot_refresh", "status": "skip", "message": "工作簿中无透视表缓存"}
                    )
            except Exception as e:
                result["actions"].append(
                    {"action": "pivot_refresh", "status": "error", "message": str(e)}
                )

        return result

    # ── 语义视图与诊断 ──

    def summarize(self, app: Any) -> dict:
        """生成工作簿结构摘要（L1 语义视图）

        设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

        返回包含工作表列表、图表、命名区域、数据预览等结构化信息的 dict。
        """
        wb = app.ActiveWorkbook

        # 元数据
        metadata = {
            "title": str(wb.BuiltInDocumentProperties("Title").Value or ""),
            "author": str(wb.BuiltInDocumentProperties("Author").Value or ""),
            "sheets": wb.Sheets.Count,
        }

        # 工作表概览
        sheets = []
        for i in range(1, wb.Sheets.Count + 1):
            ws = wb.Sheets(i)
            sheet_info: dict = {
                "index": i,
                "name": ws.Name,
            }
            try:
                used = ws.UsedRange
                if used is not None:
                    last_row = used.Rows.Count
                    last_col = used.Columns.Count
                    sheet_info["used_rows"] = last_row
                    sheet_info["used_cols"] = last_col
                    sheet_info["used_range"] = (
                        f"A1:{_col_letter(last_col)}{last_row}"
                    )
            except Exception:
                pass

            # 数据预览（前5行）
            try:
                preview = []
                used = ws.UsedRange
                if used is not None:
                    preview_rows = min(used.Rows.Count, 5)
                    preview_cols = min(used.Columns.Count, 10)
                    for r in range(1, preview_rows + 1):
                        row_data = []
                        for c in range(1, preview_cols + 1):
                            try:
                                val = used.Cells(r, c).Value
                                row_data.append(str(val) if val is not None else "")
                            except Exception:
                                row_data.append("")
                        preview.append(row_data)
                sheet_info["preview"] = preview
            except Exception:
                pass

            sheets.append(sheet_info)

        # 图表
        charts = []
        for i in range(1, wb.Sheets.Count + 1):
            ws = wb.Sheets(i)
            try:
                for j in range(1, ws.ChartObjects().Count + 1):
                    co = ws.ChartObjects(j)
                    try:
                        chart_title = co.Chart.ChartTitle.Text if co.Chart.HasTitle else ""
                    except Exception:
                        chart_title = ""
                    charts.append(
                        {
                            "sheet": ws.Name,
                            "index": j,
                            "title": chart_title,
                        }
                    )
            except Exception:
                pass

        # 命名区域
        names = []
        for i in range(1, wb.Names.Count + 1):
            try:
                nm = wb.Names(i)
                names.append(
                    {
                        "name": nm.Name,
                        "refers_to": str(nm.RefersTo),
                    }
                )
            except Exception:
                pass

        return {
            "metadata": metadata,
            "sheets": sheets,
            "charts": charts,
            "names": names,
        }

    def diagnose(self, app: Any) -> list[dict]:
        """诊断工作簿问题（参考 OfficeCLI view issues）

        设计参考: iOfficeAI/OfficeCLI (Apache 2.0)
        """
        from wps_cli.services.document_diagnostics import DocumentDiagnostics

        diag = DocumentDiagnostics()
        issues = diag.diagnose_calc(app)
        return [
            {
                "severity": i.severity,
                "category": i.category,
                "subtype": i.subtype,
                "location": i.location,
                "message": i.message,
                "suggestion": i.suggestion,
            }
            for i in issues
        ]

    def annotate(self, app: Any) -> list[str]:
        """输出工作簿带路径标注的摘要（参考 OfficeCLI view annotated）

        设计参考: iOfficeAI/OfficeCLI (Apache 2.0)
        """
        wb = app.ActiveWorkbook
        lines: list[str] = []

        for sheet_idx in range(1, wb.Sheets.Count + 1):
            try:
                ws = wb.Sheets(sheet_idx)
                sheet_name = str(ws.Name)
                lines.append(f"[/workbook/sheet[{sheet_name}]]")

                try:
                    used = ws.UsedRange
                    if used is not None:
                        max_rows = min(used.Rows.Count, 100)
                        max_cols = min(used.Columns.Count, 20)
                        for r in range(1, max_rows + 1):
                            for c in range(1, max_cols + 1):
                                try:
                                    cell = used.Cells(r, c)
                                    addr = cell.Address
                                    val = str(cell.Value) if cell.Value is not None else ""
                                    formula_str = (
                                        str(cell.Formula)
                                        if hasattr(cell, "HasFormula") and cell.HasFormula
                                        else ""
                                    )
                                    prefix = (
                                        f"[/workbook/sheet[{sheet_name}]/cell[{addr}]"
                                        f" formula={formula_str}]"
                                        if formula_str
                                        else f"[/workbook/sheet[{sheet_name}]/cell[{addr}]]"
                                    )
                                    if val and val != "None":
                                        lines.append(f"{prefix} {val[:80]}")
                                except Exception:
                                    pass
                except Exception:
                    pass
            except Exception:
                pass

        return lines

    def get_stats(self, app: Any) -> dict:
        """获取纯数字统计信息（参考 OfficeCLI view stats）

        设计参考: iOfficeAI/OfficeCLI (Apache 2.0)
        """
        wb = app.ActiveWorkbook
        stats: dict = {}

        try:
            stats["sheets"] = wb.Sheets.Count
        except Exception:
            stats["sheets"] = 0
        try:
            stats["names"] = wb.Names.Count
        except Exception:
            stats["names"] = 0

        total_used_rows = 0
        total_used_cols = 0
        total_charts = 0
        total_merged = 0

        for sheet_idx in range(1, stats["sheets"] + 1):
            try:
                ws = wb.Sheets(sheet_idx)
                try:
                    used = ws.UsedRange
                    if used is not None:
                        total_used_rows += used.Rows.Count
                        total_used_cols = max(total_used_cols, used.Columns.Count)
                except Exception:
                    pass
                try:
                    total_charts += ws.ChartObjects().Count
                except Exception:
                    pass
                try:
                    total_merged += ws.UsedRange.MergeCells.Count if ws.UsedRange else 0
                except Exception:
                    pass
            except Exception:
                pass

        stats["total_used_rows"] = total_used_rows
        stats["total_used_cols"] = total_used_cols
        stats["total_charts"] = total_charts
        stats["total_merged"] = total_merged

        return stats
