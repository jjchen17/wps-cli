# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""统一路径解析器

将 OfficeCLI 风格的路径语法解析为 COM 对象导航操作。
路径语法（1-based，逗号分隔子索引）:
  Word:   /section[1]/paragraph[3]/table[2]/cell[2,1]
  Excel:  /sheet["Sheet1"]/range["A1:B10"]/cell["C3"]
  PPT:    /slide[1]/shape[3]/text[1]

设计原则：
- 不需要 XML 命名空间知识
- 基于元素本地名称
- 路径在文档修改间保持稳定
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any


@dataclass
class PathComponent:
    """路径组件"""

    element: str  # 元素类型: section, paragraph, table, cell, slide, shape, sheet, range
    index: int | str | tuple[int, int] | None  # 索引


@dataclass
class PathResolver:
    """统一路径解析器 — 将路径字符串解析为 COM 对象引用"""

    # ── 通用 ──

    @staticmethod
    def parse(path: str) -> list[PathComponent]:
        """将路径字符串解析为组件列表

        格式: /element[idx]/element[idx]/...
        支持的数字索引: /slide[1]
        支持的字符串索引: /sheet["Sheet1"]  或 /sheet['Sheet1']
        支持的元组索引: /cell[2,1]

        Raises:
            ValidationError: 路径格式无效时抛出
        """
        from wps_cli.exceptions import ValidationError

        if not path or not path.startswith("/"):
            raise ValidationError(
                f"路径必须以 '/' 开头: {path!r}",
                suggestion="有效示例: /slide[1]/shape[2], /section[1]/paragraph[3]",
            )

        components: list[PathComponent] = []
        # 匹配 /element[index] 模式
        # index 可以是数字、带引号字符串、或逗号分隔的数字对
        pattern = re.compile(
            r"/([a-zA-Z_]+)"  # 元素名
            r"(?:\["  # 可选索引
            r"("
            r'(?:"[^"]*"|\'[^\']*\')'  # 字符串索引: "Sheet1" 或 'Sheet1'
            r"|"  # 或
            r"\s*\d+(?:\s*,\s*\d+)?\s*"  # 数字索引: 1 或 2,1（允许空格）
            r")"
            r"\])?",  # 结束索引
        )

        for m in pattern.finditer(path):
            element = m.group(1).lower()
            raw_index = m.group(2)

            if raw_index is None:
                # 无索引，如 /body
                index = None
            elif raw_index.startswith('"') or raw_index.startswith("'"):
                # 字符串索引，去掉外层引号
                index = raw_index[1:-1]
            else:
                # 数字索引
                parts = [x.strip() for x in raw_index.split(",")]
                if len(parts) == 1:
                    index = int(parts[0])
                elif len(parts) == 2:
                    index = (int(parts[0]), int(parts[1]))
                else:
                    raise ValidationError(
                        f"索引格式无效: [{raw_index}] — 最多支持两个维度 (row,col)"
                    )

            components.append(PathComponent(element=element, index=index))

        if not components:
            raise ValidationError(
                f"无法解析路径: {path!r}",
                suggestion="有效示例: /slide[1]/shape[2], /section[1]/paragraph[3]",
            )

        return components

    @staticmethod
    def _check_index(component: PathComponent, max_val: int, path: str) -> None:
        """校验 1-based 索引范围，越界时抛出带建议的 ValidationError"""
        from wps_cli.exceptions import ValidationError

        if component.index is None:
            return
        idx = component.index
        if isinstance(idx, int) and (idx < 1 or idx > max_val):
            effective = list(range(1, max_val + 1))
            raise ValidationError(
                f"路径 {path!r} 中 {component.element}[{idx}] 越界"
                f"（有效范围: {effective[0]}-{effective[-1]}）",
                suggestion=f"可用索引: {effective[:5]}{'...' if len(effective) > 5 else ''}",
            )

    # ── Writer ──

    def resolve_writer(self, app: Any, path: str) -> Any:
        """在 Word 文档中解析路径，返回 COM Range 或对象

        支持的路径格式：
          /body                              → 文档正文
          /section[1]                        → 第1节
          /section[1]/paragraph[3]           → 第1节第3段
          /section[1]/table[2]               → 第1节第2个表格
          /section[1]/table[2]/cell[3,1]     → 第1节第2个表格的第3行第1列
          /header[1]                         → 第1节页眉
          /footer[1]                         → 第1节页脚
        """
        from wps_cli.exceptions import ValidationError

        components = self.parse(path)
        doc = app.ActiveDocument

        if len(components) == 1 and components[0].element == "body" and components[0].index is None:
            return doc.Content

        current = None
        for comp in components:
            element = comp.element

            if element == "section":
                count = doc.Sections.Count
                self._check_index(comp, count, path)
                current = doc.Sections(comp.index)

            elif element == "paragraph":
                if current is None:
                    current = doc
                count = current.Paragraphs.Count
                self._check_index(comp, count, path)
                current = current.Paragraphs(comp.index)

            elif element == "table":
                if current is None:
                    current = doc
                count = current.Tables.Count
                self._check_index(comp, count, path)
                current = current.Tables(comp.index)

            elif element == "cell":
                if current is None or not hasattr(current, "Cell"):
                    raise ValidationError(
                        f"路径 {path!r}: cell 必须跟在 table 后面",
                        suggestion="有效示例: /section[1]/table[2]/cell[3,1]",
                    )
                if isinstance(comp.index, tuple):
                    row, col = comp.index
                else:
                    raise ValidationError(
                        f"路径 {path!r}: cell 需要 (row, col) 索引",
                        suggestion="有效示例: /cell[3,1]",
                    )
                self._check_index(
                    PathComponent("row", row), current.Rows.Count, path
                )
                self._check_index(
                    PathComponent("col", col), current.Columns.Count, path
                )
                current = current.Cell(row, col)

            elif element == "header":
                count = doc.Sections.Count
                self._check_index(comp, count, path)
                section = doc.Sections(comp.index)
                current = section.Headers(1)  # wdHeaderFooterPrimary

            elif element == "footer":
                count = doc.Sections.Count
                self._check_index(comp, count, path)
                section = doc.Sections(comp.index)
                current = section.Footers(1)

            elif element == "body":
                current = doc

            else:
                raise ValidationError(
                    f"路径 {path!r}: Writer 不支持元素类型 '{element}'",
                    suggestion="支持的 Writer 元素: body, section, paragraph, table, cell, header, footer",
                )

        if current is None:
            raise ValidationError(
                f"路径 {path!r}: 无法解析到有效对象",
                suggestion="路径至少需要包含一个元素",
            )

        return current

    # ── Calc ──

    def resolve_calc(self, app: Any, path: str) -> Any:
        """在 Excel 工作簿中解析路径，返回 Range 或对象

        支持的路径格式：
          /sheet["Sheet1"]                       → 工作表 Sheet1
          /sheet["Sheet1"]/cell["C12"]           → Sheet1 的 C12 单元格
          /sheet["Sheet1"]/range["A1:B10"]       → Sheet1 的 A1:B10 区域
          /sheet[1]                              → 第1个工作表
          /sheet[1]/range["A1:B10"]/cell["A1"]   → 区域中第1个单元格
          /chart[1]                              → 第1个图表

        也支持 Excel 风格的简写: $Sheet1:A1
        """
        from wps_cli.exceptions import ValidationError

        # 处理 Excel 风格简写: $SheetName:Ref
        if path.startswith("$"):
            m = re.match(r"^\$([^:]+):(.+)$", path)
            if m:
                sheet_name, ref = m.group(1), m.group(2)
                components = [
                    PathComponent(element="sheet", index=sheet_name),
                    PathComponent(
                        element="cell" if re.match(r"^[A-Za-z]+\d+$", ref) else "range",
                        index=ref,
                    ),
                ]
            else:
                raise ValidationError(
                    f"无效的 Excel 风格路径: {path!r}",
                    suggestion="有效示例: $Sheet1:A1, $Sheet1:A1:B10",
                )
        else:
            components = self.parse(path)

        wb = app.ActiveWorkbook
        current = None
        ws = None  # 跟踪当前工作表，用于 cell/range

        for comp in components:
            element = comp.element

            if element == "sheet":
                if isinstance(comp.index, str):
                    # 按名称访问
                    try:
                        ws = wb.Sheets(comp.index)
                    except Exception as exc:
                        sheet_names = [
                            wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)
                        ]
                        raise ValidationError(
                            f"路径 {path!r}: 工作表 '{comp.index}' 不存在",
                            suggestion=f"可用工作表: {sheet_names}",
                        ) from exc
                elif isinstance(comp.index, int):
                    count = wb.Sheets.Count
                    self._check_index(comp, count, path)
                    ws = wb.Sheets(comp.index)
                else:
                    raise ValidationError(
                        f"路径 {path!r}: sheet 需要名称或数字索引",
                        suggestion="有效示例: /sheet[\"Sheet1\"] 或 /sheet[1]",
                    )
                current = ws

            elif element == "cell":
                if ws is None:
                    ws = app.ActiveSheet
                ref = comp.index
                if not isinstance(ref, str) or not re.match(r"^[A-Za-z]+\d+$", ref):
                    raise ValidationError(
                        f"路径 {path!r}: cell 需要 Excel 单元格引用 (如 'C12')",
                        suggestion="有效示例: /sheet[\"Sheet1\"]/cell[\"C12\"]",
                    )
                try:
                    current = ws.Range(ref)
                except Exception as exc:
                    raise ValidationError(
                        f"路径 {path!r}: 无法访问单元格 {ref}",
                        suggestion="请检查单元格引用格式 (如 A1, B3)",
                    ) from exc

            elif element == "range":
                if ws is None:
                    ws = app.ActiveSheet
                ref = comp.index
                if not isinstance(ref, str) or ":" not in ref:
                    raise ValidationError(
                        f"路径 {path!r}: range 需要区域引用 (如 'A1:B10')",
                        suggestion="有效示例: /sheet[\"Sheet1\"]/range[\"A1:B10\"]",
                    )
                try:
                    current = ws.Range(ref)
                except Exception as exc:
                    raise ValidationError(
                        f"路径 {path!r}: 无法访问区域 {ref}",
                        suggestion="请检查区域引用格式 (如 A1:B10)",
                    ) from exc

            elif element == "chart":
                if ws is None:
                    ws = app.ActiveSheet
                count = ws.ChartObjects().Count
                self._check_index(comp, count, path)
                current = ws.ChartObjects(comp.index)

            else:
                raise ValidationError(
                    f"路径 {path!r}: Calc 不支持元素类型 '{element}'",
                    suggestion="支持的 Calc 元素: sheet, cell, range, chart",
                )

        if current is None:
            raise ValidationError(
                f"路径 {path!r}: 无法解析到有效对象",
                suggestion="路径至少需要包含一个元素",
            )

        return current

    # ── Impress ──

    def resolve_impress(self, app: Any, path: str) -> Any:
        """在 PPT 演示文稿中解析路径，返回 Shape 或对象

        支持的路径格式：
          /slide[1]                       → 第1张幻灯片
          /slide[1]/shape[2]              → 第1张幻灯片的第2个形状
          /slide[1]/shape[2]/text[1]      → 第2个形状的第1段文本
          /slide[3]/notes                 → 第3张的备注
        """
        from wps_cli.exceptions import ValidationError

        components = self.parse(path)
        pres = app.ActivePresentation
        current = None

        for comp in components:
            element = comp.element

            if element == "slide":
                count = pres.Slides.Count
                self._check_index(comp, count, path)
                current = pres.Slides(comp.index)

            elif element == "shape":
                if current is None or not hasattr(current, "Shapes"):
                    raise ValidationError(
                        f"路径 {path!r}: shape 必须跟在 slide 后面",
                        suggestion="有效示例: /slide[1]/shape[2]",
                    )
                count = current.Shapes.Count
                self._check_index(comp, count, path)
                current = current.Shapes(comp.index)

            elif element == "text":
                if current is None or not hasattr(current, "TextFrame"):
                    raise ValidationError(
                        f"路径 {path!r}: text 必须跟在 shape 后面",
                        suggestion="有效示例: /slide[1]/shape[2]/text[1]",
                    )
                if isinstance(comp.index, int):
                    try:
                        # 尝试获取 TextRange.Paragraphs(comp.index)
                        tr = current.TextFrame.TextRange
                        count = tr.Paragraphs().Count if hasattr(tr, "Paragraphs") else 1
                        current = tr.Paragraphs(comp.index) if comp.index <= count else tr
                    except Exception:
                        current = current.TextFrame.TextRange
                else:
                    current = current.TextFrame.TextRange

            elif element == "notes":
                if current is None or not hasattr(current, "NotesPage"):
                    raise ValidationError(
                        f"路径 {path!r}: notes 必须跟在 slide 后面",
                        suggestion="有效示例: /slide[3]/notes",
                    )
                current = current.NotesPage

            else:
                raise ValidationError(
                    f"路径 {path!r}: Impress 不支持元素类型 '{element}'",
                    suggestion="支持的 Impress 元素: slide, shape, text, notes",
                )

        if current is None:
            raise ValidationError(
                f"路径 {path!r}: 无法解析到有效对象",
                suggestion="路径至少需要包含一个元素",
            )

        return current

    # ── 统一入口 ──

    def resolve(self, app: Any, app_type: str, path: str) -> Any:
        """统一入口：根据 app_type 路由到对应解析器

        Args:
            app: COM application 对象
            app_type: "writer" | "calc" | "impress"
            path: 路径字符串

        Returns:
            解析到的 COM 对象

        Raises:
            ValidationError: 路径无效或 app_type 不支持时抛出
        """
        from wps_cli.exceptions import ValidationError

        resolvers = {
            "writer": self.resolve_writer,
            "calc": self.resolve_calc,
            "impress": self.resolve_impress,
        }
        if app_type not in resolvers:
            raise ValidationError(
                f"不支持的应用类型: {app_type!r}",
                suggestion="支持的类型: writer, calc, impress",
            )
        return resolvers[app_type](app, path)


# 模块级单例，方便直接使用
resolver = PathResolver()
