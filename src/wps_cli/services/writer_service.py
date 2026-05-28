"""Writer 文档操作业务逻辑"""

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from wps_cli.consts import (
    ALIGN_CENTER,
    ALIGN_JUSTIFY,
    ALIGN_LEFT,
    ALIGN_RIGHT,
    WD_DO_NOT_SAVE_CHANGES,
    WD_FORMAT_PDF,
    WD_LINE_SPACE_MULTIPLE,
    WD_PAGE_BREAK,
    WD_REPLACE_ALL,
    WD_SAVE_CHANGES,
    WD_STATISTIC_CHARACTERS,
    WD_STATISTIC_PAGES,
    WD_STATISTIC_WORDS,
    WD_STORY,
)
from wps_cli.services.session_manager import SessionManager


@dataclass
class WriterService:
    """Word 文档操作"""

    manager: SessionManager

    # ── 文档生命周期 ──

    def new(self, output: Path | None = None) -> Path:
        with self.manager.session("writer") as app:
            doc = app.Documents.Add()
            if output:
                doc.SaveAs(str(output))
            path = doc.FullName
            doc.Close(WD_DO_NOT_SAVE_CHANGES)
        return Path(path)

    def open(self, path: Path, readonly: bool = False) -> Any:
        session = self.manager.start("writer")
        try:
            session.app.Documents.Open(str(path), ReadOnly=readonly)
            return session
        except Exception:
            self.manager.stop(session.session_id)
            raise

    def save(self, app: Any, path: Path | None = None) -> Path:
        doc = app.ActiveDocument
        if path:
            doc.SaveAs(str(path))
        else:
            doc.Save()
        return Path(doc.FullName)

    def close(self, app: Any, save: bool = False) -> None:
        doc = app.ActiveDocument
        if save:
            doc.Save()
        doc.Close(WD_DO_NOT_SAVE_CHANGES if not save else WD_SAVE_CHANGES)

    def info(self, path: Path) -> dict:
        with self.manager.session("writer") as app:
            doc = app.Documents.Open(str(path))
            result = {
                "path": str(Path(doc.FullName)),
                "pages": doc.ComputeStatistics(WD_STATISTIC_PAGES),
                "words": doc.ComputeStatistics(WD_STATISTIC_WORDS),
                "characters": doc.ComputeStatistics(WD_STATISTIC_CHARACTERS),
                "paragraphs": doc.Paragraphs.Count,
                "author": doc.BuiltInDocumentProperties("Author").Value,
                "created": str(doc.BuiltInDocumentProperties("Creation Date").Value),
                "modified": str(doc.BuiltInDocumentProperties("Last Save Time").Value),
            }
            doc.Close(WD_DO_NOT_SAVE_CHANGES)
        return result

    # ── 文本操作 ──

    def text_insert(self, app: Any, text: str, position: str = "end") -> None:
        sel = app.Selection
        if position == "end":
            sel.EndKey(WD_STORY)
        sel.TypeText(text)

    def text_replace(
        self, app: Any, old: str, new: str, wildcard: bool = False, case: bool = False
    ) -> int:
        """查找替换文本

        Args:
            wildcard: 启用 WPS 通配符模式（* 任意字符, ? 单字符, [abc] 字符集）
        """
        # 先计数
        rng = app.ActiveDocument.Content
        rng.Find.Text = old
        rng.Find.MatchCase = case
        rng.Find.MatchWildcards = wildcard
        count = 0
        while rng.Find.Execute():
            count += 1
        # 再替换
        find = app.ActiveDocument.Content.Find
        find.Text = old
        find.Replacement.Text = new
        find.MatchCase = case
        find.MatchWildcards = wildcard
        find.Execute(Replace=WD_REPLACE_ALL)
        return count

    def text_get(self, app: Any, start: int = 0, end: int = -1) -> str:
        doc = app.ActiveDocument
        rng = doc.Range(start, end if end >= 0 else doc.Range().End)
        return rng.Text

    def text_count(self, app: Any) -> dict:
        doc = app.ActiveDocument
        return {
            "words": doc.ComputeStatistics(WD_STATISTIC_WORDS),
            "characters": doc.ComputeStatistics(WD_STATISTIC_CHARACTERS),
            "paragraphs": doc.Paragraphs.Count,
            "pages": doc.ComputeStatistics(WD_STATISTIC_PAGES),
        }

    # ── 段落操作 ──

    def heading_insert(self, app: Any, text: str, level: int = 1) -> None:
        sel = app.Selection
        sel.Style = f"标题 {level}"
        sel.TypeText(text)
        sel.TypeParagraph()

    def paragraph_format(
        self,
        app: Any,
        align: str | None = None,
        indent_left: float | None = None,
        indent_first: float | None = None,
        line_spacing: float | None = None,
    ) -> None:
        pf = app.Selection.ParagraphFormat
        align_map = {"left": ALIGN_LEFT, "center": ALIGN_CENTER, "right": ALIGN_RIGHT, "justify": ALIGN_JUSTIFY}
        if align is not None:
            pf.Alignment = align_map.get(align, ALIGN_LEFT)
        if indent_left is not None:
            pf.LeftIndent = indent_left
        if indent_first is not None:
            pf.FirstLineIndent = indent_first
        if line_spacing is not None:
            pf.LineSpacingRule = WD_LINE_SPACE_MULTIPLE
            pf.LineSpacing = line_spacing * 12

    # ── 表格操作 ──

    def table_insert(
        self, app: Any, rows: int, cols: int, data: list[list[str]] | None = None
    ) -> int:
        doc = app.ActiveDocument
        rng = app.Selection.Range
        table = doc.Tables.Add(rng, rows, cols)
        table.Borders.Enable = True
        if data:
            for i, row_data in enumerate(data):
                for j, cell_text in enumerate(row_data):
                    table.Cell(i + 1, j + 1).Range.Text = str(cell_text)
        return table.Index

    def table_get(self, app: Any, index: int) -> list[list[str]]:
        doc = app.ActiveDocument
        table = doc.Tables(index)
        result = []
        for i in range(1, table.Rows.Count + 1):
            row = []
            for j in range(1, table.Columns.Count + 1):
                row.append(table.Cell(i, j).Range.Text.strip())
            result.append(row)
        return result

    # ── 图片操作 ──

    def image_insert(
        self,
        app: Any,
        path: Path,
        width: float | None = None,
        height: float | None = None,
    ) -> None:
        sel = app.Selection
        shape = sel.InlineShapes.AddPicture(str(path))
        if width is not None:
            shape.Width = width
        if height is not None:
            shape.Height = height

    # ── 页面布局 ──

    def page_setup(
        self,
        app: Any,
        width_mm: float | None = None,
        height_mm: float | None = None,
        margin_top: float | None = None,
        margin_bottom: float | None = None,
        margin_left: float | None = None,
        margin_right: float | None = None,
    ) -> None:
        page = app.ActiveDocument.PageSetup
        MM_TO_PT = 2.835
        if width_mm is not None:
            page.PageWidth = width_mm * MM_TO_PT
        if height_mm is not None:
            page.PageHeight = height_mm * MM_TO_PT
        if margin_top is not None:
            page.TopMargin = margin_top * MM_TO_PT
        if margin_bottom is not None:
            page.BottomMargin = margin_bottom * MM_TO_PT
        if margin_left is not None:
            page.LeftMargin = margin_left * MM_TO_PT
        if margin_right is not None:
            page.RightMargin = margin_right * MM_TO_PT

    def page_break(self, app: Any) -> None:
        app.Selection.InsertBreak(WD_PAGE_BREAK)

    # ── 导出 ──

    def export_pdf(self, app: Any, output: Path) -> Path:
        doc = app.ActiveDocument
        doc.ExportAsFixedFormat(str(output), WD_FORMAT_PDF)
        return output
