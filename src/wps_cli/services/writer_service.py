"""Writer 文档操作业务逻辑"""

import re
from dataclasses import dataclass
from pathlib import Path

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
            doc.Close(0)  # wdDoNotSaveChanges
        return Path(path)

    def open(self, path: Path, readonly: bool = False) -> object:
        session = self.manager.start("writer")
        app = session.app
        app.Documents.Open(str(path), ReadOnly=readonly)
        return session

    def save(self, app: object, path: Path | None = None) -> Path:
        doc = app.ActiveDocument
        if path:
            doc.SaveAs(str(path))
        else:
            doc.Save()
        return Path(doc.FullName)

    def close(self, app: object, save: bool = False) -> None:
        doc = app.ActiveDocument
        if save:
            doc.Save()
        doc.Close(0 if not save else -1)  # -1 = wdSaveChanges

    def info(self, path: Path) -> dict:
        with self.manager.session("writer") as app:
            doc = app.Documents.Open(str(path))
            result = {
                "path": str(Path(doc.FullName)),
                "pages": doc.ComputeStatistics(2),  # wdStatisticPages
                "words": doc.ComputeStatistics(0),  # wdStatisticWords
                "characters": doc.ComputeStatistics(3),  # wdStatisticCharacters
                "paragraphs": doc.Paragraphs.Count,
                "author": doc.BuiltInDocumentProperties("Author").Value,
                "created": str(doc.BuiltInDocumentProperties("Creation Date").Value),
                "modified": str(doc.BuiltInDocumentProperties("Last Save Time").Value),
            }
            doc.Close(0)
        return result

    # ── 文本操作 ──

    def text_insert(self, app: object, text: str, position: str = "end") -> None:
        sel = app.Selection
        if position == "end":
            sel.EndKey(6)  # wdStory
        sel.TypeText(text)

    def text_replace(
        self, app: object, old: str, new: str, regex: bool = False, case: bool = False
    ) -> int:
        # 先计数
        rng = app.ActiveDocument.Content
        rng.Find.Text = old
        rng.Find.MatchCase = case
        rng.Find.MatchWildcards = regex
        count = 0
        while rng.Find.Execute():
            count += 1
        # 再替换
        find = app.ActiveDocument.Content.Find
        find.Text = old
        find.Replacement.Text = new
        find.MatchCase = case
        find.MatchWildcards = regex
        find.Execute(Replace=2)  # wdReplaceAll
        return count

    def text_get(self, app: object, start: int = 0, end: int = -1) -> str:
        doc = app.ActiveDocument
        rng = doc.Range(start, end if end >= 0 else doc.Range().End)
        return rng.Text

    def text_count(self, app: object) -> dict:
        doc = app.ActiveDocument
        return {
            "words": doc.ComputeStatistics(0),
            "characters": doc.ComputeStatistics(3),
            "paragraphs": doc.Paragraphs.Count,
            "pages": doc.ComputeStatistics(2),
        }

    # ── 段落操作 ──

    def heading_insert(self, app: object, text: str, level: int = 1) -> None:
        sel = app.Selection
        sel.Style = f"标题 {level}"
        sel.TypeText(text)
        sel.TypeParagraph()

    def paragraph_format(
        self,
        app: object,
        align: str | None = None,
        indent_left: float | None = None,
        indent_first: float | None = None,
        line_spacing: float | None = None,
    ) -> None:
        pf = app.Selection.ParagraphFormat
        align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
        if align is not None:
            pf.Alignment = align_map.get(align, 0)
        if indent_left is not None:
            pf.LeftIndent = indent_left
        if indent_first is not None:
            pf.FirstLineIndent = indent_first
        if line_spacing is not None:
            pf.LineSpacingRule = 4  # wdLineSpaceMultiple
            pf.LineSpacing = line_spacing * 12

    # ── 表格操作 ──

    def table_insert(
        self, app: object, rows: int, cols: int, data: list[list[str]] | None = None
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

    def table_get(self, app: object, index: int) -> list[list[str]]:
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
        app: object,
        path: Path,
        width: float | None = None,
        height: float | None = None,
    ) -> None:
        sel = app.Selection
        shape = sel.InlineShapes.AddPicture(str(path))
        if width:
            shape.Width = width
        if height:
            shape.Height = height

    # ── 页面布局 ──

    def page_setup(
        self,
        app: object,
        width_mm: float | None = None,
        height_mm: float | None = None,
        margin_top: float | None = None,
        margin_bottom: float | None = None,
        margin_left: float | None = None,
        margin_right: float | None = None,
    ) -> None:
        page = app.ActiveDocument.PageSetup
        if width_mm:
            page.PageWidth = width_mm * 2.835
        if height_mm:
            page.PageHeight = height_mm * 2.835
        if margin_top:
            page.TopMargin = margin_top * 2.835
        if margin_bottom:
            page.BottomMargin = margin_bottom * 2.835
        if margin_left:
            page.LeftMargin = margin_left * 2.835
        if margin_right:
            page.RightMargin = margin_right * 2.835

    def page_break(self, app: object) -> None:
        app.Selection.InsertBreak(7)  # wdPageBreak

    # ── 导出 ──

    def export_pdf(self, app: object, output: Path) -> Path:
        doc = app.ActiveDocument
        doc.ExportAsFixedFormat(str(output), 17)  # wdExportFormatPDF
        return output
