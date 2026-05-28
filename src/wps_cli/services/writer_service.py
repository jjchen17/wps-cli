"""Writer 文档操作业务逻辑"""

from __future__ import annotations

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
from wps_cli.services.session_manager import Session, SessionManager


@dataclass
class WriterService:
    """Word 文档操作"""

    manager: SessionManager

    @staticmethod
    def _open_doc(app: Any, path: Path | str, *, readonly: bool = False) -> Any:
        """统一的 Documents.Open 入口

        ``ConfirmConversions=False`` + ``AddToRecentFiles=False`` 减少弹窗，
        宏自动执行已在 ``ComBackend.harden`` 中通过 ``AutomationSecurity`` 全局禁用。
        """
        return app.Documents.Open(
            str(path),
            ConfirmConversions=False,
            ReadOnly=readonly,
            AddToRecentFiles=False,
        )

    # ── 文档生命周期 ──

    def new(self, output: Path | None = None) -> Path:
        with self.manager.session("writer") as app:
            doc = app.Documents.Add()
            if output:
                doc.SaveAs(str(output))
            path = doc.FullName
            doc.Close(WD_DO_NOT_SAVE_CHANGES)
        return Path(path)

    def open_document(self, path: Path, readonly: bool = False) -> Session:
        """打开文档并返回会话；命名避免遮蔽内置 ``open``"""
        session = self.manager.start("writer")
        try:
            self._open_doc(session.app, path, readonly=readonly)
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
            doc = self._open_doc(app, path, readonly=True)
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
        """查找替换文本，返回替换次数

        Args:
            wildcard: 启用 WPS 通配符模式（* 任意字符, ? 单字符, [abc] 字符集）

        非通配符模式下用 Find API 逐次计数（精确，不受 ``new`` 是否包含 ``old`` 影响）。
        通配符模式下计数语义复杂（同一位置可被不同通配规则反复匹配），返回 -1 表示未知。
        """
        doc = app.ActiveDocument

        find = doc.Content.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        find.Text = old
        find.Replacement.Text = new
        find.MatchCase = case
        find.MatchWildcards = wildcard
        find.Forward = True
        find.Wrap = 0  # wdFindStop 防止 wrap 导致重复计数

        if not wildcard:
            count = 0
            scan = doc.Content.Find
            scan.ClearFormatting()
            scan.Text = old
            scan.MatchCase = case
            scan.MatchWildcards = False
            scan.Forward = True
            scan.Wrap = 0
            while scan.Execute(Replace=0):
                count += 1
            find.Execute(Replace=WD_REPLACE_ALL)
            return count

        find.Execute(Replace=WD_REPLACE_ALL)
        return -1

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
        align_map = {
            "left": ALIGN_LEFT,
            "center": ALIGN_CENTER,
            "right": ALIGN_RIGHT,
            "justify": ALIGN_JUSTIFY,
        }
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
        mm_to_pt = 2.835
        if width_mm is not None:
            page.PageWidth = width_mm * mm_to_pt
        if height_mm is not None:
            page.PageHeight = height_mm * mm_to_pt
        if margin_top is not None:
            page.TopMargin = margin_top * mm_to_pt
        if margin_bottom is not None:
            page.BottomMargin = margin_bottom * mm_to_pt
        if margin_left is not None:
            page.LeftMargin = margin_left * mm_to_pt
        if margin_right is not None:
            page.RightMargin = margin_right * mm_to_pt

    def page_break(self, app: Any) -> None:
        app.Selection.InsertBreak(WD_PAGE_BREAK)

    # ── 导出 ──

    def export_pdf(self, app: Any, output: Path) -> Path:
        doc = app.ActiveDocument
        doc.ExportAsFixedFormat(str(output), WD_FORMAT_PDF)
        return output
