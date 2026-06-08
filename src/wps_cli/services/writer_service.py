"""Writer 文档操作业务逻辑"""

from __future__ import annotations

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

    # ── 模板合并 ──

    @staticmethod
    def template_fill(app: Any, data: dict[str, str]) -> dict:
        """替换文档中所有 {{key}} 占位符为实际值。

        设计参考: iOfficeAI/OfficeCLI (Apache 2.0)
        """
        from wps_cli.services.template_engine import TemplateEngine

        engine = TemplateEngine()
        return engine.fill(app, data)

    # ── 导出 ──

    def export_pdf(self, app: Any, output: Path) -> Path:
        doc = app.ActiveDocument
        doc.ExportAsFixedFormat(str(output), WD_FORMAT_PDF)
        return output

    # ── Refresh 刷新 ──
    # 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

    def refresh_fields(self, app: Any, field_type: str | None = None) -> dict:
        """刷新文档字段

        Args:
            field_type: 字段类型过滤 (toc/page/all), None 或 "all" 表示全部
        """
        doc = app.ActiveDocument
        result: dict = {"field_type": field_type or "all", "actions": []}

        if field_type is None or field_type == "all":
            try:
                count = doc.Fields.Count
                doc.Fields.Update()
                result["actions"].append(
                    {"action": "fields_update", "count": count, "status": "ok"}
                )
            except Exception as e:
                result["actions"].append(
                    {"action": "fields_update", "status": "error", "message": str(e)}
                )

            try:
                toc_count = doc.TablesOfContents.Count
                for i in range(1, toc_count + 1):
                    try:
                        doc.TablesOfContents(i).Update()
                        result["actions"].append(
                            {"action": f"toc_{i}_update", "status": "ok"}
                        )
                    except Exception as e:
                        result["actions"].append(
                            {
                                "action": f"toc_{i}_update",
                                "status": "error",
                                "message": str(e),
                            }
                        )
            except Exception as e:
                result["actions"].append(
                    {"action": "toc_update", "status": "error", "message": str(e)}
                )

        elif field_type == "toc":
            try:
                toc_count = doc.TablesOfContents.Count
                for i in range(1, toc_count + 1):
                    try:
                        doc.TablesOfContents(i).Update()
                        result["actions"].append(
                            {"action": f"toc_{i}_update", "status": "ok"}
                        )
                    except Exception as e:
                        result["actions"].append(
                            {
                                "action": f"toc_{i}_update",
                                "status": "error",
                                "message": str(e),
                            }
                        )
                if toc_count == 0:
                    result["actions"].append(
                        {"action": "toc_update", "status": "skip", "message": "文档中无目录"}
                    )
            except Exception as e:
                result["actions"].append(
                    {"action": "toc_update", "status": "error", "message": str(e)}
                )

        elif field_type == "page":
            try:
                doc.Fields.Update()
                result["actions"].append(
                    {"action": "page_fields_update", "status": "ok"}
                )
            except Exception as e:
                result["actions"].append(
                    {"action": "page_fields_update", "status": "error", "message": str(e)}
                )

        return result

    # ── 表单域与内容控件 ──
    # 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

    def formfield_list(self, app: Any) -> list[dict]:
        """列出所有表单域（旧式 FormFields）"""
        doc = app.ActiveDocument
        fields: list[dict] = []
        try:
            total = doc.FormFields.Count
            for i in range(1, total + 1):
                try:
                    ff = doc.FormFields(i)
                    ff_type = ""
                    try:
                        type_val = ff.Type
                        # wdFieldFormTextInput=70, wdFieldFormCheckBox=71, wdFieldFormDropDown=83
                        type_map = {70: "text", 71: "checkbox", 83: "dropdown"}
                        ff_type = type_map.get(type_val, f"unknown({type_val})")
                    except Exception:
                        ff_type = "unknown"

                    result_text = ""
                    try:
                        result_text = str(ff.Result) if ff.Result else ""
                    except Exception:
                        pass

                    name = ""
                    try:
                        name = str(ff.Name) if ff.Name else ""
                    except Exception:
                        pass

                    fields.append(
                        {
                            "index": i,
                            "name": name,
                            "type": ff_type,
                            "result": result_text,
                        }
                    )
                except Exception:
                    fields.append(
                        {
                            "index": i,
                            "name": "",
                            "type": "unknown",
                            "result": "",
                            "error": "无法读取表单域信息",
                        }
                    )
        except Exception:
            pass
        return fields

    def formfield_get(self, app: Any, index: int) -> dict:
        """获取指定表单域信息"""
        doc = app.ActiveDocument
        ff = doc.FormFields(index)

        ff_type = ""
        try:
            type_val = ff.Type
            type_map = {70: "text", 71: "checkbox", 83: "dropdown"}
            ff_type = type_map.get(type_val, f"unknown({type_val})")
        except Exception:
            ff_type = "unknown"

        result_text = ""
        try:
            result_text = str(ff.Result) if ff.Result else ""
        except Exception:
            pass

        name = ""
        try:
            name = str(ff.Name) if ff.Name else ""
        except Exception:
            pass

        help_text = ""
        try:
            help_text = str(ff.StatusText) if ff.StatusText else ""
        except Exception:
            pass

        return {
            "index": index,
            "name": name,
            "type": ff_type,
            "result": result_text,
            "help_text": help_text,
        }

    def formfield_set(self, app: Any, index: int, value: str) -> None:
        """设置表单域值"""
        doc = app.ActiveDocument
        ff = doc.FormFields(index)
        ff.Result = value

    def content_control_list(self, app: Any) -> list[dict]:
        """列出所有内容控件（ContentControls）"""
        doc = app.ActiveDocument
        controls: list[dict] = []
        try:
            total = doc.ContentControls.Count
            for i in range(1, total + 1):
                try:
                    cc = doc.ContentControls(i)
                    cc_type = ""
                    try:
                        type_val = cc.Type
                        # wdContentControlRichText=0, wdContentControlText=1,
                        # wdContentControlPicture=2, wdContentControlComboBox=3,
                        # wdContentControlDropdownList=4, wdContentControlBuildingBlockGallery=5,
                        # wdContentControlDate=6, wdContentControlCheckBox=7
                        type_map = {
                            0: "rich_text",
                            1: "plain_text",
                            2: "picture",
                            3: "combobox",
                            4: "dropdown",
                            5: "building_block",
                            6: "date",
                            7: "checkbox",
                        }
                        cc_type = type_map.get(type_val, f"unknown({type_val})")
                    except Exception:
                        cc_type = "unknown"

                    text = ""
                    try:
                        text = str(cc.Range.Text).strip()[:200]
                    except Exception:
                        pass

                    title = ""
                    try:
                        title = str(cc.Title) if cc.Title else ""
                    except Exception:
                        pass

                    tag = ""
                    try:
                        tag = str(cc.Tag) if cc.Tag else ""
                    except Exception:
                        pass

                    lock = ""
                    try:
                        if cc.LockContents:
                            lock = "content_locked"
                        elif cc.LockContentControl:
                            lock = "cannot_delete"
                    except Exception:
                        pass

                    controls.append(
                        {
                            "index": i,
                            "title": title,
                            "tag": tag,
                            "type": cc_type,
                            "text": text,
                            "lock": lock,
                        }
                    )
                except Exception:
                    controls.append(
                        {
                            "index": i,
                            "title": "",
                            "tag": "",
                            "type": "unknown",
                            "text": "",
                            "lock": "",
                            "error": "无法读取内容控件信息",
                        }
                    )
        except Exception:
            pass
        return controls

    def content_control_set(self, app: Any, index: int, text: str) -> None:
        """设置内容控件文本"""
        doc = app.ActiveDocument
        cc = doc.ContentControls(index)
        cc.Range.Text = text

    # ── 语义视图与诊断 ──

    def summarize(self, app: Any) -> dict:
        """生成文档结构摘要（L1 语义视图）

        设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

        返回包含标题层级、表格、图片等结构化信息的 dict，便于 AI Agent 理解文档全貌。
        """
        doc = app.ActiveDocument

        # 元数据
        metadata = {
            "title": str(doc.BuiltInDocumentProperties("Title").Value or ""),
            "author": str(doc.BuiltInDocumentProperties("Author").Value or ""),
            "pages": doc.ComputeStatistics(WD_STATISTIC_PAGES),
            "words": doc.ComputeStatistics(WD_STATISTIC_WORDS),
            "paragraphs": doc.Paragraphs.Count,
        }

        # 标题
        headings = []
        for p in doc.Paragraphs:
            try:
                style_name = str(p.Style)
                if "标题" in style_name or "Heading" in style_name or "heading" in style_name:
                    headings.append(
                        {
                            "level": self._heading_level(style_name),
                            "text": p.Range.Text.strip()[:100],
                            "page": p.Range.Information(3),  # wdActiveEndPageNumber
                        }
                    )
            except Exception:
                pass

        # 表格
        tables = []
        for i in range(1, doc.Tables.Count + 1):
            try:
                t = doc.Tables(i)
                tables.append(
                    {
                        "index": i,
                        "rows": t.Rows.Count,
                        "cols": t.Columns.Count,
                    }
                )
            except Exception:
                pass

        # 图片
        images = []
        for i in range(1, doc.InlineShapes.Count + 1):
            try:
                shape = doc.InlineShapes(i)
                images.append(
                    {
                        "index": i,
                        "type": "InlineShape",
                        "width": shape.Width,
                        "height": shape.Height,
                        "has_alt_text": bool(shape.AlternativeText),
                    }
                )
            except Exception:
                pass

        return {
            "metadata": metadata,
            "headings": headings,
            "tables": tables,
            "images": images,
        }

    def diagnose(self, app: Any) -> list[dict]:
        """诊断文档问题（参考 OfficeCLI view issues）

        设计参考: iOfficeAI/OfficeCLI (Apache 2.0)
        """
        from wps_cli.services.document_diagnostics import DocumentDiagnostics

        diag = DocumentDiagnostics()
        issues = diag.diagnose_writer(app)
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
        """输出文档内容并标注每个元素的路径和样式（参考 OfficeCLI view annotated）

        设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

        遍历文档段落/表格/InlineShape，每行前缀标注路径和样式，便于 AI Agent 定位元素。
        """
        doc = app.ActiveDocument
        lines: list[str] = []
        para_idx = 0

        try:
            for obj in doc.StoryRanges:
                try:
                    for para in obj.Paragraphs:
                        try:
                            para_idx += 1
                            style_name = str(para.Style)
                            text = para.Range.Text.strip()
                            if text and text not in ("\r", "\x0c", "\x0d"):
                                lines.append(
                                    f"[/section[1]/paragraph[{para_idx}] "
                                    f"style={style_name}] {text}"
                                )
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception:
            # 回退：直接遍历 Paragraphs
            try:
                for para in doc.Paragraphs:
                    try:
                        para_idx += 1
                        style_name = str(para.Style)
                        text = para.Range.Text.strip()
                        if text and text not in ("\r", "\x0c", "\x0d"):
                            lines.append(
                                f"[/body/section[1]/paragraph[{para_idx}] "
                                f"style={style_name}] {text}"
                            )
                    except Exception:
                        pass
            except Exception:
                pass

        # 表格
        try:
            for i in range(1, doc.Tables.Count + 1):
                try:
                    t = doc.Tables(i)
                    lines.append(
                        f"[/body/section[1]/table[{i}] "
                        f"rows={t.Rows.Count} cols={t.Columns.Count}]"
                    )
                except Exception:
                    pass
        except Exception:
            pass

        # InlineShape
        try:
            for i in range(1, doc.InlineShapes.Count + 1):
                try:
                    shape = doc.InlineShapes(i)
                    alt = str(shape.AlternativeText) if shape.AlternativeText else "(无)"
                    lines.append(
                        f"[/body/section[1]/inline_shape[{i}] "
                        f"width={shape.Width} height={shape.Height} "
                        f"alt_text={alt}]"
                    )
                except Exception:
                    pass
        except Exception:
            pass

        return lines

    def get_stats(self, app: Any) -> dict:
        """获取纯数字统计信息（参考 OfficeCLI view stats）

        设计参考: iOfficeAI/OfficeCLI (Apache 2.0)
        """
        doc = app.ActiveDocument
        stats: dict = {}

        try:
            stats["pages"] = doc.ComputeStatistics(WD_STATISTIC_PAGES)
        except Exception:
            stats["pages"] = 0
        try:
            stats["words"] = doc.ComputeStatistics(WD_STATISTIC_WORDS)
        except Exception:
            stats["words"] = 0
        try:
            stats["characters"] = doc.ComputeStatistics(WD_STATISTIC_CHARACTERS)
        except Exception:
            stats["characters"] = 0
        try:
            stats["paragraphs"] = doc.Paragraphs.Count
        except Exception:
            stats["paragraphs"] = 0
        try:
            stats["tables"] = doc.Tables.Count
        except Exception:
            stats["tables"] = 0
        try:
            stats["inline_shapes"] = doc.InlineShapes.Count
        except Exception:
            stats["inline_shapes"] = 0

        # 统计字体种类
        try:
            fonts: set[str] = set()
            for p in doc.Paragraphs:
                try:
                    fn = str(p.Range.Font.Name)
                    if fn:
                        fonts.add(fn)
                except Exception:
                    pass
            stats["fonts_used"] = len(fonts)
        except Exception:
            stats["fonts_used"] = 0

        return stats

    @staticmethod
    def _heading_level(style_name: str) -> int:
        """从样式名提取标题级别"""
        m = re.search(r"(\d+)", style_name)
        return int(m.group(1)) if m else 1
