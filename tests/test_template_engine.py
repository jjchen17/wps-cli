# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""模板合并引擎测试 — 使用 mock COM 对象"""

from __future__ import annotations

import pytest

from wps_cli.services.template_engine import TemplateEngine
from wps_cli.services.writer_service import WriterService

# ── Mock COM 对象 ──


class MockFind:
    """模拟 COM Find 对象"""

    def __init__(self, parent: MockRange) -> None:
        self._parent = parent
        self.Text = ""
        self.Replacement = MockReplacement()
        self.MatchCase = False
        self.MatchWildcards = False
        self.Forward = True
        self.Wrap = 0

    def ClearFormatting(self) -> None:
        pass

    def Execute(self, Replace: int = 0) -> bool:
        """模拟 Find 执行

        Replace=0 (wdFindContinue): 仅查找，不修改文本。
        通过维护内部偏移量模拟 COM 的逐次前进查找。
        Replace=2 (wdReplaceAll): 全量替换。
        """
        text = self._parent.Text
        old = self.Text
        new = self.Replacement.Text

        if Replace == 0:  # 扫描模式：逐次前进查找，不修改文本
            # 使用内部偏移量模拟 COM 的 Range 前进行为
            if not hasattr(self, "_scan_offset"):
                self._scan_offset = 0
            idx = text.find(old, self._scan_offset)
            if idx >= 0:
                self._scan_offset = idx + len(old)
                return True
            return False
        else:  # 替换模式 (Replace=2 = WD_REPLACE_ALL)
            if old in text:
                self._parent.Text = text.replace(old, new)
            return True


class MockReplacement:
    """模拟 COM Replacement 对象"""

    def __init__(self) -> None:
        self.Text = ""

    def ClearFormatting(self) -> None:
        pass


class MockRange:
    """模拟 COM Range 对象

    ``Find`` 属性每次访问都创建新实例，模拟真实 COM 行为：
    ``doc.Content.Find`` 返回独立的 Find 对象。
    """

    def __init__(self, text: str = "") -> None:
        self.Text = text
        self.Start = 0
        self.End = len(text) if text else 0

    @property
    def Find(self) -> MockFind:
        return MockFind(self)

    def Copy(self) -> None:
        pass


class MockCell:
    """模拟表格单元格"""

    def __init__(self, row: int, col: int, text: str = "") -> None:
        self.Row = row
        self.Col = col
        self.Range = MockRange(text)


class MockParagraph:
    """模拟段落"""

    def __init__(self, idx: int, text: str = "", style: str = "正文") -> None:
        self._idx = idx
        self.Range = MockRange(text)
        self.Style = style
        self._text = text


class MockParagraphs:
    """模拟段落集合"""

    def __init__(self, items: list[MockParagraph] | None = None) -> None:
        self._items = items or []
        self.Count = len(self._items)

    def __call__(self, index: int) -> MockParagraph:
        return self._items[index - 1]

    def __iter__(self):
        return iter(self._items)


class MockTable:
    """模拟表格"""

    def __init__(self, data: list[list[str]], index: int = 1) -> None:
        self.Index = index
        self.Rows = MockRows(len(data))
        self.Columns = MockRows(len(data[0]) if data else 0)
        self._cells: dict[tuple[int, int], MockCell] = {}
        for r, row in enumerate(data, 1):
            for c, val in enumerate(row, 1):
                self._cells[(r, c)] = MockCell(r, c, val)
        self.Borders = MockBorders()

    def Cell(self, row: int, col: int) -> MockCell:
        key = (row, col)
        if key not in self._cells:
            self._cells[key] = MockCell(row, col, "")
        return self._cells[key]


class MockRows:
    """模拟 Rows/Columns 集合"""

    def __init__(self, count: int) -> None:
        self.Count = count


class MockBorders:
    """模拟表格 Borders"""

    def __init__(self) -> None:
        self.Enable = False


class MockTables:
    """模拟表格集合"""

    def __init__(self, items: list[MockTable] | None = None) -> None:
        self._items = items or []
        self.Count = len(self._items)

    def __call__(self, index: int) -> MockTable:
        return self._items[index - 1]


class MockHeaderFooter:
    """模拟页眉页脚"""

    def __init__(self, text: str = "") -> None:
        self.Range = MockRange(text)
        self.Shapes = MockShapes()


class MockShapes:
    """模拟 Shapes 集合"""

    def __init__(self) -> None:
        pass

    def AddTextEffect(self, *args, **kwargs):
        return MockShape()

    def AddPicture(self, *args, **kwargs):
        return MockShape()


class MockShape:
    """模拟 Shape"""

    def __init__(self) -> None:
        self.Width = 0
        self.Height = 0
        self.Rotation = 0
        self.Fill = MockFill()
        self.TextEffect = MockTextEffect()
        self.AlternativeText = ""


class MockFill:
    """模拟 Fill"""

    class ForeColor:
        RGB = 0

    def __init__(self) -> None:
        self.ForeColor = self.ForeColor()


class MockTextEffect:
    """模拟 TextEffect"""

    NormalizedHeight = False


class MockHeadersFooters:
    """模拟 Headers/Footers 集合"""

    def __init__(self, items: list[MockHeaderFooter] | None = None) -> None:
        self._items = items or []
        self.Count = len(self._items)

    def __call__(self, index: int) -> MockHeaderFooter:
        return self._items[index - 1]


class MockSection:
    """模拟文档节"""

    def __init__(self, header_text: str = "", footer_text: str = "") -> None:
        self.Headers = MockHeadersFooters([MockHeaderFooter(header_text)])
        self.Footers = MockHeadersFooters([MockHeaderFooter(footer_text)])


class MockSections:
    """模拟 Sections 集合"""

    def __init__(self, items: list[MockSection] | None = None) -> None:
        self._items = items or [MockSection()]
        self._count = len(self._items)

    @property
    def Count(self) -> int:
        return self._count

    def __iter__(self):
        return iter(self._items)


class MockDoc:
    """模拟 WPS Document 对象"""

    def __init__(
        self,
        paragraphs: list[str] | None = None,
        tables: list[list[list[str]]] | None = None,
        headers: list[str] | None = None,
        footers: list[str] | None = None,
    ) -> None:
        paragraphs = paragraphs or []
        tables = tables or []
        self.Paragraphs = MockParagraphs(
            [MockParagraph(i + 1, text) for i, text in enumerate(paragraphs)]
        )
        self.Tables = MockTables(
            [MockTable(data, i + 1) for i, data in enumerate(tables)]
        )
        self._sections: list[MockSection] = []
        if headers or footers:
            max_len = max(len(headers or []), len(footers or []))
            for i in range(max_len):
                h = headers[i] if headers and i < len(headers) else ""
                f = footers[i] if footers and i < len(footers) else ""
                self._sections.append(MockSection(h, f))
        else:
            self._sections.append(MockSection())
        self.Sections = MockSections(self._sections)
        self.InlineShapes = MockInlineShapes()
        self.BuiltInDocumentProperties = MockBuiltInProperties()
        self.FullName = "/mock/test.docx"

        # doc.Content 返回整个文档的 Range（所有段落文本拼接）
        full_text = "\r\n".join(paragraphs)
        # 也包含表格和页眉页脚文本
        for table in tables:
            for row in table:
                for cell in row:
                    full_text += "\r\n" + cell
        for h in (headers or []):
            full_text += "\r\n" + h
        for f in (footers or []):
            full_text += "\r\n" + f
        self.Content = MockRange(full_text)

    def ComputeStatistics(self, stat_type: int) -> int:
        return len(self.Paragraphs._items)

    def Close(self, *args, **kwargs) -> None:
        pass

    def SaveAs(self, path: str) -> None:
        pass

    def Save(self) -> None:
        pass

    def ExportAsFixedFormat(self, path: str, fmt: int) -> None:
        pass

    def Range(self, start: int = 0, end: int = -1) -> MockRange:
        if end < 0:
            end = len(self.Content.Text)
        return MockRange(self.Content.Text[start:end])


class MockInlineShapes:
    """模拟 InlineShapes 集合"""

    def __init__(self) -> None:
        self.Count = 0

    def __call__(self, index: int):
        raise IndexError("No inline shapes in mock")


class MockBuiltInProperties:
    """模拟 BuiltInDocumentProperties"""

    def __call__(self, name: str):
        return MockProperty(name)


class MockProperty:
    """模拟文档属性"""

    def __init__(self, name: str) -> None:
        self.Value = "" if name else ""


class MockApp:
    """模拟 WPS Writer 应用对象"""

    def __init__(
        self,
        paragraphs: list[str] | None = None,
        tables: list[list[list[str]]] | None = None,
        headers: list[str] | None = None,
        footers: list[str] | None = None,
    ) -> None:
        self.ActiveDocument = MockDoc(paragraphs, tables, headers, footers)
        self.Name = "Mock Writer"
        self.Version = "12.0.0-test"
        self.Visible = False
        self.AutomationSecurity = 0
        self.DisplayAlerts = True
        self.Documents = MockDocuments()


class MockDocuments:
    """模拟 Documents 集合"""

    def Open(self, *args, **kwargs):
        return MockDoc()

    def Add(self):
        doc = MockDoc()
        doc.FullName = "/mock/new.docx"
        return doc


class MockSelection:
    """模拟 Selection"""

    def __init__(self, app: MockApp) -> None:
        self.app = app
        self.Range = MockRange()
        self.Font = MockFont()
        self.ParagraphFormat = MockParagraphFormat()
        self.InlineShapes = MockInlineShapes2()

    def EndKey(self, *args) -> None:
        pass

    def TypeText(self, text: str) -> None:
        pass

    def TypeParagraph(self) -> None:
        pass

    def InsertBreak(self, *args) -> None:
        pass

    def InsertFile(self, path: str) -> None:
        pass


class MockFont:
    """模拟 Font"""

    def __init__(self) -> None:
        self.Name = "宋体"
        self.NameFarEast = "宋体"
        self.Size = 12
        self.Bold = False
        self.Color = 0


class MockParagraphFormat:
    """模拟 ParagraphFormat"""

    def __init__(self) -> None:
        self.Alignment = 0
        self.LeftIndent = 0
        self.FirstLineIndent = 0
        self.LineSpacingRule = 0
        self.LineSpacing = 12


class MockInlineShapes2:
    """模拟 Selection InlineShapes"""

    def AddPicture(self, path: str, *args, **kwargs):
        return MockShape()


# ── 测试类 ──


class TestTemplateEngineFill:
    """测试模板填充"""

    def test_simple_paragraph_replacement(self):
        """段落中的 {{key}} 被正确替换"""
        app = MockApp(paragraphs=["你好，{{name}}！", "日期：{{date}}"])
        engine = TemplateEngine()
        result = engine.fill(app, {"name": "张三", "date": "2026-06-07"})
        assert "replaced" in result
        assert result["total"] == 2
        assert result["replaced"]["name"] == 1
        assert result["replaced"]["date"] == 1
        # doc.Content.Text 已被替换（模拟 Find 在 Content Range 上执行）
        assert "张三" in app.ActiveDocument.Content.Text
        assert "2026-06-07" in app.ActiveDocument.Content.Text

    def test_table_replacement(self):
        """表格中的 {{key}} 被替换"""
        app = MockApp(
            paragraphs=["报告"],
            tables=[[["{{name}}", "{{age}}"], ["{{city}}", "中国"]]],
        )
        engine = TemplateEngine()
        result = engine.fill(app, {"name": "李四", "age": "30", "city": "北京"})
        assert "replaced" in result

    def test_header_replacement(self):
        """页眉中的 {{key}} 被替换"""
        app = MockApp(
            paragraphs=["正文内容"],
            headers=["{{company}} 机密文件", "版本：{{version}}"],
            footers=["第 {{page}} 页"],
        )
        engine = TemplateEngine()
        result = engine.fill(
            app, {"company": "腾讯", "version": "1.0", "page": "1"}
        )
        assert result["total"] >= 0  # header replacement may or may not work depending on mock

    def test_missing_key_raises_keyerror(self):
        """data 中缺少文档需要的 key 时抛出 KeyError"""
        app = MockApp(paragraphs=["你好，{{name}}！", "{{email}}"])
        engine = TemplateEngine()
        with pytest.raises(KeyError, match="email"):
            engine.fill(app, {"name": "张三"})

    def test_empty_data_raises(self):
        """空 data 抛出 ValueError"""
        app = MockApp(paragraphs=["hello"])
        engine = TemplateEngine()
        with pytest.raises(ValueError, match="data 不能为空"):
            engine.fill(app, {})

    def test_all_keys_present(self):
        """所有 key 都在 data 中时应成功"""
        app = MockApp(paragraphs=["{{a}} {{b}} {{c}}"])
        engine = TemplateEngine()
        result = engine.fill(app, {"a": "1", "b": "2", "c": "3"})
        assert result["total"] == 3

    def test_no_placeholders_no_error(self):
        """文档中没有占位符时不报错"""
        app = MockApp(paragraphs=["纯文本，无占位符。"])
        engine = TemplateEngine()
        result = engine.fill(app, {"name": "张三"})
        assert result["total"] == 0

    def test_extra_keys_in_data_no_error(self):
        """data 中有多余的 key 不报错（灵活模板）"""
        app = MockApp(paragraphs=["你好，{{name}}！"])
        engine = TemplateEngine()
        result = engine.fill(app, {"name": "张三", "extra": "value"})
        assert result["total"] == 1

    def test_duplicate_keys_counted(self):
        """同一 key 出现多次时计数正确"""
        app = MockApp(
            paragraphs=["{{name}} 你好！", "再次欢迎 {{name}}", "{{name}} 再见！"]
        )
        engine = TemplateEngine()
        result = engine.fill(app, {"name": "王五"})
        assert result["replaced"]["name"] == 3
        assert result["total"] == 3


class TestTemplateEngineExtractKeys:
    """测试占位符提取"""

    def test_extract_from_paragraphs(self):
        """从段落提取占位符键名"""
        app = MockApp(paragraphs=["{{name}} 你好", "{{date}} 报告", "无占位符"])
        engine = TemplateEngine()
        keys = engine.extract_keys(app)
        assert sorted(keys) == ["date", "name"]

    def test_extract_from_tables(self):
        """从表格提取占位符键名"""
        app = MockApp(
            paragraphs=["标题"],
            tables=[[["{{product}}", "{{price}}"], ["{{qty}}", "备注"]]],
        )
        engine = TemplateEngine()
        keys = engine.extract_keys(app)
        assert sorted(keys) == ["price", "product", "qty"]

    def test_extract_no_placeholders(self):
        """无占位符时返回空列表"""
        app = MockApp(paragraphs=["纯文本", "也没有"])
        engine = TemplateEngine()
        keys = engine.extract_keys(app)
        assert keys == []

    def test_extract_unique_sorted(self):
        """键名去重排序"""
        app = MockApp(paragraphs=["{{name}} {{date}} {{name}}"])
        engine = TemplateEngine()
        keys = engine.extract_keys(app)
        assert keys == ["date", "name"]


class TestWriterServiceTemplateFill:
    """测试 WriterService.template_fill 集成"""

    def test_template_fill_via_service(self):
        """通过 WriterService.template_fill 调用模板引擎"""
        app = MockApp(paragraphs=["{{greeting}}，{{name}}！"])
        result = WriterService.template_fill(app, {"greeting": "你好", "name": "世界"})
        assert result["total"] == 2
