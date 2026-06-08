"""Dump 往返序列化测试

设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""

from __future__ import annotations

import tempfile
from pathlib import Path

from tests.conftest import MockApp, MockComBackend
from wps_cli.services.dump_service import DumpService
from wps_cli.services.session_manager import SessionManager


def _make_svc() -> DumpService:
    return DumpService(manager=SessionManager(backend=MockComBackend()))


# ── Writer Mock 对象 ──────────────────────────────────────────────


class _FakePageSetup:
    TopMargin = 72.0
    BottomMargin = 72.0
    LeftMargin = 90.0
    RightMargin = 90.0
    PageWidth = 612.0
    PageHeight = 792.0
    Orientation = "0"


class _FakeFont:
    Name = "宋体"
    Size = 12.0
    Bold = False
    Italic = False
    Color = "0"


class _FakeParaRange:
    Text = "测试段落文本\r\x07"
    Font = _FakeFont()


class _FakeParagraph:
    def __init__(self, index: int):
        self.Range = _FakeParaRange()
        self.Alignment = "1"  # wdAlignParagraphLeft
        self.Style = type("Style", (), {"NameLocal": "正文"})()
        self._index = index


class _FakeParagraphs:
    def __init__(self, count: int = 3):
        self._paras = [_FakeParagraph(i) for i in range(1, count + 1)]

    @property
    def Count(self):
        return len(self._paras)

    def __call__(self, index):
        return self._paras[index - 1]


class _FakeCell:
    def __init__(self, text: str):
        self.Range = type("Range", (), {"Text": text + "\r\x07"})()


class _FakeTable:
    def __init__(self, index: int):
        self._index = index
        self.Rows = type("Rows", (), {"Count": 2})()
        self.Columns = type("Columns", (), {"Count": 2})()

    def Cell(self, row, col):
        return _FakeCell(f"R{row}C{col}")


class _FakeTables:
    def __init__(self, count: int = 1):
        self._tables = [_FakeTable(i) for i in range(1, count + 1)]

    @property
    def Count(self):
        return len(self._tables)

    def __call__(self, index):
        return self._tables[index - 1]


class _FakeInlineShape:
    def __init__(self, index: int):
        self.Width = 100.0
        self.Height = 100.0
        self.AlternativeText = f"图片{index}"
        self.LinkFormat = None


class _FakeInlineShapes:
    def __init__(self, count: int = 1):
        self._shapes = [_FakeInlineShape(i) for i in range(1, count + 1)]

    @property
    def Count(self):
        return len(self._shapes)

    def __call__(self, index):
        return self._shapes[index - 1]


class _FakeHeaderFooter:
    def __init__(self, exists: bool = True):
        self.Exists = exists
        self.Range = type(
            "Range",
            (),
            {"Paragraphs": _FakeParagraphs(1), "Text": "页眉文本\r\x07"},
        )()


class _FakeHeadersFooters:
    def __init__(self, exists: bool = True):
        self._hf = _FakeHeaderFooter(exists)

    def __call__(self, index):
        return self._hf


class _FakeSection:
    def __init__(self, index: int):
        self._index = index
        self.Headers = _FakeHeadersFooters(True)
        self.Footers = _FakeHeadersFooters(True)


class _FakeSections:
    def __init__(self, count: int = 1):
        self._sections = [_FakeSection(i) for i in range(1, count + 1)]

    @property
    def Count(self):
        return len(self._sections)

    def __call__(self, index):
        return self._sections[index - 1]


class _FakeWriterDoc:
    def __init__(self):
        self.PageSetup = _FakePageSetup()
        self.Paragraphs = _FakeParagraphs(3)
        self.Tables = _FakeTables(1)
        self.InlineShapes = _FakeInlineShapes(1)
        self.Sections = _FakeSections(1)

    def Close(self):
        pass


# ── Impress Mock 对象 ──────────────────────────────────────────


class _FakeShape:
    def __init__(self, index: int):
        self.Left = 100.0
        self.Top = 100.0
        self.Width = 300.0
        self.Height = 200.0
        self.Type = 17  # msoTextBox
        self.HasTextFrame = True
        self.TextFrame = type(
            "TextFrame",
            (),
            {"TextRange": type("TextRange", (), {"Text": f"文本框{index}"})()},
        )()
        self.AlternativeText = ""


class _FakeShapes:
    def __init__(self, count: int = 2):
        self._shapes = [_FakeShape(i) for i in range(1, count + 1)]

    @property
    def Count(self):
        return len(self._shapes)

    def __call__(self, index):
        return self._shapes[index - 1]


class _FakeSlide:
    def __init__(self, index: int):
        self._index = index
        self.Shapes = _FakeShapes(2)
        self.Layout = type("Layout", (), {"Name": "标题和内容"})()


class _FakeSlides:
    def __init__(self, count: int = 2):
        self._slides = [_FakeSlide(i) for i in range(1, count + 1)]

    @property
    def Count(self):
        return len(self._slides)

    def __call__(self, index):
        return self._slides[index - 1]


class _FakeImpressPres:
    def __init__(self):
        self.PageSetup = type(
            "PageSetup",
            (),
            {"SlideWidth": 960.0, "SlideHeight": 540.0},
        )()
        self.Slides = _FakeSlides(2)

    def Close(self):
        pass


# ── 测试 ────────────────────────────────────────────────────────


class TestDumpWriter:
    """D-1：Writer 文档序列化"""

    def test_dump_writer_generates_commands(self):
        svc = _make_svc()
        app = MockApp("writer")
        app.ActiveDocument = _FakeWriterDoc()

        commands = svc.dump_writer(app)
        assert len(commands) > 0
        # 应包含 page_setup + paragraphs + table + inline_shape + header + footer
        actions = {cmd["action"] for cmd in commands}
        assert "writer.page_setup" in actions
        assert "writer.insert_text" in actions
        assert "writer.insert_table" in actions
        assert "writer.insert_image" in actions

    def test_dump_writer_page_setup_has_margins(self):
        svc = _make_svc()
        app = MockApp("writer")
        app.ActiveDocument = _FakeWriterDoc()

        commands = svc.dump_writer(app)
        ps_cmd = next(c for c in commands if c["action"] == "writer.page_setup")
        params = ps_cmd["params"]
        assert params["top_margin"] == 72.0
        assert params["page_width"] == 612.0

    def test_dump_writer_paragraph_has_text(self):
        svc = _make_svc()
        app = MockApp("writer")
        app.ActiveDocument = _FakeWriterDoc()

        commands = svc.dump_writer(app)
        text_cmds = [c for c in commands if c["action"] == "writer.insert_text"]
        assert len(text_cmds) >= 1
        first_text = text_cmds[0]["params"]["text"]
        assert "测试段落文本" in first_text


class TestDumpImpress:
    """D-2：PPT 文档序列化"""

    def test_dump_impress_generates_commands(self):
        svc = _make_svc()
        app = MockApp("impress")
        app.ActivePresentation = _FakeImpressPres()

        commands = svc.dump_impress(app)
        assert len(commands) > 0
        actions = {cmd["action"] for cmd in commands}
        assert "impress.slide_setup" in actions
        assert "impress.set_layout" in actions
        assert "impress.set_text" in actions

    def test_dump_impress_slide_setup_dimensions(self):
        svc = _make_svc()
        app = MockApp("impress")
        app.ActivePresentation = _FakeImpressPres()

        commands = svc.dump_impress(app)
        ss_cmd = next(c for c in commands if c["action"] == "impress.slide_setup")
        assert ss_cmd["params"]["width"] == 960.0
        assert ss_cmd["params"]["height"] == 540.0


class TestDumpFileIO:
    """D-3：JSON 文件读写"""

    def test_to_file_and_from_file_roundtrip(self):
        svc = _make_svc()
        commands = [
            {"action": "writer.page_setup", "params": {"page_width": 612}},
            {"action": "writer.insert_text", "params": {"text": "Hello"}},
        ]
        with tempfile.TemporaryDirectory() as tmp:
            out_path = Path(tmp) / "test.json"
            result_path = svc.to_file(commands, out_path)
            assert result_path.exists()

            loaded = svc.from_file(out_path)
            assert loaded == commands

    def test_to_file_creates_parent_dir(self):
        svc = _make_svc()
        commands: list[dict] = []
        with tempfile.TemporaryDirectory() as tmp:
            out_path = Path(tmp) / "subdir" / "output.json"
            svc.to_file(commands, out_path)
            assert out_path.exists()


class TestDumpReplayWriter:
    """D-4：Writer batch 回放"""

    def test_replay_page_setup(self):
        svc = _make_svc()
        app = MockApp("writer")
        app.ActiveDocument = _FakeWriterDoc()

        commands = [
            {
                "action": "writer.page_setup",
                "params": {
                    "top_margin": 50.0,
                    "page_width": 800.0,
                },
            },
        ]
        executed = svc.replay_writer(app, commands)
        assert executed == 1
        # 验证页面设置已应用
        doc = app.ActiveDocument
        assert doc.PageSetup.TopMargin == 50.0
        assert doc.PageSetup.PageWidth == 800.0

    def test_replay_skips_unknown_action(self):
        svc = _make_svc()
        app = MockApp("writer")
        app.ActiveDocument = _FakeWriterDoc()

        commands = [{"action": "unknown.action", "params": {}}]
        executed = svc.replay_writer(app, commands)
        assert executed == 0

    def test_replay_handles_empty_commands(self):
        svc = _make_svc()
        app = MockApp("writer")
        app.ActiveDocument = _FakeWriterDoc()

        executed = svc.replay_writer(app, [])
        assert executed == 0


class TestDumpReplayImpress:
    """D-5：Impress batch 回放"""

    def test_replay_slide_setup(self):
        svc = _make_svc()
        app = MockApp("impress")
        app.ActivePresentation = _FakeImpressPres()

        commands = [
            {
                "action": "impress.slide_setup",
                "params": {"width": 1024.0, "height": 768.0},
            },
        ]
        executed = svc.replay_impress(app, commands)
        assert executed == 1


class TestEdgeCases:
    """D-6：边界情况"""

    def test_dump_empty_document_skeleton(self):
        """空骨架文档也能正常序列化"""
        svc = _make_svc()
        app = MockApp("writer")

        class _EmptyDoc:
            PageSetup = _FakePageSetup()
            Paragraphs = _FakeParagraphs(0)
            Tables = _FakeTables(0)
            InlineShapes = _FakeInlineShapes(0)
            Sections = _FakeSections(0)

        app.ActiveDocument = _EmptyDoc()
        commands = svc.dump_writer(app)
        # 至少应有 page_setup
        assert any(c["action"] == "writer.page_setup" for c in commands)

    def test_safe_str_handles_none(self):
        from wps_cli.services.dump_service import _safe_str

        assert _safe_str(None) == ""

    def test_safe_str_handles_string(self):
        from wps_cli.services.dump_service import _safe_str

        assert _safe_str("hello") == "hello"
