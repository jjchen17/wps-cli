"""WriterService 表单域与内容控件单元测试

设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""

from __future__ import annotations

from tests.conftest import MockComBackend
from wps_cli.services.session_manager import SessionManager
from wps_cli.services.writer_service import WriterService


def _make_svc() -> WriterService:
    return WriterService(manager=SessionManager(backend=MockComBackend()))


# ── 模拟 COM 对象 ──


class MockFormField:
    """模拟旧式 FormField"""

    def __init__(self, index: int, ff_type: int, name: str, result: str):
        self._index = index
        self._type = ff_type
        self._name = name
        self._result = result
        self.StatusText = ""  # help text via StatusText

    @property
    def Type(self):
        return self._type

    @property
    def Name(self):
        return self._name

    @property
    def Result(self):
        return self._result

    @Result.setter
    def Result(self, value):
        self._result = value


class MockFormFields:
    """模拟 FormFields 集合 — 支持 .Count 和 .__call__(idx)"""

    def __init__(self, fields: list[MockFormField]):
        self._fields = fields

    @property
    def Count(self):
        return len(self._fields)

    def __call__(self, index):
        return self._fields[index - 1]


class MockContentControl:
    """模拟 ContentControl"""

    def __init__(
        self,
        index: int,
        cc_type: int,
        title: str,
        tag: str,
        text: str,
        lock_contents: bool = False,
        lock_cc: bool = False,
    ):
        self._index = index
        self.Type = cc_type
        self.Title = title
        self.Tag = tag
        self._text = text
        self.LockContents = lock_contents
        self.LockContentControl = lock_cc

    @property
    def Range(self):
        # 每次访问返回相同的 _Range 实例以保证 set 持久
        if not hasattr(self, "_range_obj"):
            self._range_obj = self._Range(self)
        return self._range_obj

    class _Range:
        def __init__(self, parent):
            self._parent = parent

        @property
        def Text(self):
            return self._parent._text

        @Text.setter
        def Text(self, value):
            self._parent._text = value


class MockContentControls:
    """模拟 ContentControls 集合 — 支持 .Count 和 .__call__(idx)"""

    def __init__(self, controls: list[MockContentControl]):
        self._controls = controls

    @property
    def Count(self):
        return len(self._controls)

    def __call__(self, index):
        return self._controls[index - 1]


class MockDocument:
    """模拟 Word 文档"""

    def __init__(
        self,
        formfields: MockFormFields | None = None,
        contentcontrols: MockContentControls | None = None,
    ):
        self.FormFields = formfields or MockFormFields([])
        self.ContentControls = contentcontrols or MockContentControls([])


class MockApp:
    """模拟 WPS Word 应用 — 提供 ActiveDocument"""

    def __init__(self, doc: MockDocument):
        self.ActiveDocument = doc


# ── FormField 测试 ──


class TestFormFieldList:
    """formfield_list 测试"""

    def test_empty_document(self):
        svc = _make_svc()
        doc = MockDocument()
        app = MockApp(doc)
        result = svc.formfield_list(app)
        assert result == []

    def test_list_text_field(self):
        svc = _make_svc()
        fields = MockFormFields([MockFormField(1, 70, "NameField", "张三")])
        doc = MockDocument(formfields=fields)
        app = MockApp(doc)
        result = svc.formfield_list(app)
        assert len(result) == 1
        assert result[0]["index"] == 1
        assert result[0]["name"] == "NameField"
        assert result[0]["type"] == "text"
        assert result[0]["result"] == "张三"

    def test_list_checkbox_field(self):
        svc = _make_svc()
        fields = MockFormFields([MockFormField(1, 71, "AgreeCheck", "1")])
        doc = MockDocument(formfields=fields)
        app = MockApp(doc)
        result = svc.formfield_list(app)
        assert len(result) == 1
        assert result[0]["type"] == "checkbox"
        assert result[0]["result"] == "1"

    def test_list_dropdown_field(self):
        svc = _make_svc()
        fields = MockFormFields([MockFormField(1, 83, "CitySelect", "北京")])
        doc = MockDocument(formfields=fields)
        app = MockApp(doc)
        result = svc.formfield_list(app)
        assert len(result) == 1
        assert result[0]["type"] == "dropdown"
        assert result[0]["result"] == "北京"

    def test_list_unknown_type(self):
        svc = _make_svc()
        fields = MockFormFields([MockFormField(1, 999, "Unknown", "")])
        doc = MockDocument(formfields=fields)
        app = MockApp(doc)
        result = svc.formfield_list(app)
        assert len(result) == 1
        assert result[0]["type"] == "unknown(999)"

    def test_list_multiple_fields(self):
        svc = _make_svc()
        fields = MockFormFields(
            [
                MockFormField(1, 70, "Name", "张三"),
                MockFormField(2, 71, "OK", ""),
                MockFormField(3, 83, "City", "上海"),
            ]
        )
        doc = MockDocument(formfields=fields)
        app = MockApp(doc)
        result = svc.formfield_list(app)
        assert len(result) == 3
        assert result[0]["type"] == "text"
        assert result[1]["type"] == "checkbox"
        assert result[2]["type"] == "dropdown"

    def test_formfield_exception_handled(self):
        svc = _make_svc()

        class BadDoc:
            @property
            def FormFields(self):
                raise RuntimeError("COM error")

        app = MockApp(BadDoc())
        result = svc.formfield_list(app)
        assert result == []


class TestFormFieldGet:
    """formfield_get 测试"""

    def test_get_existing_field(self):
        svc = _make_svc()
        fields = MockFormFields([MockFormField(1, 70, "NameField", "李四")])
        doc = MockDocument(formfields=fields)
        app = MockApp(doc)
        result = svc.formfield_get(app, 1)
        assert result["index"] == 1
        assert result["name"] == "NameField"
        assert result["type"] == "text"
        assert result["result"] == "李四"
        assert "help_text" in result


class TestFormFieldSet:
    """formfield_set 测试"""

    def test_set_text_field_value(self):
        svc = _make_svc()
        field = MockFormField(1, 70, "NameField", "")
        fields = MockFormFields([field])
        doc = MockDocument(formfields=fields)
        app = MockApp(doc)

        svc.formfield_set(app, 1, "王五")
        assert field.Result == "王五"

    def test_set_checkbox_value(self):
        svc = _make_svc()
        field = MockFormField(1, 71, "AgreeCheck", "")
        fields = MockFormFields([field])
        doc = MockDocument(formfields=fields)
        app = MockApp(doc)

        svc.formfield_set(app, 1, "1")
        assert field.Result == "1"


# ── ContentControl 测试 ──


class TestContentControlList:
    """content_control_list 测试"""

    def test_empty_document(self):
        svc = _make_svc()
        doc = MockDocument()
        app = MockApp(doc)
        result = svc.content_control_list(app)
        assert result == []

    def test_list_rich_text_control(self):
        svc = _make_svc()
        controls = MockContentControls(
            [MockContentControl(1, 0, "Title1", "tag1", "富文本内容")]
        )
        doc = MockDocument(contentcontrols=controls)
        app = MockApp(doc)
        result = svc.content_control_list(app)
        assert len(result) == 1
        assert result[0]["index"] == 1
        assert result[0]["title"] == "Title1"
        assert result[0]["tag"] == "tag1"
        assert result[0]["type"] == "rich_text"
        assert result[0]["text"] == "富文本内容"

    def test_list_plain_text_control(self):
        svc = _make_svc()
        controls = MockContentControls(
            [MockContentControl(1, 1, "Input1", "plain", "纯文本")]
        )
        doc = MockDocument(contentcontrols=controls)
        app = MockApp(doc)
        result = svc.content_control_list(app)
        assert result[0]["type"] == "plain_text"

    def test_list_date_control(self):
        svc = _make_svc()
        controls = MockContentControls(
            [MockContentControl(1, 6, "Date1", "date", "2026-06-08")]
        )
        doc = MockDocument(contentcontrols=controls)
        app = MockApp(doc)
        result = svc.content_control_list(app)
        assert result[0]["type"] == "date"

    def test_list_checkbox_control(self):
        svc = _make_svc()
        controls = MockContentControls(
            [MockContentControl(1, 7, "Check1", "check", "")]
        )
        doc = MockDocument(contentcontrols=controls)
        app = MockApp(doc)
        result = svc.content_control_list(app)
        assert result[0]["type"] == "checkbox"

    def test_list_dropdown_control(self):
        svc = _make_svc()
        controls = MockContentControls(
            [MockContentControl(1, 4, "Dropdown1", "dd", "选项1")]
        )
        doc = MockDocument(contentcontrols=controls)
        app = MockApp(doc)
        result = svc.content_control_list(app)
        assert result[0]["type"] == "dropdown"

    def test_lock_status(self):
        svc = _make_svc()
        controls = MockContentControls(
            [
                MockContentControl(1, 1, "Locked", "t", "内容", lock_contents=True),
                MockContentControl(2, 1, "NoDelete", "t2", "内容2", lock_cc=True),
                MockContentControl(3, 1, "Free", "t3", "内容3"),
            ]
        )
        doc = MockDocument(contentcontrols=controls)
        app = MockApp(doc)
        result = svc.content_control_list(app)
        assert result[0]["lock"] == "content_locked"
        assert result[1]["lock"] == "cannot_delete"
        assert result[2]["lock"] == ""

    def test_contentcontrol_exception_handled(self):
        svc = _make_svc()

        class BadDoc:
            @property
            def ContentControls(self):
                raise RuntimeError("COM error")

        app = MockApp(BadDoc())
        result = svc.content_control_list(app)
        assert result == []


class TestContentControlSet:
    """content_control_set 测试"""

    def test_set_rich_text(self):
        svc = _make_svc()
        cc = MockContentControl(1, 0, "Title", "tag", "旧内容")
        controls = MockContentControls([cc])
        doc = MockDocument(contentcontrols=controls)
        app = MockApp(doc)

        svc.content_control_set(app, 1, "新内容")
        assert cc.Range.Text == "新内容"

    def test_set_plain_text(self):
        svc = _make_svc()
        cc = MockContentControl(1, 1, "Input", "tag", "")
        controls = MockContentControls([cc])
        doc = MockDocument(contentcontrols=controls)
        app = MockApp(doc)

        svc.content_control_set(app, 1, "用户输入")
        assert cc.Range.Text == "用户输入"
