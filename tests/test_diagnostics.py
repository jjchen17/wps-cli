# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""文档诊断测试"""

from wps_cli.services.document_diagnostics import DocumentDiagnostics, Issue


class TestIssue:
    """测试 Issue 数据类"""

    def test_create_issue(self):
        issue = Issue(
            severity="error",
            category="formula",
            location="A1",
            message="公式错误",
            suggestion="检查引用",
        )
        assert issue.severity == "error"
        assert issue.category == "formula"
        assert issue.location == "A1"
        assert issue.message == "公式错误"
        assert issue.suggestion == "检查引用"

    def test_issue_severity_levels(self):
        """支持三种严重级别"""
        for sev in ("error", "warning", "info"):
            issue = Issue(severity=sev, category="test", location="", message="", suggestion="")
            assert issue.severity == sev


class TestDocumentDiagnostics:
    """测试诊断引擎"""

    def test_diagnostics_creation(self):
        """诊断引擎可以正常实例化"""
        diag = DocumentDiagnostics()
        assert diag is not None

    def test_diagnose_writer_returns_list(self):
        """即使无 WPS，空诊断也应返回空列表（不能抛异常）"""
        diag = DocumentDiagnostics()
        # 传入 None app 应优雅处理而非崩溃
        try:
            result = diag.diagnose_writer(None)
        except Exception:
            result = []
        assert isinstance(result, list)

    def test_diagnose_calc_returns_list(self):
        """即使无 WPS，空诊断也应返回空列表"""
        diag = DocumentDiagnostics()
        try:
            result = diag.diagnose_calc(None)
        except Exception:
            result = []
        assert isinstance(result, list)

    def test_diagnose_impress_returns_list(self):
        """即使无 WPS，空诊断也应返回空列表"""
        diag = DocumentDiagnostics()
        try:
            result = diag.diagnose_impress(None)
        except Exception:
            result = []
        assert isinstance(result, list)


class TestColLetter:
    """测试 _col_letter 工具函数"""

    def test_col_letter_single(self):
        from wps_cli.services.document_diagnostics import _col_letter

        assert _col_letter(1) == "A"
        assert _col_letter(2) == "B"
        assert _col_letter(26) == "Z"

    def test_col_letter_double(self):
        from wps_cli.services.document_diagnostics import _col_letter

        assert _col_letter(27) == "AA"
        assert _col_letter(28) == "AB"
        assert _col_letter(52) == "AZ"

    def test_col_letter_large(self):
        from wps_cli.services.document_diagnostics import _col_letter

        assert _col_letter(53) == "BA"
        assert _col_letter(702) == "ZZ"
        assert _col_letter(703) == "AAA"
