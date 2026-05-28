"""ExportService 单元测试"""

from pathlib import Path

import pytest

from tests.conftest import MockComBackend
from wps_cli.exceptions import FormatError, ValidationError
from wps_cli.services.export_service import ExportService
from wps_cli.services.session_manager import SessionManager


def _make_svc() -> ExportService:
    return ExportService(manager=SessionManager(backend=MockComBackend()))


class TestExportService:
    def test_detect_app_type(self):
        svc = _make_svc()
        assert svc.detect_app_type(Path("a.docx")) == "writer"
        assert svc.detect_app_type(Path("b.xlsx")) == "calc"
        assert svc.detect_app_type(Path("c.pptx")) == "impress"
        assert svc.detect_app_type(Path("d.doc")) == "writer"
        assert svc.detect_app_type(Path("e.csv")) == "calc"
        with pytest.raises(FormatError):
            svc.detect_app_type(Path("f.xyz"))

    def test_format_map_separated(self):
        svc = _make_svc()
        fm = svc._get_format_map("writer")
        assert "pdf" in fm
        assert "docx" in fm
        fm = svc._get_format_map("calc")
        assert "xlsx" in fm
        assert "csv" in fm
        fm = svc._get_format_map("impress")
        assert "pptx" in fm
        assert "pdf" in fm

    def test_unknown_target_format_raises_format_error(self, tmp_path):
        svc = _make_svc()
        f = tmp_path / "a.docx"
        f.write_text("x", encoding="utf-8")
        with pytest.raises(FormatError):
            svc.convert(f, "unsupported_format")

    def test_batch_rejects_absolute_glob(self, tmp_path):
        svc = _make_svc()
        with pytest.raises(ValidationError, match="绝对路径"):
            svc.batch_convert("C:\\*.docx", "pdf", tmp_path)

    def test_batch_rejects_unc_glob(self, tmp_path):
        svc = _make_svc()
        with pytest.raises(ValidationError, match="UNC"):
            svc.batch_convert("\\\\server\\share\\*", "pdf", tmp_path)

    def test_batch_rejects_empty_glob(self, tmp_path):
        svc = _make_svc()
        with pytest.raises(ValidationError, match="不能为空"):
            svc.batch_convert("", "pdf", tmp_path)
