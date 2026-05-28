"""ExportService 单元测试"""

import pytest
from pathlib import Path


class TestExportService:
    def test_detect_app_type(self):
        from wps_cli.services.export_service import ExportService
        from wps_cli.services.session_manager import SessionManager
        from tests.conftest import MockComBackend

        svc = ExportService(manager=SessionManager(backend=MockComBackend()))
        assert svc.detect_app_type(Path("a.docx")) == "writer"
        assert svc.detect_app_type(Path("b.xlsx")) == "calc"
        assert svc.detect_app_type(Path("c.pptx")) == "impress"
        assert svc.detect_app_type(Path("d.doc")) == "writer"
        assert svc.detect_app_type(Path("e.csv")) == "calc"
        with pytest.raises(ValueError):
            svc.detect_app_type(Path("f.xyz"))

    def test_format_map_separated(self):
        from wps_cli.services.export_service import ExportService
        from wps_cli.services.session_manager import SessionManager
        from tests.conftest import MockComBackend

        svc = ExportService(manager=SessionManager(backend=MockComBackend()))
        fm = svc._get_format_map("writer")
        assert "pdf" in fm
        assert "docx" in fm
        fm = svc._get_format_map("calc")
        assert "xlsx" in fm
        assert "csv" in fm
        fm = svc._get_format_map("impress")
        assert "pptx" in fm
        assert "pdf" in fm


class TestPdfService:
    def test_parse_pages(self):
        from wps_cli.services.pdf_service import PdfService
        from wps_cli.services.session_manager import SessionManager
        from tests.conftest import MockComBackend

        svc = PdfService(manager=SessionManager(backend=MockComBackend()))
        assert svc._parse_pages("1-3") == [1, 2, 3]
        assert svc._parse_pages("1,3,5") == [1, 3, 5]
        assert svc._parse_pages("1-3,5,7-9") == [1, 2, 3, 5, 7, 8, 9]
        assert svc._parse_pages("3-3") == [3]
        with pytest.raises(ValueError, match="起始页不能大于结束页"):
            svc._parse_pages("5-3")

    def test_merge_empty_inputs(self):
        from wps_cli.services.pdf_service import PdfService
        from wps_cli.services.session_manager import SessionManager
        from tests.conftest import MockComBackend

        svc = PdfService(manager=SessionManager(backend=MockComBackend()))
        with pytest.raises(ValueError, match="至少需要一个输入文件"):
            svc.merge([], Path("out.pdf"))


class TestPathUtils:
    def test_validate_path(self):
        from wps_cli.utils.path_utils import validate_path

        p = validate_path(".")
        assert p.is_absolute()

    def test_sanitize_filename(self):
        from wps_cli.utils.path_utils import sanitize_filename

        assert sanitize_filename("hello") == "hello"
        assert sanitize_filename("file<name>.txt") == "file_name_.txt"
        with pytest.raises(ValueError):
            sanitize_filename("")
