"""PdfService 单元测试 — 重点覆盖 _parse_pages 的边界条件"""

import pytest

from tests.conftest import MockComBackend
from wps_cli.consts import MAX_PDF_PAGE_NUMBER, MAX_PDF_PAGE_RANGE_SIZE
from wps_cli.exceptions import ValidationError
from wps_cli.services.pdf_service import PdfService
from wps_cli.services.session_manager import SessionManager


def _make_svc() -> PdfService:
    return PdfService(manager=SessionManager(backend=MockComBackend()))


class TestParsePages:
    def test_single_range(self):
        assert PdfService._parse_pages("1-3") == [1, 2, 3]

    def test_comma_list(self):
        assert PdfService._parse_pages("1,3,5") == [1, 3, 5]

    def test_mixed(self):
        assert PdfService._parse_pages("1-3,5,7-9") == [1, 2, 3, 5, 7, 8, 9]

    def test_single_value_range(self):
        assert PdfService._parse_pages("3-3") == [3]

    def test_deduplicate_and_sort(self):
        assert PdfService._parse_pages("3,1,2,1") == [1, 2, 3]

    def test_inverted_range_rejected(self):
        with pytest.raises(ValidationError, match="起始页不能大于结束页"):
            PdfService._parse_pages("5-3")

    def test_zero_page_rejected(self):
        with pytest.raises(ValidationError):
            PdfService._parse_pages("0-3")

    def test_negative_page_rejected(self):
        with pytest.raises(ValidationError):
            PdfService._parse_pages("-1")

    def test_overflow_upper_bound(self):
        big = MAX_PDF_PAGE_NUMBER + 1
        with pytest.raises(ValidationError, match="超出"):
            PdfService._parse_pages(f"1-{big}")

    def test_memory_bomb_blocked(self):
        # 这是审查报告里的 H-4：1-999999 必须被拒
        with pytest.raises(ValidationError):
            PdfService._parse_pages("1-999999")

    def test_oversized_range_blocked(self):
        too_wide = MAX_PDF_PAGE_RANGE_SIZE + 5
        with pytest.raises(ValidationError, match="跨度"):
            PdfService._parse_pages(f"1-{too_wide}")

    def test_invalid_format_rejected(self):
        with pytest.raises(ValidationError):
            PdfService._parse_pages("abc")

    def test_empty_input_rejected(self):
        with pytest.raises(ValidationError):
            PdfService._parse_pages("")


class TestPdfServiceMerge:
    def test_merge_empty_inputs(self, tmp_path):
        svc = _make_svc()
        with pytest.raises(ValidationError, match="至少需要一个输入文件"):
            svc.merge([], tmp_path / "out.pdf")


class TestPdfServiceWatermark:
    def test_watermark_text_too_long(self, tmp_path):
        svc = _make_svc()
        with pytest.raises(ValidationError, match="过长"):
            svc.watermark(tmp_path / "in.pdf", "x" * 200, tmp_path / "out.pdf")
