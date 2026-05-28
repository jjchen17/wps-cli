"""PDF 处理业务逻辑"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from wps_cli.consts import (
    MAX_PDF_PAGE_NUMBER,
    MAX_PDF_PAGE_RANGE_SIZE,
    MSO_TEXT_EFFECT_1,
    WD_DO_NOT_SAVE_CHANGES,
    WD_FORMAT_PDF,
    WD_HEADER_FOOTER_PRIMARY,
    WD_PAGE,
    WD_STATISTIC_PAGES,
    WD_STORY,
)
from wps_cli.exceptions import ValidationError
from wps_cli.services.session_manager import SessionManager


@dataclass
class PdfService:
    """PDF 文档操作（基于 WPS 导出能力）"""

    manager: SessionManager

    def info(self, path: Path) -> dict:
        """获取 PDF 元信息"""
        stat = path.stat()
        return {
            "path": str(path),
            "size_bytes": stat.st_size,
            "size_mb": round(stat.st_size / 1024 / 1024, 2),
            "modified": stat.st_mtime,
        }

    def merge(self, inputs: list[Path], output: Path) -> Path:
        """合并多个 PDF（通过 Writer 打开再导出）"""
        if not inputs:
            raise ValidationError("至少需要一个输入文件")
        with self.manager.session("writer") as app:
            doc = app.Documents.Open(
                str(inputs[0]),
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            for p in inputs[1:]:
                sel = app.Selection
                sel.EndKey(WD_STORY)
                sel.InsertFile(str(p))
            doc.SaveAs(str(output), WD_FORMAT_PDF)
            doc.Close(WD_DO_NOT_SAVE_CHANGES)
        return output

    def extract_pages(self, input_path: Path, pages: str, output: Path) -> Path:
        """提取指定页面（pages 格式: "1-3,5,7-9"）"""
        page_list = self._parse_pages(pages)
        with self.manager.session("writer") as app:
            doc = app.Documents.Open(
                str(input_path),
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            total = doc.ComputeStatistics(WD_STATISTIC_PAGES)
            pages_to_delete = sorted(
                [p for p in range(1, total + 1) if p not in page_list], reverse=True
            )
            for page_num in pages_to_delete:
                start_rng = doc.Range()
                start_rng.GoTo(WD_PAGE, 1, page_num)
                start_pos = start_rng.Start
                if page_num < total:
                    end_rng = doc.Range()
                    end_rng.GoTo(WD_PAGE, 1, page_num + 1)
                    end_pos = end_rng.Start
                else:
                    end_pos = doc.Range().End
                page_rng = doc.Range(start_pos, end_pos)
                page_rng.Delete()
            doc.SaveAs(str(output), WD_FORMAT_PDF)
            doc.Close(WD_DO_NOT_SAVE_CHANGES)
        return output

    def split(self, input_path: Path, every: int, output_dir: Path) -> list[Path]:
        """按每 N 页拆分"""
        if every <= 0:
            raise ValidationError("拆分粒度必须为正整数")
        output_dir.mkdir(parents=True, exist_ok=True)
        with self.manager.session("writer") as app:
            doc = app.Documents.Open(
                str(input_path),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
            )
            total = doc.ComputeStatistics(WD_STATISTIC_PAGES)
            results = []
            for start in range(1, total + 1, every):
                end = min(start + every - 1, total)
                rng = doc.Range()
                rng.GoTo(WD_PAGE, 1, start)
                start_pos = rng.Start
                if end < total:
                    end_rng = doc.Range()
                    end_rng.GoTo(WD_PAGE, 1, end + 1)
                    end_pos = end_rng.Start
                else:
                    end_pos = doc.Range().End
                part_rng = doc.Range(start_pos, end_pos)
                part_path = output_dir / f"part_{start}-{end}.pdf"
                part_rng.Copy()
                new_doc = app.Documents.Add()
                new_doc.Range().Paste()
                new_doc.SaveAs(str(part_path), WD_FORMAT_PDF)
                new_doc.Close(WD_DO_NOT_SAVE_CHANGES)
                results.append(part_path)
            doc.Close(WD_DO_NOT_SAVE_CHANGES)
        return results

    def watermark(self, input_path: Path, text: str, output: Path) -> Path:
        """添加文字水印"""
        if len(text) > 100:
            raise ValidationError("水印文字过长（最多 100 字符）")
        with self.manager.session("writer") as app:
            doc = app.Documents.Open(
                str(input_path),
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            for section in doc.Sections:
                header = section.Headers(WD_HEADER_FOOTER_PRIMARY)
                shape = header.Shapes.AddTextEffect(
                    MSO_TEXT_EFFECT_1, text, "宋体", 36, False, False, 0, 0
                )
                shape.Rotation = -45
                shape.Fill.ForeColor.RGB = 0xCCCCCC
                shape.TextEffect.NormalizedHeight = False
            doc.SaveAs(str(output), WD_FORMAT_PDF)
            doc.Close(WD_DO_NOT_SAVE_CHANGES)
        return output

    @staticmethod
    def _parse_pages(pages: str) -> list[int]:
        """解析页码字符串: "1-3,5,7-9" -> [1,2,3,5,7,8,9]

        - 单个页码上限 :data:`MAX_PDF_PAGE_NUMBER`
        - 单个区间跨度上限 :data:`MAX_PDF_PAGE_RANGE_SIZE`
        - 总页码数上限 :data:`MAX_PDF_PAGE_RANGE_SIZE`

        防止 ``"1-999999"`` 这种内存炸弹。
        """
        if not pages or not pages.strip():
            raise ValidationError("页码字符串不能为空")
        result: list[int] = []
        for part in pages.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                start_s, end_s = part.split("-", 1)
                try:
                    start, end = int(start_s), int(end_s)
                except ValueError as exc:
                    raise ValidationError(f"页码格式无效: {part!r}") from exc
                if start < 1:
                    raise ValidationError(f"页码必须从 1 开始: {part}")
                if start > end:
                    raise ValidationError(f"页码范围无效: {start}-{end}（起始页不能大于结束页）")
                if end > MAX_PDF_PAGE_NUMBER:
                    raise ValidationError(f"页码超出上限 {MAX_PDF_PAGE_NUMBER}: {end}")
                if (end - start + 1) > MAX_PDF_PAGE_RANGE_SIZE:
                    raise ValidationError(
                        f"页码区间跨度超出 {MAX_PDF_PAGE_RANGE_SIZE}: {start}-{end}"
                    )
                result.extend(range(start, end + 1))
            else:
                try:
                    n = int(part)
                except ValueError as exc:
                    raise ValidationError(f"页码格式无效: {part!r}") from exc
                if n < 1 or n > MAX_PDF_PAGE_NUMBER:
                    raise ValidationError(f"页码超出范围 [1, {MAX_PDF_PAGE_NUMBER}]: {n}")
                result.append(n)
        if len(result) > MAX_PDF_PAGE_RANGE_SIZE:
            raise ValidationError(f"页码总数超出上限 {MAX_PDF_PAGE_RANGE_SIZE}")
        return sorted(set(result))
