"""统一导出服务"""

from dataclasses import dataclass
from pathlib import Path

from wps_cli.consts import (
    CALC_FORMATS,
    IMPRESS_FORMATS,
    WD_DO_NOT_SAVE_CHANGES,
    WRITER_FORMATS,
)
from wps_cli.services.session_manager import SessionManager


@dataclass
class ExportService:
    """跨应用格式转换"""

    manager: SessionManager

    def detect_app_type(self, path: Path) -> str:
        """根据文件扩展名判断应用类型"""
        ext = path.suffix.lower()
        writer_exts = {".doc", ".docx", ".wps", ".rtf", ".txt", ".html"}
        calc_exts = {".xls", ".xlsx", ".et", ".csv"}
        impress_exts = {".ppt", ".pptx", ".dps"}
        if ext in writer_exts:
            return "writer"
        if ext in calc_exts:
            return "calc"
        if ext in impress_exts:
            return "impress"
        raise ValueError(f"不支持的文件格式: {ext}")

    def _get_format_map(self, app_type: str) -> dict[str, int]:
        """获取指定应用类型的格式映射"""
        if app_type == "writer":
            return WRITER_FORMATS
        if app_type == "calc":
            return CALC_FORMATS
        if app_type == "impress":
            return IMPRESS_FORMATS
        raise ValueError(f"未知应用类型: {app_type}")

    def convert(self, input_path: Path, output_format: str, output_path: Path | None = None) -> Path:
        """通用格式转换"""
        app_type = self.detect_app_type(input_path)
        fmt_map = self._get_format_map(app_type)
        fmt_code = fmt_map.get(output_format)
        if fmt_code is None:
            raise ValueError(f"不支持的目标格式: {output_format}")

        if output_path is None:
            output_path = input_path.with_suffix(f".{output_format}")

        with self.manager.session(app_type) as app:
            if app_type == "writer":
                doc = app.Documents.Open(str(input_path))
                doc.SaveAs(str(output_path), fmt_code)
                doc.Close(WD_DO_NOT_SAVE_CHANGES)
            elif app_type == "calc":
                wb = app.Workbooks.Open(str(input_path))
                wb.SaveAs(str(output_path), fmt_code)
                wb.Close(WD_DO_NOT_SAVE_CHANGES)
            elif app_type == "impress":
                pres = app.Presentations.Open(str(input_path))
                pres.SaveAs(str(output_path), fmt_code)
                pres.Close()

        return output_path

    def batch_convert(
        self, glob_pattern: str, output_format: str, output_dir: Path
    ) -> list[dict]:
        """批量格式转换"""
        import glob as glob_mod
        output_dir.mkdir(parents=True, exist_ok=True)
        files = glob_mod.glob(glob_pattern)
        results = []
        for f in files:
            fpath = Path(f)
            out = output_dir / fpath.with_suffix(f".{output_format}").name
            try:
                self.convert(fpath, output_format, out)
                results.append({"input": str(fpath), "output": str(out), "status": "ok"})
            except Exception as e:
                results.append({"input": str(fpath), "error": str(e), "status": "failed"})
        return results
