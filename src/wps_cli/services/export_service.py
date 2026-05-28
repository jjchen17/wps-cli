"""统一导出服务"""

from pathlib import Path

from wps_cli.services.session_manager import SessionManager


class ExportService:
    """跨应用格式转换"""

    def __init__(self, manager: SessionManager):
        self.manager = manager

    FORMAT_MAP: dict[str, int] = {
        "docx": 16,    # wdFormatDocumentDefault
        "doc": 0,      # wdFormatDocument
        "rtf": 6,      # wdFormatRTF
        "txt": 2,      # wdFormatText
        "html": 8,     # wdFormatHTML
        "pdf": 17,     # wdFormatPDF
        "xlsx": 51,    # xlOpenXMLWorkbook
        "csv": 6,      # xlCSV
        "pptx": 24,    # ppSaveAsOpenXMLPresentation
        "ppt": 0,      # ppSaveAsPresentation
    }

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

    def convert(self, input_path: Path, output_format: str, output_path: Path | None = None) -> Path:
        """通用格式转换"""
        app_type = self.detect_app_type(input_path)
        fmt_code = self.FORMAT_MAP.get(output_format)
        if fmt_code is None:
            raise ValueError(f"不支持的目标格式: {output_format}")

        if output_path is None:
            output_path = input_path.with_suffix(f".{output_format}")

        with self.manager.session(app_type) as app:
            if app_type == "writer":
                doc = app.Documents.Open(str(input_path))
                doc.SaveAs(str(output_path), fmt_code)
                doc.Close(0)
            elif app_type == "calc":
                wb = app.Workbooks.Open(str(input_path))
                wb.SaveAs(str(output_path), fmt_code)
                wb.Close(0)
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
