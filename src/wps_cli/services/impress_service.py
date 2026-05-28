"""Impress 演示文稿操作业务逻辑"""

from dataclasses import dataclass
from pathlib import Path

from wps_cli.services.session_manager import SessionManager


@dataclass
class ImpressService:
    """PPT 演示文稿操作"""

    manager: SessionManager

    # ── 文档生命周期 ──

    def new(self, output: Path | None = None) -> Path:
        with self.manager.session("impress") as app:
            pres = app.Presentations.Add()
            if output:
                pres.SaveAs(str(output))
            path = pres.FullName
            pres.Close()
        return Path(path)

    def info(self, path: Path) -> dict:
        with self.manager.session("impress") as app:
            pres = app.Presentations.Open(str(path))
            result = {
                "path": str(Path(pres.FullName)),
                "slides": pres.Slides.Count,
                "title": pres.BuiltInDocumentProperties("Title").Value or "",
                "author": pres.BuiltInDocumentProperties("Author").Value or "",
            }
            pres.Close()
        return result

    # ── 幻灯片管理 ──

    def slide_list(self, app: object) -> list[dict]:
        pres = app.ActivePresentation
        slides = []
        for i in range(1, pres.Slides.Count + 1):
            sl = pres.Slides(i)
            title = ""
            for shape in sl.Shapes:
                if shape.HasTextFrame and shape.PlaceholderFormat.Type == 1:
                    title = shape.TextFrame.TextRange.Text[:50]
                    break
            slides.append({
                "index": i,
                "title": title,
                "layout": sl.Layout,
                "has_notes": bool(sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text.strip()),
            })
        return slides

    def slide_add(self, app: object, layout: int = 1, at: int | None = None, title: str = "") -> int:
        pres = app.ActivePresentation
        idx = at if at else pres.Slides.Count + 1
        sl = pres.Slides.Add(idx, layout)
        if title:
            for shape in sl.Shapes:
                if shape.HasTextFrame and shape.PlaceholderFormat.Type == 1:
                    shape.TextFrame.TextRange.Text = title
                    break
        return idx

    def slide_delete(self, app: object, index: int) -> None:
        pres = app.ActivePresentation
        pres.Slides(index).Delete()

    def slide_copy(self, app: object, src: int, dest: int) -> None:
        pres = app.ActivePresentation
        pres.Slides(src).Copy()
        pres.Slides.Paste(dest)

    def slide_move(self, app: object, from_idx: int, to_idx: int) -> None:
        pres = app.ActivePresentation
        pres.Slides(from_idx).MoveTo(to_idx)

    # ── 内容操作 ──

    def text_set(self, app: object, slide_idx: int, placeholder: str, text: str) -> None:
        pres = app.ActivePresentation
        sl = pres.Slides(slide_idx)
        ph_map = {"title": 1, "body": 2, "subtitle": 3}
        ph_type = ph_map.get(placeholder, 1)
        for shape in sl.Shapes:
            if shape.HasTextFrame and shape.PlaceholderFormat.Type == ph_type:
                shape.TextFrame.TextRange.Text = text
                break

    def text_get(self, app: object, slide_idx: int) -> str:
        pres = app.ActivePresentation
        sl = pres.Slides(slide_idx)
        texts = []
        for shape in sl.Shapes:
            if shape.HasTextFrame:
                texts.append(shape.TextFrame.TextRange.Text)
        return "\n".join(texts)

    def textbox_add(
        self, app: object, slide_idx: int, text: str,
        left: float = 100, top: float = 100, width: float = 400, height: float = 100,
    ) -> None:
        pres = app.ActivePresentation
        sl = pres.Slides(slide_idx)
        shape = sl.Shapes.AddTextbox(1, left, top, width, height)  # msoTextOrientationHorizontal
        shape.TextFrame.TextRange.Text = text

    def image_insert(
        self, app: object, slide_idx: int, path: Path,
        left: float = 100, top: float = 100, width: float | None = None, height: float | None = None,
    ) -> None:
        pres = app.ActivePresentation
        sl = pres.Slides(slide_idx)
        shape = sl.Shapes.AddPicture(str(path), True, True, left, top)
        if width:
            shape.Width = width
        if height:
            shape.Height = height

    def notes_set(self, app: object, slide_idx: int, text: str) -> None:
        pres = app.ActivePresentation
        sl = pres.Slides(slide_idx)
        sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = text

    def notes_get(self, app: object, slide_idx: int) -> str:
        pres = app.ActivePresentation
        sl = pres.Slides(slide_idx)
        return sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text

    # ── 切换效果 ──

    def transition_set(self, app: object, slide_idx: int, effect: int = 3844, duration: float = 1.0) -> None:
        """设置幻灯片切换效果。effect 为 WPS 常量 ID"""
        pres = app.ActivePresentation
        sl = pres.Slides(slide_idx)
        sl.SlideShowTransition.EntryEffect = effect
        sl.SlideShowTransition.Duration = duration

    # ── 保存与导出 ──

    def save(self, app: object, path: Path | None = None) -> Path:
        pres = app.ActivePresentation
        if path:
            pres.SaveAs(str(path))
        else:
            pres.Save()
        return Path(pres.FullName)

    def export_pdf(self, app: object, output: Path) -> Path:
        pres = app.ActivePresentation
        pres.SaveAs(str(output), 32)  # ppSaveAsPDF
        return output
