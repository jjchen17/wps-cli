"""Impress 演示文稿操作业务逻辑"""

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from wps_cli.consts import (
    MSO_TEXT_ORIENTATION_HORIZONTAL,
    PP_PLACEHOLDER_BODY,
    PP_PLACEHOLDER_SUBTITLE,
    PP_PLACEHOLDER_TITLE,
    PP_SAVE_AS_PDF,
    PP_TRANSITION_RANDOM,
)
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
            pres = app.Presentations.Open(str(path), ReadOnly=True)
            result = {
                "path": str(Path(pres.FullName)),
                "slides": pres.Slides.Count,
                "title": pres.BuiltInDocumentProperties("Title").Value or "",
                "author": pres.BuiltInDocumentProperties("Author").Value or "",
            }
            pres.Close()
        return result

    # ── 幻灯片管理 ──

    def slide_list(self, app: Any) -> list[dict]:
        pres = app.ActivePresentation
        slides = []
        for i in range(1, pres.Slides.Count + 1):
            sl = pres.Slides(i)
            title = ""
            for shape in sl.Shapes:
                try:
                    if shape.HasTextFrame and shape.PlaceholderFormat.Type == PP_PLACEHOLDER_TITLE:
                        title = shape.TextFrame.TextRange.Text[:50]
                        break
                except AttributeError:
                    # 非占位符形状没有 PlaceholderFormat，跳过
                    continue
                except Exception:
                    # COM 调用偶发失败，跳过单个 shape，整体流程继续
                    continue
            has_notes = False
            try:
                notes_text = sl.NotesPage.Shapes.Placeholders(
                    PP_PLACEHOLDER_BODY
                ).TextFrame.TextRange.Text
                has_notes = bool(notes_text.strip())
            except (AttributeError, Exception):
                pass
            slides.append(
                {
                    "index": i,
                    "title": title,
                    "layout": sl.Layout,
                    "has_notes": has_notes,
                }
            )
        return slides

    def slide_add(self, app: Any, layout: int = 1, at: int | None = None, title: str = "") -> int:
        pres = app.ActivePresentation
        idx = at if at else pres.Slides.Count + 1
        sl = pres.Slides.Add(idx, layout)
        if title:
            for shape in sl.Shapes:
                if shape.HasTextFrame and shape.PlaceholderFormat.Type == PP_PLACEHOLDER_TITLE:
                    shape.TextFrame.TextRange.Text = title
                    break
        return idx

    def slide_delete(self, app: Any, index: int) -> None:
        pres = app.ActivePresentation
        pres.Slides(index).Delete()

    def slide_copy(self, app: Any, src: int, dest: int) -> None:
        pres = app.ActivePresentation
        pres.Slides(src).Copy()
        pres.Slides.Paste(dest)

    def slide_move(self, app: Any, from_idx: int, to_idx: int) -> None:
        pres = app.ActivePresentation
        pres.Slides(from_idx).MoveTo(to_idx)

    # ── 内容操作 ──

    def text_set(self, app: Any, slide_idx: int, placeholder: str, text: str) -> None:
        pres = app.ActivePresentation
        sl = pres.Slides(slide_idx)
        ph_map = {
            "title": PP_PLACEHOLDER_TITLE,
            "body": PP_PLACEHOLDER_BODY,
            "subtitle": PP_PLACEHOLDER_SUBTITLE,
        }
        ph_type = ph_map.get(placeholder, PP_PLACEHOLDER_TITLE)
        for shape in sl.Shapes:
            if shape.HasTextFrame and shape.PlaceholderFormat.Type == ph_type:
                shape.TextFrame.TextRange.Text = text
                break

    def text_get(self, app: Any, slide_idx: int) -> str:
        pres = app.ActivePresentation
        sl = pres.Slides(slide_idx)
        texts = []
        for shape in sl.Shapes:
            if shape.HasTextFrame:
                texts.append(shape.TextFrame.TextRange.Text)
        return "\n".join(texts)

    def textbox_add(
        self,
        app: Any,
        slide_idx: int,
        text: str,
        left: float = 100,
        top: float = 100,
        width: float = 400,
        height: float = 100,
    ) -> None:
        pres = app.ActivePresentation
        sl = pres.Slides(slide_idx)
        shape = sl.Shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL, left, top, width, height)
        shape.TextFrame.TextRange.Text = text

    def image_insert(
        self,
        app: Any,
        slide_idx: int,
        path: Path,
        left: float = 100,
        top: float = 100,
        width: float | None = None,
        height: float | None = None,
    ) -> None:
        pres = app.ActivePresentation
        sl = pres.Slides(slide_idx)
        shape = sl.Shapes.AddPicture(str(path), True, True, left, top)
        if width is not None:
            shape.Width = width
        if height is not None:
            shape.Height = height

    def notes_set(self, app: Any, slide_idx: int, text: str) -> None:
        pres = app.ActivePresentation
        sl = pres.Slides(slide_idx)
        sl.NotesPage.Shapes.Placeholders(PP_PLACEHOLDER_BODY).TextFrame.TextRange.Text = text

    def notes_get(self, app: Any, slide_idx: int) -> str:
        pres = app.ActivePresentation
        sl = pres.Slides(slide_idx)
        return sl.NotesPage.Shapes.Placeholders(PP_PLACEHOLDER_BODY).TextFrame.TextRange.Text

    # ── 切换效果 ──

    def transition_set(
        self, app: Any, slide_idx: int, effect: int = PP_TRANSITION_RANDOM, duration: float = 1.0
    ) -> None:
        """设置幻灯片切换效果。effect 为 WPS 常量 ID"""
        pres = app.ActivePresentation
        sl = pres.Slides(slide_idx)
        sl.SlideShowTransition.EntryEffect = effect
        sl.SlideShowTransition.Duration = duration

    # ── 保存与导出 ──

    def save(self, app: Any, path: Path | None = None) -> Path:
        pres = app.ActivePresentation
        if path:
            pres.SaveAs(str(path))
        else:
            pres.Save()
        return Path(pres.FullName)

    def export_pdf(self, app: Any, output: Path) -> Path:
        pres = app.ActivePresentation
        pres.SaveAs(str(output), PP_SAVE_AS_PDF)
        return output
