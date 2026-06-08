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

    # ── Refresh 刷新 ──
    # 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

    def refresh(self, app: Any) -> dict:
        """刷新演示文稿链接

        调用 ActivePresentation.UpdateLinks() 更新所有 OLE 链接、
        图表数据和嵌入对象。
        """
        pres = app.ActivePresentation
        result: dict = {"actions": []}

        try:
            pres.UpdateLinks()
            result["actions"].append(
                {"action": "update_links", "status": "ok"}
            )
        except Exception as e:
            result["actions"].append(
                {"action": "update_links", "status": "error", "message": str(e)}
            )

        return result

    # ── 语义视图与诊断 ──

    def summarize(self, app: Any) -> dict:
        """生成演示文稿结构摘要（L1 语义视图）

        设计参考: iOfficeAI/OfficeCLI (Apache 2.0)

        返回包含幻灯片概览、切换效果、备注统计等结构化信息的 dict。
        """
        pres = app.ActivePresentation

        # 元数据
        metadata = {
            "title": str(pres.BuiltInDocumentProperties("Title").Value or ""),
            "author": str(pres.BuiltInDocumentProperties("Author").Value or ""),
            "slides": pres.Slides.Count,
        }

        # 幻灯片概览
        slides = []
        total_shapes = 0
        total_notes = 0
        for i in range(1, pres.Slides.Count + 1):
            sl = pres.Slides(i)

            # 标题
            title = ""
            shape_count = sl.Shapes.Count
            total_shapes += shape_count

            for shape in sl.Shapes:
                try:
                    if (
                        shape.HasTextFrame
                        and hasattr(shape, "PlaceholderFormat")
                        and shape.PlaceholderFormat.Type == PP_PLACEHOLDER_TITLE
                    ):
                        title = shape.TextFrame.TextRange.Text.strip()[:80]
                        break
                except Exception:
                    continue

            # 备注
            has_notes = False
            notes_text = ""
            try:
                notes_text = (
                    sl.NotesPage.Shapes.Placeholders(PP_PLACEHOLDER_BODY)
                    .TextFrame.TextRange.Text.strip()
                )
                has_notes = bool(notes_text)
                if has_notes:
                    total_notes += 1
            except Exception:
                pass

            # 切换效果
            transition_info = {}
            try:
                tr = sl.SlideShowTransition
                transition_info = {
                    "effect": tr.EntryEffect,
                    "duration": tr.Duration,
                }
            except Exception:
                pass

            slides.append(
                {
                    "index": i,
                    "title": title,
                    "shapes": shape_count,
                    "has_notes": has_notes,
                    "notes_preview": notes_text[:100] if has_notes else "",
                    "transition": transition_info,
                }
            )

        return {
            "metadata": metadata,
            "slides": slides,
            "stats": {
                "total_shapes": total_shapes,
                "slides_with_notes": total_notes,
            },
        }

    def diagnose(self, app: Any) -> list[dict]:
        """诊断演示文稿问题（参考 OfficeCLI view issues）

        设计参考: iOfficeAI/OfficeCLI (Apache 2.0)
        """
        from wps_cli.services.document_diagnostics import DocumentDiagnostics

        diag = DocumentDiagnostics()
        issues = diag.diagnose_impress(app)
        return [
            {
                "severity": i.severity,
                "category": i.category,
                "subtype": i.subtype,
                "location": i.location,
                "message": i.message,
                "suggestion": i.suggestion,
            }
            for i in issues
        ]

    def annotate(self, app: Any) -> list[str]:
        """输出演示文稿带路径标注的摘要（参考 OfficeCLI view annotated）

        设计参考: iOfficeAI/OfficeCLI (Apache 2.0)
        """
        pres = app.ActivePresentation
        lines: list[str] = []

        for i in range(1, pres.Slides.Count + 1):
            try:
                sl = pres.Slides(i)
                lines.append(f"[/presentation/slide[{i}] "
                             f"layout={sl.Layout}]")

                for j, sh in enumerate(sl.Shapes, 1):
                    try:
                        shape_type = str(sh.Type) if hasattr(sh, "Type") else "unknown"
                        shape_name = str(sh.Name) if hasattr(sh, "Name") else "shape"
                        alt_text = (
                            str(sh.AlternativeText)
                            if hasattr(sh, "AlternativeText") and sh.AlternativeText
                            else "(无)"
                        )
                        text_preview = ""
                        if sh.HasTextFrame:
                            try:
                                text_preview = sh.TextFrame.TextRange.Text[:60].strip()
                            except Exception:
                                pass

                        lines.append(
                            f"[/presentation/slide[{i}]/shape[{j}] "
                            f"name={shape_name} type={shape_type} "
                            f"alt_text={alt_text}] {text_preview}"
                        )
                    except Exception:
                        pass

            except Exception:
                pass

        return lines

    def get_stats(self, app: Any) -> dict:
        """获取纯数字统计信息（参考 OfficeCLI view stats）

        设计参考: iOfficeAI/OfficeCLI (Apache 2.0)
        """
        pres = app.ActivePresentation
        stats: dict = {}

        try:
            stats["slides"] = pres.Slides.Count
        except Exception:
            stats["slides"] = 0

        total_shapes = 0
        total_text_frames = 0
        total_images = 0
        slides_with_notes = 0

        for i in range(1, stats["slides"] + 1):
            try:
                sl = pres.Slides(i)
                try:
                    shape_count = sl.Shapes.Count
                    total_shapes += shape_count
                except Exception:
                    pass

                # 统计形状类型
                try:
                    for sh in sl.Shapes:
                        try:
                            if sh.HasTextFrame:
                                total_text_frames += 1
                            if hasattr(sh, "Type") and sh.Type == 13:  # msoPicture
                                total_images += 1
                        except Exception:
                            pass
                except Exception:
                    pass

                # 备注
                try:
                    notes_text = (
                        sl.NotesPage.Shapes.Placeholders(PP_PLACEHOLDER_BODY)
                        .TextFrame.TextRange.Text.strip()
                    )
                    if notes_text:
                        slides_with_notes += 1
                except Exception:
                    pass

            except Exception:
                pass

        stats["total_shapes"] = total_shapes
        stats["total_text_frames"] = total_text_frames
        stats["total_images"] = total_images
        stats["slides_with_notes"] = slides_with_notes

        return stats
