"""Dump 往返序列化

将 Word/PPT 文档序列化为可重放的 batch JSON 指令数组。
支持全文档 dump 和子树 dump（通过路径语法指定范围）。

设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from wps_cli.services.session_manager import SessionManager


def _safe_str(value: Any) -> str:
    """安全转换为字符串，处理 None 和 COM 对象"""
    if value is None:
        return ""
    try:
        return str(value)
    except Exception:
        return ""


@dataclass
class DumpService:
    """文档序列化服务

    将 WPS 文档（Writer/Impress）的结构和内容序列化为
    可重放的 batch JSON 指令数组。支持全文档 dump 和
    通过路径过滤的子树 dump。

    设计理念（参考 OfficeCLI）：
    - dump 输出是一组 batch 命令，可被 ``wps batch`` 重放
    - 支持"从样本学习 → 批量生成变体"的工作流
    - 每个命令包含 action + params，语义自描述
    """

    manager: SessionManager
    _indent: int = field(default=2, repr=False)

    # ── Writer dump ─────────────────────────────────────────────

    def dump_writer(self, app: Any, path_filter: str = "") -> list[dict]:
        """序列化 Word 文档为 batch 命令列表

        遍历顺序: 页面设置 → 各段落(文本+样式) → 表格 → 图片 → 页眉页脚

        Args:
            app: WPS Writer 应用对象
            path_filter: 路径过滤表达式 (如 "/section[1]/table[2]")，
                        为空则 dump 全部

        Returns:
            batch 命令列表，每条包含 action + params
        """
        commands: list[dict] = []
        doc = app.ActiveDocument

        # 1. 页面设置
        commands.append(self._dump_page_setup(doc))

        # 2. 段落
        if not path_filter or "paragraph" in path_filter or "section" in path_filter:
            for i in range(1, doc.Paragraphs.Count + 1):
                try:
                    para = doc.Paragraphs(i)
                    cmd = self._dump_paragraph(para, i)
                    commands.append(cmd)
                except Exception:
                    pass

        # 3. 表格
        if not path_filter or "table" in path_filter:
            for i in range(1, doc.Tables.Count + 1):
                try:
                    table = doc.Tables(i)
                    cmd = self._dump_table(table, i)
                    commands.append(cmd)
                except Exception:
                    pass

        # 4. 图片 (InlineShapes)
        if not path_filter or "image" in path_filter or "shape" in path_filter:
            for i in range(1, doc.InlineShapes.Count + 1):
                try:
                    shape = doc.InlineShapes(i)
                    cmd = self._dump_inline_shape(shape, i)
                    commands.append(cmd)
                except Exception:
                    pass

        # 5. 页眉页脚
        if not path_filter or "header" in path_filter or "footer" in path_filter:
            try:
                for section_i in range(1, doc.Sections.Count + 1):
                    section = doc.Sections(section_i)
                    # 页眉
                    try:
                        header = section.Headers(1)  # wdHeaderFooterPrimary
                        if header.Exists:
                            hdr_cmd = self._dump_header_footer(header, "header", section_i)
                            commands.append(hdr_cmd)
                    except Exception:
                        pass
                    # 页脚
                    try:
                        footer = section.Footers(1)
                        if footer.Exists:
                            ftr_cmd = self._dump_header_footer(footer, "footer", section_i)
                            commands.append(ftr_cmd)
                    except Exception:
                        pass
            except Exception:
                pass

        return commands

    def _dump_page_setup(self, doc: Any) -> dict:
        """序列化页面设置"""
        try:
            ps = doc.PageSetup
            return {
                "action": "writer.page_setup",
                "params": {
                    "top_margin": round(float(ps.TopMargin), 1),
                    "bottom_margin": round(float(ps.BottomMargin), 1),
                    "left_margin": round(float(ps.LeftMargin), 1),
                    "right_margin": round(float(ps.RightMargin), 1),
                    "page_width": round(float(ps.PageWidth), 1),
                    "page_height": round(float(ps.PageHeight), 1),
                    "orientation": _safe_str(ps.Orientation),
                },
            }
        except Exception:
            return {"action": "writer.page_setup", "params": {}, "_error": "无法读取页面设置"}

    def _dump_paragraph(self, para: Any, index: int) -> dict:
        """序列化段落"""
        try:
            text = _safe_str(para.Range.Text).rstrip("\r\x07")
            alignment = _safe_str(para.Alignment) if hasattr(para, "Alignment") else ""
            style = _safe_str(para.Style.NameLocal) if hasattr(para, "Style") else ""
        except Exception:
            text, alignment, style = "", "", ""

        # 提取字体信息
        font_info: dict = {}
        try:
            rng = para.Range
            font_info = {
                "name": _safe_str(rng.Font.Name),
                "size": round(float(rng.Font.Size), 1) if rng.Font.Size else 0,
                "bold": bool(rng.Font.Bold),
                "italic": bool(rng.Font.Italic),
                "color": _safe_str(rng.Font.Color),
            }
        except Exception:
            pass

        return {
            "action": "writer.insert_text",
            "params": {
                "index": index,
                "text": text,
                "style": style,
                "alignment": alignment,
                **({"font": font_info} if font_info else {}),
            },
        }

    def _dump_table(self, table: Any, index: int) -> dict:
        """序列化表格"""
        rows_data: list[list[str]] = []
        try:
            for r in range(1, table.Rows.Count + 1):
                row_data: list[str] = []
                for c in range(1, table.Columns.Count + 1):
                    try:
                        cell_text = _safe_str(table.Cell(r, c).Range.Text).rstrip("\r\x07")
                        row_data.append(cell_text)
                    except Exception:
                        row_data.append("")
                rows_data.append(row_data)
        except Exception:
            pass

        return {
            "action": "writer.insert_table",
            "params": {
                "index": index,
                "rows": len(rows_data),
                "cols": len(rows_data[0]) if rows_data else 0,
                "data": rows_data,
            },
        }

    def _dump_inline_shape(self, shape: Any, index: int) -> dict:
        """序列化内嵌形状/图片"""
        try:
            width = round(float(shape.Width), 1)
            height = round(float(shape.Height), 1)
            alt_text = _safe_str(getattr(shape, "AlternativeText", ""))
        except Exception:
            width, height, alt_text = 0, 0, ""

        result: dict = {
            "action": "writer.insert_image",
            "params": {
                "index": index,
                "width": width,
                "height": height,
                "alt_text": alt_text,
            },
        }

        # 尝试提取图片文件路径
        try:
            if hasattr(shape, "LinkFormat") and shape.LinkFormat:
                result["params"]["source_path"] = _safe_str(shape.LinkFormat.SourceFullName)
        except Exception:
            pass

        return result

    def _dump_header_footer(self, hf: Any, kind: str, section_index: int) -> dict:
        """序列化页眉或页脚"""
        paragraphs: list[dict] = []
        try:
            rng = hf.Range
            for i in range(1, rng.Paragraphs.Count + 1):
                try:
                    p = rng.Paragraphs(i)
                    paragraphs.append({
                        "text": _safe_str(p.Range.Text).rstrip("\r\x07"),
                        "alignment": _safe_str(p.Alignment) if hasattr(p, "Alignment") else "",
                    })
                except Exception:
                    pass
        except Exception:
            pass

        return {
            "action": f"writer.set_{kind}",
            "params": {
                "section_index": section_index,
                "paragraphs": paragraphs,
            },
        }

    # ── Impress dump ────────────────────────────────────────────

    def dump_impress(self, app: Any, path_filter: str = "") -> list[dict]:
        """序列化 PPT 演示文稿为 batch 命令列表

        遍历顺序: 幻灯片尺寸 → 各幻灯片 → 各形状(文本/图片/表格)

        Args:
            app: WPS Impress 应用对象
            path_filter: 路径过滤表达式，为空则 dump 全部

        Returns:
            batch 命令列表
        """
        commands: list[dict] = []
        pres = app.ActivePresentation

        # 幻灯片尺寸
        try:
            commands.append({
                "action": "impress.slide_setup",
                "params": {
                    "width": round(float(pres.PageSetup.SlideWidth), 1),
                    "height": round(float(pres.PageSetup.SlideHeight), 1),
                },
            })
        except Exception:
            pass

        # 各幻灯片
        for i in range(1, pres.Slides.Count + 1):
            try:
                slide = pres.Slides(i)
                slide_cmds = self._dump_slide(slide, i)
                commands.extend(slide_cmds)
            except Exception:
                pass

        return commands

    def _dump_slide(self, slide: Any, slide_index: int) -> list[dict]:
        """序列化单张幻灯片"""
        cmds: list[dict] = []

        # 幻灯片布局
        try:
            cmds.append({
                "action": "impress.set_layout",
                "params": {
                    "slide_index": slide_index,
                    "layout_name": _safe_str(slide.Layout.Name) if slide.Layout else "",
                },
            })
        except Exception:
            pass

        # 各形状
        for j in range(1, slide.Shapes.Count + 1):
            try:
                shape = slide.Shapes(j)
                shape_cmd = self._dump_shape(shape, slide_index, j)
                if shape_cmd:
                    cmds.append(shape_cmd)
            except Exception:
                pass

        return cmds

    def _dump_shape(self, shape: Any, slide_index: int, shape_index: int) -> dict | None:
        """序列化形状（文本框/图片/表格等）"""
        try:
            shape_type = int(shape.Type)
        except Exception:
            return None

        base_params = {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "left": round(float(shape.Left), 1),
            "top": round(float(shape.Top), 1),
            "width": round(float(shape.Width), 1),
            "height": round(float(shape.Height), 1),
        }

        # 1 = msoAutoShape, 14 = msoPlaceholder, 17 = msoTextBox
        if shape_type in (1, 14, 17):
            text = ""
            try:
                if shape.HasTextFrame:
                    text = _safe_str(shape.TextFrame.TextRange.Text)
            except Exception:
                pass
            return {
                "action": "impress.set_text",
                "params": {**base_params, "text": text, "shape_type": shape_type},
            }

        # 13 = msoPicture
        if shape_type == 13:
            alt_text = _safe_str(getattr(shape, "AlternativeText", ""))
            return {
                "action": "impress.insert_image",
                "params": {**base_params, "alt_text": alt_text},
            }

        # 19 = msoTable
        if shape_type == 19:
            rows_data: list[list[str]] = []
            try:
                table = shape.Table
                for r in range(1, table.Rows.Count + 1):
                    row = [_safe_str(table.Cell(r, c).Shape.TextFrame.TextRange.Text)
                           for c in range(1, table.Columns.Count + 1)]
                    rows_data.append(row)
            except Exception:
                pass
            return {
                "action": "impress.insert_table",
                "params": {**base_params, "data": rows_data},
            }

        # 其他形状类型：只记录位置信息
        return {
            "action": "impress.shape_info",
            "params": {**base_params, "shape_type": shape_type},
        }

    # ── 持久化辅助 ──────────────────────────────────────────────

    def to_file(self, commands: list[dict], output: Path) -> Path:
        """将 batch 命令列表写入 JSON 文件"""
        output.parent.mkdir(parents=True, exist_ok=True)
        with open(output, "w", encoding="utf-8") as f:
            json.dump(commands, f, ensure_ascii=False, indent=self._indent)
        return output

    def from_file(self, path: Path) -> list[dict]:
        """从 JSON 文件读取 batch 命令列表"""
        with open(path, encoding="utf-8") as f:
            return json.load(f)

    # ── Batch 回放 ─────────────────────────────────────────────

    def replay_writer(self, app: Any, commands: list[dict]) -> int:
        """将 batch 命令列表回放到 Writer 文档

        Returns:
            成功执行的命令数
        """
        executed = 0
        known_actions = {
            "writer.page_setup",
            "writer.insert_text",
            "writer.insert_table",
            "writer.insert_image",
            "writer.set_header",
            "writer.set_footer",
        }
        for cmd in commands:
            action = cmd.get("action", "")
            if action not in known_actions:
                continue
            params = cmd.get("params", {})

            try:
                if action == "writer.page_setup":
                    self._replay_page_setup(app.ActiveDocument, params)
                elif action == "writer.insert_text":
                    self._replay_insert_text(app.ActiveDocument, params)
                elif action == "writer.insert_table":
                    self._replay_insert_table(app.ActiveDocument, params)
                elif action == "writer.insert_image":
                    self._replay_insert_image(app.ActiveDocument, params)
                executed += 1
            except Exception:
                pass

        return executed

    def replay_impress(self, app: Any, commands: list[dict]) -> int:
        """将 batch 命令列表回放到 Impress 文档"""
        executed = 0
        pres = app.ActivePresentation
        for cmd in commands:
            try:
                action = cmd.get("action", "")
                params = cmd.get("params", {})
                if action == "impress.slide_setup":
                    ps = pres.PageSetup
                    ps.SlideWidth = params.get("width", ps.SlideWidth)
                    ps.SlideHeight = params.get("height", ps.SlideHeight)
                elif action == "impress.set_text":
                    slide = pres.Slides(params["slide_index"])
                    shape = slide.Shapes(params["shape_index"])
                    if shape.HasTextFrame:
                        shape.TextFrame.TextRange.Text = params.get("text", "")
                executed += 1
            except Exception:
                pass

        return executed

    # ── 回放辅助 ───────────────────────────────────────────────

    @staticmethod
    def _replay_page_setup(doc: Any, params: dict) -> None:
        ps = doc.PageSetup
        if "top_margin" in params:
            ps.TopMargin = params["top_margin"]
        if "bottom_margin" in params:
            ps.BottomMargin = params["bottom_margin"]
        if "left_margin" in params:
            ps.LeftMargin = params["left_margin"]
        if "right_margin" in params:
            ps.RightMargin = params["right_margin"]
        if "page_width" in params:
            ps.PageWidth = params["page_width"]
        if "page_height" in params:
            ps.PageHeight = params["page_height"]

    @staticmethod
    def _replay_insert_text(doc: Any, params: dict) -> None:
        text = params.get("text", "")
        if not text:
            return
        idx = params.get("index", 0)
        if idx > 0 and idx <= doc.Paragraphs.Count:
            para = doc.Paragraphs(idx)
            para.Range.Text = text
            if "alignment" in params and params["alignment"]:
                try:
                    para.Alignment = int(params["alignment"])
                except (ValueError, TypeError):
                    pass

    @staticmethod
    def _replay_insert_table(doc: Any, params: dict) -> None:
        data = params.get("data", [])
        if not data:
            return
        rows = len(data)
        cols = len(data[0]) if rows > 0 else 1
        table = doc.Tables.Add(doc.Range(0, 0), rows, cols)
        for r_idx, row_data in enumerate(data, start=1):
            for c_idx, cell_value in enumerate(row_data, start=1):
                try:
                    table.Cell(r_idx, c_idx).Range.Text = cell_value
                except Exception:
                    pass

    @staticmethod
    def _replay_insert_image(doc: Any, params: dict) -> None:
        source = params.get("source_path", "")
        if source and os.path.exists(source):
            doc.InlineShapes.AddPicture(source)


# ── 独立函数：直接从文件 dump ──────────────────────────────────


def dump_writer_to_file(
    manager: SessionManager, input_path: Path, output_path: Path, path_filter: str = ""
) -> int:
    """打开文档 → dump → 写入 JSON 文件 → 返回命令数"""
    svc = DumpService(manager)
    with manager.session("writer") as app:
        app.Documents.Open(str(input_path), ReadOnly=True)
        commands = svc.dump_writer(app, path_filter)
        svc.to_file(commands, output_path)
        app.ActiveDocument.Close()
    return len(commands)


def dump_impress_to_file(
    manager: SessionManager, input_path: Path, output_path: Path, path_filter: str = ""
) -> int:
    """打开演示文稿 → dump → 写入 JSON 文件 → 返回命令数"""
    svc = DumpService(manager)
    with manager.session("impress") as app:
        app.Presentations.Open(str(input_path), ReadOnly=True)
        commands = svc.dump_impress(app, path_filter)
        svc.to_file(commands, output_path)
        app.ActivePresentation.Close()
    return len(commands)


def batch_replay_to_file(
    manager: SessionManager,
    app_type: str,
    input_json: Path,
    output_path: Path,
) -> int:
    """读取 JSON → 回放到新文档 → 保存 → 返回命令数"""
    svc = DumpService(manager)
    commands = svc.from_file(input_json)

    with manager.session(app_type) as app:
        if app_type == "writer":
            app.Documents.Add()
            executed = svc.replay_writer(app, commands)
            app.ActiveDocument.SaveAs(str(output_path))
        elif app_type == "impress":
            app.Presentations.Add()
            executed = svc.replay_impress(app, commands)
            app.ActivePresentation.SaveAs(str(output_path))
        else:
            raise ValueError(f"不支持的文档类型: {app_type}")

    return executed
