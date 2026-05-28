"""样式管理引擎"""

from dataclasses import dataclass
from typing import Any, ClassVar

from wps_cli.consts import (
    ALIGN_CENTER,
    ALIGN_JUSTIFY,
    ALIGN_LEFT,
    ALIGN_RIGHT,
    WD_LINE_SPACE_MULTIPLE,
)


@dataclass
class StyleEngine:
    """管理文档样式模板"""

    # 公文格式预设（GB/T 9704-2012）
    PRESETS: ClassVar[dict[str, dict]] = {
        "公文标题": {
            "font": "方正小标宋简体",
            "size": 22,
            "bold": True,
            "align": "center",
        },
        "公文一级标题": {
            "font": "黑体",
            "size": 16,
            "bold": False,
            "align": "left",
        },
        "公文二级标题": {
            "font": "楷体_GB2312",
            "size": 16,
            "bold": True,
            "align": "left",
        },
        "公文正文": {
            "font": "仿宋_GB2312",
            "size": 16,
            "bold": False,
            "align": "justify",
            "first_indent": 32,
            "line_spacing": 1.5,
        },
        "报告标题": {
            "font": "微软雅黑",
            "size": 28,
            "bold": True,
            "align": "center",
        },
        "报告正文": {
            "font": "微软雅黑",
            "size": 12,
            "bold": False,
            "align": "justify",
            "first_indent": 24,
            "line_spacing": 1.5,
        },
    }

    def get_preset(self, name: str) -> dict:
        """获取样式预设"""
        if name not in self.PRESETS:
            raise ValueError(f"未知样式预设: {name}，可选: {list(self.PRESETS.keys())}")
        return self.PRESETS[name]

    def list_presets(self) -> list[str]:
        """列出所有可用预设"""
        return list(self.PRESETS.keys())

    def apply_preset(self, app: Any, preset_name: str) -> None:
        """将预设应用到当前选区"""
        preset = self.get_preset(preset_name)
        sel = app.Selection

        if "font" in preset:
            sel.Font.Name = preset["font"]
            sel.Font.NameFarEast = preset["font"]
        if "size" in preset:
            sel.Font.Size = preset["size"]
        if "bold" in preset:
            sel.Font.Bold = preset["bold"]

        pf = sel.ParagraphFormat
        align_map = {"left": ALIGN_LEFT, "center": ALIGN_CENTER, "right": ALIGN_RIGHT, "justify": ALIGN_JUSTIFY}
        if "align" in preset:
            pf.Alignment = align_map.get(preset["align"], ALIGN_LEFT)
        if "first_indent" in preset:
            pf.FirstLineIndent = preset["first_indent"]
        if "line_spacing" in preset:
            pf.LineSpacingRule = WD_LINE_SPACE_MULTIPLE
            pf.LineSpacing = preset["line_spacing"] * 12
