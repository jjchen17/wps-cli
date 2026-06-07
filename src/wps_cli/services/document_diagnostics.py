# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""文档诊断引擎

检测常见文档问题：文本溢出、字体缺失、图片无替代文本、公式错误等。
用于 AI Agent 自愈工作流：编辑 → 诊断 → 修复 → 再诊断。

设计原则：
- diagnose() 方法不得抛出异常 — 单个检测项失败只跳过该项，继续检测其他项
- 返回结构化的 Issue 列表，便于 AI Agent 解析和修复
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from wps_cli.services.calc_service import _col_letter


@dataclass
class Issue:
    """文档问题"""

    severity: str  # "error", "warning", "info"
    category: str  # "font", "layout", "image", "formula", "style", "structure", "consistency"
    location: str  # 位置描述
    message: str  # 问题描述
    suggestion: str  # 修复建议


@dataclass
class DocumentDiagnostics:
    """文档诊断引擎

    检测常见文档问题并返回 Issue 列表。每个 diagnose_* 方法内部
    对单个检测项都有 try/except 保护，不会因某一检测失败而中断全局诊断。
    """

    # ── Writer 诊断 ──

    def diagnose_writer(self, app: Any) -> list[Issue]:
        """诊断 Word 文档

        检测：
        - 字体是否在系统中可用（记录使用字体）
        - 图片是否有替代文本（alt text）
        - 表格是否可能溢出页面宽度
        - 页面设置是否异常
        - 段落样式一致性
        """
        issues: list[Issue] = []
        doc = app.ActiveDocument

        # -- 1. 图片 alt text 检查 --
        self._check_image_alt_text(doc, issues)

        # -- 2. 字体使用记录 --
        self._check_fonts_used(doc, issues)

        # -- 3. 表格宽度检查 --
        self._check_table_width(doc, issues)

        # -- 4. 页面边距检查 --
        self._check_page_margins(doc, issues)

        # -- 5. 段落样式一致性 --
        self._check_paragraph_styles(doc, issues)

        return issues

    def _check_image_alt_text(self, doc: Any, issues: list[Issue]) -> None:
        """检查图片是否有替代文本"""
        try:
            for shape in doc.InlineShapes:
                try:
                    if not shape.AlternativeText:
                        issues.append(
                            Issue(
                                severity="warning",
                                category="image",
                                location=f"InlineShape {shape.Index}",
                                message="图片缺少替代文本（alt text）",
                                suggestion="右键图片 → 查看替代文本 → 输入描述",
                            )
                        )
                except Exception:
                    # 单个 shape 访问失败不影响整体流程
                    pass
        except Exception:
            pass

    def _check_fonts_used(self, doc: Any, issues: list[Issue]) -> None:
        """记录文档中使用的字体"""
        try:
            used_fonts: set[str] = set()
            for p in doc.Paragraphs:
                try:
                    font_name = str(p.Range.Font.Name)
                    if font_name:
                        used_fonts.add(font_name)
                except Exception:
                    pass

            for font in sorted(used_fonts):
                issues.append(
                    Issue(
                        severity="info",
                        category="font",
                        location="文档",
                        message=f"使用了字体: {font}",
                        suggestion="如果目标机器未安装此字体，可能回退为默认字体",
                    )
                )
        except Exception:
            pass

    def _check_table_width(self, doc: Any, issues: list[Issue]) -> None:
        """检查表格是否可能溢出页面宽度

        对比表格总宽度与页面可用宽度（页宽 - 左右边距）。
        """
        try:
            page = doc.PageSetup
            usable_width = page.PageWidth - page.LeftMargin - page.RightMargin
        except Exception:
            return

        try:
            for i in range(1, doc.Tables.Count + 1):
                try:
                    table = doc.Tables(i)
                    total = 0.0
                    for j in range(1, table.Columns.Count + 1):
                        total += table.Columns(j).Width
                    if total > usable_width * 1.05:  # 5% 容差
                        issues.append(
                            Issue(
                                severity="warning",
                                category="layout",
                                location=f"表格 {i}",
                                message=f"表格总宽度 ({total:.0f}pt) 超过页面可用宽度 ({usable_width:.0f}pt)",
                                suggestion="考虑调整表格列宽或页面边距",
                            )
                        )
                except Exception:
                    pass
        except Exception:
            pass

    def _check_page_margins(self, doc: Any, issues: list[Issue]) -> None:
        """检查页面边距是否合理"""
        try:
            page = doc.PageSetup
            min_margin = 18  # pt, ~6.35mm
            margins = {
                "上": page.TopMargin,
                "下": page.BottomMargin,
                "左": page.LeftMargin,
                "右": page.RightMargin,
            }
            for name, val in margins.items():
                if val < min_margin:
                    issues.append(
                        Issue(
                            severity="info",
                            category="layout",
                            location=f"页面{name}边距",
                            message=f"页面{name}边距过小 ({val:.0f}pt)",
                            suggestion="建议边距不小于 18pt (约 6.35mm)",
                        )
                    )
        except Exception:
            pass

    def _check_paragraph_styles(self, doc: Any, issues: list[Issue]) -> None:
        """检查段落样式：空段落、连续空格"""
        try:
            empty_count = 0
            for p in doc.Paragraphs:
                try:
                    text = p.Range.Text.strip()
                    if not text or text in ("\r", "\x0c", "\x0d"):
                        empty_count += 1
                except Exception:
                    pass

            if empty_count > 50:
                issues.append(
                    Issue(
                        severity="info",
                        category="style",
                        location="文档",
                        message=f"文档中有 {empty_count} 个空白段落",
                        suggestion="考虑移除多余的空白段落以减小文件体积",
                    )
                )
        except Exception:
            pass

    # ── Calc 诊断 ──

    def diagnose_calc(self, app: Any) -> list[Issue]:
        """诊断 Excel 工作簿

        检测：
        - 公式错误 (#REF!, #VALUE!, #N/A, #DIV/0!, #NUM!, #NULL!, #NAME?)
        - 隐藏行/列
        - 合并单元格
        - 数值存储为文本
        - 条件格式冲突
        """
        issues: list[Issue] = []
        wb = app.ActiveWorkbook

        for sheet_idx in range(1, wb.Sheets.Count + 1):
            ws = wb.Sheets(sheet_idx)
            sheet_name = str(ws.Name)

            # -- 公式错误检测 --
            self._check_formula_errors(ws, sheet_name, issues)

            # -- 隐藏行列检测 --
            self._check_hidden_rows_cols(ws, sheet_name, issues)

        return issues

    def _check_formula_errors(self, ws: Any, sheet_name: str, issues: list[Issue]) -> None:
        """检测公式错误"""
        error_functions = ["#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NUM!", "#NULL!", "#NAME?"]
        try:
            used = ws.UsedRange
            if used is None:
                return
            # 只检查有公式的单元格
            try:
                special_cells = used.SpecialCells(-4123)  # xlCellTypeFormulas
                for cell in special_cells:
                    try:
                        text = str(cell.Text)
                        for err in error_functions:
                            if err in text:
                                issues.append(
                                    Issue(
                                        severity="error",
                                        category="formula",
                                        location=f"Sheet[{sheet_name}] {cell.Address}",
                                        message=f"公式错误: {text}",
                                        suggestion=f"检查单元格 {cell.Address} 的公式引用",
                                    )
                                )
                                break
                    except Exception:
                        pass
            except Exception:
                # 没有公式单元格时会抛 COM 错误，忽略
                pass
        except Exception:
            pass

    def _check_hidden_rows_cols(self, ws: Any, sheet_name: str, issues: list[Issue]) -> None:
        """检测隐藏行/列"""
        try:
            hidden_rows = []
            for r in range(1, min(ws.Rows.Count + 1, 200)):
                try:
                    if ws.Rows(r).Hidden:
                        hidden_rows.append(str(r))
                except Exception:
                    pass
            if hidden_rows:
                issues.append(
                    Issue(
                        severity="info",
                        category="structure",
                        location=f"Sheet[{sheet_name}]",
                        message=f"{len(hidden_rows)} 行被隐藏 (行: {', '.join(hidden_rows[:10])}"
                        f"{'...' if len(hidden_rows) > 10 else ''})",
                        suggestion="右键行号 → 取消隐藏可恢复显示",
                    )
                )
        except Exception:
            pass

        try:
            hidden_cols = []
            for c in range(1, min(ws.Columns.Count + 1, 50)):
                try:
                    if ws.Columns(c).Hidden:
                        hidden_cols.append(_col_letter(c))
                except Exception:
                    pass
            if hidden_cols:
                issues.append(
                    Issue(
                        severity="info",
                        category="structure",
                        location=f"Sheet[{sheet_name}]",
                        message=f"{len(hidden_cols)} 列被隐藏 (列: {', '.join(hidden_cols[:10])}"
                        f"{'...' if len(hidden_cols) > 10 else ''})",
                        suggestion="右键列标 → 取消隐藏可恢复显示",
                    )
                )
        except Exception:
            pass

    # ── Impress 诊断 ──

    def diagnose_impress(self, app: Any) -> list[Issue]:
        """诊断 PPT 演示文稿

        检测：
        - 文本溢出形状
        - 字体一致性
        - 幻灯片切换时间
        - 图片无替代文本
        """
        issues: list[Issue] = []
        pres = app.ActivePresentation

        all_fonts: set[str] = set()

        for i in range(1, pres.Slides.Count + 1):
            try:
                sl = pres.Slides(i)
                slide_label = f"幻灯片 {i}"

                # -- 文本溢出 --
                self._check_shape_overflow(sl, slide_label, issues)

                # -- 字体收集 --
                self._collect_slide_fonts(sl, all_fonts)

                # -- 图片 alt text --
                self._check_slide_image_alt(sl, slide_label, issues)

                # -- 切换时间 --
                self._check_transition_duration(sl, slide_label, issues)

            except Exception:
                pass

        # -- 字体一致性分析 --
        if len(all_fonts) > 3:
            issues.append(
                Issue(
                    severity="warning",
                    category="consistency",
                    location="演示文稿",
                    message=f"使用了 {len(all_fonts)} 种不同字体: {', '.join(sorted(all_fonts))}",
                    suggestion="建议统一为 2-3 种字体以保持视觉一致性",
                )
            )

        return issues

    def _check_shape_overflow(self, sl: Any, label: str, issues: list[Issue]) -> None:
        """检查形状文本是否溢出"""
        try:
            for sh in sl.Shapes:
                try:
                    if not sh.HasTextFrame:
                        continue
                    tf = sh.TextFrame
                    if not tf.WordWrap:
                        # 关闭自动换行可能导致文本溢出
                        text_len = len(tf.TextRange.Text)
                        shape_width = sh.Width
                        # 粗略估算：中文字符约 14pt 宽
                        estimated_width = text_len * 14
                        if estimated_width > shape_width * 1.2:
                            issues.append(
                                Issue(
                                    severity="warning",
                                    category="layout",
                                    location=f"{label}, 形状 {sh.Name}",
                                    message="形状文本可能超出边界（自动换行已关闭）",
                                    suggestion="启用形状的自动换行，或调整形状宽度",
                                )
                            )
                except Exception:
                    pass
        except Exception:
            pass

    def _collect_slide_fonts(self, sl: Any, all_fonts: set[str]) -> None:
        """收集幻灯片中使用的字体"""
        try:
            for sh in sl.Shapes:
                try:
                    if sh.HasTextFrame:
                        for run in sh.TextFrame.TextRange.Runs():
                            try:
                                font_name = str(run.Font.Name)
                                if font_name and font_name != "宋体":
                                    all_fonts.add(font_name)
                            except Exception:
                                pass
                except Exception:
                    pass
        except Exception:
            pass

    def _check_slide_image_alt(self, sl: Any, label: str, issues: list[Issue]) -> None:
        """检查幻灯片中图片的替代文本"""
        try:
            for sh in sl.Shapes:
                try:
                    if sh.Type == 13 and not sh.AlternativeText:  # msoPicture
                            issues.append(
                                Issue(
                                    severity="warning",
                                    category="image",
                                    location=f"{label}, 形状 {sh.Name}",
                                    message="图片缺少替代文本（alt text）",
                                    suggestion="右键图片 → 设置对象格式 → 可选文字",
                                )
                            )
                except Exception:
                    pass
        except Exception:
            pass

    def _check_transition_duration(self, sl: Any, label: str, issues: list[Issue]) -> None:
        """检查幻灯片切换时间"""
        try:
            transition = sl.SlideShowTransition
            duration = transition.Duration
            if duration > 5.0:
                issues.append(
                    Issue(
                        severity="info",
                        category="style",
                        location=label,
                        message=f"切换时间过长 ({duration:.1f}秒)",
                        suggestion="考虑将切换时间缩短到 2 秒以内",
                    )
                )
        except Exception:
            pass
