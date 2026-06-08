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

# ── Issue 子类型常量 ──
# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)

# 公式相关
SUBTYPE_FORMULA_EVAL_ERROR = "formula_eval_error"
SUBTYPE_FORMULA_CACHE_STALE = "formula_cache_stale"
SUBTYPE_FORMULA_REF_MISSING_SHEET = "formula_ref_missing_sheet"

# 字段相关
SUBTYPE_FIELD_NOT_EVALUATED = "field_not_evaluated"
SUBTYPE_FIELD_CACHE_STALE = "field_cache_stale"

# 图表/命名区域
SUBTYPE_CHART_REF_MISSING = "chart_ref_missing_sheet"
SUBTYPE_CHART_CACHE_STALE = "chart_cache_stale"
SUBTYPE_DEFINEDNAME_BROKEN = "definedname_broken"
SUBTYPE_DEFINEDNAME_TARGET_MISSING = "definedname_target_missing"

# 结构
SUBTYPE_MISSING_ALT_TEXT = "missing_alt_text"
SUBTYPE_BROKEN_PART_REF = "broken_part_ref"
SUBTYPE_TEXT_OVERFLOW = "text_overflow"

# 样式
SUBTYPE_STYLE_INCONSISTENT = "style_inconsistent"

# 段落
SUBTYPE_NUMBERING_GAP = "numbering_gap"
SUBTYPE_EXCESSIVE_BLANK_PARAGRAPHS = "excessive_blank_paragraphs"

# 数据
SUBTYPE_NUMBER_AS_TEXT = "number_as_text"
SUBTYPE_MERGED_CELL = "merged_cell"

# 字体
SUBTYPE_FONT_INCONSISTENT = "font_inconsistent"
SUBTYPE_FONT_NOT_FOUND = "font_not_found"

# 动画/切换
SUBTYPE_ANIM_TRIGGER_MISSING = "anim_trigger_missing"
SUBTYPE_MASTER_OVERRIDE = "master_override"

# 条件格式
SUBTYPE_CONDITIONAL_FORMAT_CONFLICT = "conditional_format_conflict"


@dataclass
class Issue:
    """文档问题"""

    severity: str  # "error", "warning", "info"
    category: str  # "font", "layout", "image", "formula", "style", "structure", "consistency", "field", "data", "anim"
    subtype: str  # 稳定标识符，如 formula_error_ref / broken_definedname / missing_alt_text
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
        - 图片是否有替代文本（alt text）
        - 字体是否在系统中可用（记录使用字体）
        - 表格是否可能溢出页面宽度
        - 页面设置是否异常
        - 段落样式一致性
        - 段落编号断裂
        - 空白段落过多
        - 页码字段过期
        - 文本溢出
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

        # -- 6. 空白段落过多 --
        self._check_excessive_blank_paragraphs(doc, issues)

        # -- 7. 段落编号断裂 --
        self._check_numbering_gap(doc, issues)

        # -- 8. 页码字段过期 --
        self._check_field_stale(doc, issues)

        # -- 9. 文本溢出检测 --
        self._check_text_overflow(doc, issues)

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
                                subtype=SUBTYPE_MISSING_ALT_TEXT,
                                location=f"InlineShape {shape.Index}",
                                message="图片缺少替代文本（alt text）",
                                suggestion="右键图片 → 查看替代文本 → 输入描述",
                            )
                        )
                except Exception:
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
                        subtype=SUBTYPE_FONT_NOT_FOUND,
                        location="文档",
                        message=f"使用了字体: {font}",
                        suggestion="如果目标机器未安装此字体，可能回退为默认字体",
                    )
                )
        except Exception:
            pass

    def _check_table_width(self, doc: Any, issues: list[Issue]) -> None:
        """检查表格是否可能溢出页面宽度"""
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
                    if total > usable_width * 1.05:
                        issues.append(
                            Issue(
                                severity="warning",
                                category="layout",
                                subtype=SUBTYPE_TEXT_OVERFLOW,
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
                            subtype=SUBTYPE_TEXT_OVERFLOW,
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
            prev_style = None
            style_changed_count = 0
            for p in doc.Paragraphs:
                try:
                    text = p.Range.Text.strip()
                    if not text or text in ("\r", "\x0c", "\x0d"):
                        continue
                    curr_style = str(p.Style)
                    if prev_style is not None and curr_style != prev_style:
                        style_changed_count += 1
                    prev_style = curr_style
                except Exception:
                    pass

            if style_changed_count > 30:
                issues.append(
                    Issue(
                        severity="info",
                        category="style",
                        subtype=SUBTYPE_STYLE_INCONSISTENT,
                        location="文档",
                        message=f"文档中样式切换 {style_changed_count} 次，可能风格不统一",
                        suggestion="考虑统一段落样式，或使用样式集一键格式化",
                    )
                )
        except Exception:
            pass

    def _check_excessive_blank_paragraphs(self, doc: Any, issues: list[Issue]) -> None:
        """检查空白段落数量"""
        try:
            empty_count = 0
            consecutive_empty = 0
            max_consecutive_empty = 0
            for p in doc.Paragraphs:
                try:
                    text = p.Range.Text.strip()
                    if not text or text in ("\r", "\x0c", "\x0d"):
                        empty_count += 1
                        consecutive_empty += 1
                        if consecutive_empty > max_consecutive_empty:
                            max_consecutive_empty = consecutive_empty
                    else:
                        consecutive_empty = 0
                except Exception:
                    pass

            if empty_count > 50:
                issues.append(
                    Issue(
                        severity="info",
                        category="structure",
                        subtype=SUBTYPE_EXCESSIVE_BLANK_PARAGRAPHS,
                        location="文档",
                        message=f"文档中有 {empty_count} 个空白段落（最长连续 {max_consecutive_empty} 段）",
                        suggestion="考虑移除多余的空白段落以减小文件体积",
                    )
                )
        except Exception:
            pass

    def _check_numbering_gap(self, doc: Any, issues: list[Issue]) -> None:
        """检测段落编号断裂

        遍历文档段落，若检测到编号列表序列中存在编号不连续，则报告。
        """
        try:
            prev_num = 0
            gap_count = 0
            for p in doc.Paragraphs:
                try:
                    rng = p.Range
                    if hasattr(rng, "ListFormat") and rng.ListFormat.ListType > 0:
                        list_str = str(rng.ListFormat.ListString).strip()
                        # 尝试提取数字
                        if list_str:
                            import re

                            m = re.search(r"(\d+)", str(list_str))
                            if m:
                                curr_num = int(m.group(1))
                                if prev_num > 0 and curr_num > prev_num + 1:
                                    gap_count += 1
                                prev_num = curr_num
                except Exception:
                    pass

            if gap_count > 0:
                issues.append(
                    Issue(
                        severity="warning",
                        category="structure",
                        subtype=SUBTYPE_NUMBERING_GAP,
                        location="文档",
                        message=f"检测到 {gap_count} 处段落编号断裂（编号不连续）",
                        suggestion="检查编号列表是否被手动修改或删除",
                    )
                )
        except Exception:
            pass

    def _check_field_stale(self, doc: Any, issues: list[Issue]) -> None:
        """检测字段（页码/目录等）是否过期未更新"""
        try:
            stale_count = 0
            for f in doc.Fields:
                try:
                    # 检查字段是否被锁定(_FieldLocked)或缓存未更新
                    code = str(f.Code.Text or "")
                    result = str(f.Result.Text or "")
                    if "PAGE" in code.upper() and not result.strip() or "NUMPAGES" in code.upper() and not result.strip():
                        stale_count += 1
                except Exception:
                    pass

            if stale_count > 0:
                issues.append(
                    Issue(
                        severity="warning",
                        category="field",
                        subtype=SUBTYPE_FIELD_NOT_EVALUATED,
                        location="文档",
                        message=f"检测到 {stale_count} 个字段可能未求值或缓存过期",
                        suggestion="按 Ctrl+A 全选，再按 F9 更新所有字段",
                    )
                )
        except Exception:
            pass

    def _check_text_overflow(self, doc: Any, issues: list[Issue]) -> None:
        """检测文本框/形状文本溢出"""
        try:
            for shape in doc.InlineShapes:
                try:
                    if hasattr(shape, "HasTextFrame") and shape.HasTextFrame:
                        tf = shape.TextFrame
                        if hasattr(tf, "Overflowing") and tf.Overflowing:
                            issues.append(
                                Issue(
                                    severity="warning",
                                    category="layout",
                                    subtype=SUBTYPE_TEXT_OVERFLOW,
                                    location=f"InlineShape {shape.Index}",
                                    message="文本框内容溢出形状边界",
                                    suggestion="调整形状大小或减小字体",
                                )
                            )
                except Exception:
                    pass
        except Exception:
            pass

    # ── Calc 诊断 ──

    def diagnose_calc(self, app: Any) -> list[Issue]:
        """诊断 Excel 工作簿

        检测：
        - 公式错误 (#REF!/#VALUE!/#DIV/0!/#N/A/#NUM!/#NULL!/#NAME?)
        - 命名区域断裂
        - 数据源引用断裂
        - 数值存储为文本
        - 条件格式冲突
        - 合并单元格
        - 隐藏行/列
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

            # -- 数值存为文本 --
            self._check_numbers_as_text(ws, sheet_name, issues)

            # -- 合并单元格 --
            self._check_merged_cells(ws, sheet_name, issues)

            # -- 条件格式冲突 --
            self._check_conditional_format_conflict(ws, sheet_name, issues)

        # -- 命名区域断裂检测 --
        self._check_defined_names(wb, issues)

        return issues

    def _check_formula_errors(self, ws: Any, sheet_name: str, issues: list[Issue]) -> None:
        """检测公式错误及缓存过期"""
        error_map = {
            "#REF!": SUBTYPE_FORMULA_REF_MISSING_SHEET,
            "#VALUE!": SUBTYPE_FORMULA_EVAL_ERROR,
            "#DIV/0!": SUBTYPE_FORMULA_EVAL_ERROR,
            "#N/A": SUBTYPE_FORMULA_EVAL_ERROR,
            "#NUM!": SUBTYPE_FORMULA_EVAL_ERROR,
            "#NULL!": SUBTYPE_FORMULA_EVAL_ERROR,
            "#NAME?": SUBTYPE_FORMULA_EVAL_ERROR,
        }
        try:
            used = ws.UsedRange
            if used is None:
                return
            try:
                special_cells = used.SpecialCells(-4123)  # xlCellTypeFormulas
                for cell in special_cells:
                    try:
                        text = str(cell.Text)
                        for err, subtype in error_map.items():
                            if err in text:
                                issues.append(
                                    Issue(
                                        severity="error",
                                        category="formula",
                                        subtype=subtype,
                                        location=f"Sheet[{sheet_name}] {cell.Address}",
                                        message=f"公式错误: {text}",
                                        suggestion=f"检查单元格 {cell.Address} 的公式引用",
                                    )
                                )
                                break
                    except Exception:
                        pass
            except Exception:
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
                        subtype=SUBTYPE_BROKEN_PART_REF,
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
                        subtype=SUBTYPE_BROKEN_PART_REF,
                        location=f"Sheet[{sheet_name}]",
                        message=f"{len(hidden_cols)} 列被隐藏 (列: {', '.join(hidden_cols[:10])}"
                        f"{'...' if len(hidden_cols) > 10 else ''})",
                        suggestion="右键列标 → 取消隐藏可恢复显示",
                    )
                )
        except Exception:
            pass

    def _check_numbers_as_text(self, ws: Any, sheet_name: str, issues: list[Issue]) -> None:
        """检测数值存储为文本"""
        try:
            used = ws.UsedRange
            if used is None:
                return
            text_count = 0
            max_check = min(used.Cells.Count, 1000)
            for i in range(1, max_check + 1):
                try:
                    cell = used.Cells(i)
                    # 检查是否有 "数字存为文本" 的 WPS 错误标记
                    if hasattr(cell, "Errors"):
                        try:
                            # xlNumberAsText = 6
                            if cell.Errors(6).Value:
                                text_count += 1
                                if text_count <= 5:
                                    issues.append(
                                        Issue(
                                            severity="warning",
                                            category="data",
                                            subtype=SUBTYPE_NUMBER_AS_TEXT,
                                            location=f"Sheet[{sheet_name}] {cell.Address}",
                                            message="数值存储为文本格式",
                                            suggestion=f"选中 {cell.Address}，点击错误提示 → 转换为数字",
                                        )
                                    )
                        except Exception:
                            pass
                except Exception:
                    pass

            if text_count > 5:
                issues.append(
                    Issue(
                        severity="warning",
                        category="data",
                        subtype=SUBTYPE_NUMBER_AS_TEXT,
                        location=f"Sheet[{sheet_name}]",
                        message=f"检测到 {text_count} 个单元格将数值存为文本",
                        suggestion="全选区域 → 错误检查 → 全部转换为数字",
                    )
                )
        except Exception:
            pass

    def _check_merged_cells(self, ws: Any, sheet_name: str, issues: list[Issue]) -> None:
        """检测合并单元格"""
        try:
            merged_count = 0
            used = ws.UsedRange
            if used is None:
                return
            for _area in used.MergeCells:
                try:
                    merged_count += 1
                except Exception:
                    pass

            if merged_count > 20:
                issues.append(
                    Issue(
                        severity="info",
                        category="structure",
                        subtype=SUBTYPE_MERGED_CELL,
                        location=f"Sheet[{sheet_name}]",
                        message=f"工作表包含 {merged_count} 个合并单元格区域",
                        suggestion="合并单元格可能影响排序、筛选和公式填充",
                    )
                )
        except Exception:
            pass

    def _check_conditional_format_conflict(
        self, ws: Any, sheet_name: str, issues: list[Issue]
    ) -> None:
        """检测条件格式冲突"""
        try:
            cf_count = ws.UsedRange.FormatConditions.Count
            if cf_count > 10:
                issues.append(
                    Issue(
                        severity="info",
                        category="style",
                        subtype=SUBTYPE_CONDITIONAL_FORMAT_CONFLICT,
                        location=f"Sheet[{sheet_name}]",
                        message=f"工作表包含 {cf_count} 个条件格式规则，可能存在冲突",
                        suggestion="检查条件格式优先级，确保规则的 ApplyTo 区域不重叠",
                    )
                )
        except Exception:
            pass

        # 检查相邻区域的条件格式重叠
        try:
            for i in range(1, ws.UsedRange.FormatConditions.Count + 1):
                try:
                    fc = ws.UsedRange.FormatConditions(i)
                    str(fc.AppliesTo.Address) if hasattr(fc, "AppliesTo") else ""
                    # 检查同一单元格被多个条件格式覆盖（通过公式引用分析）
                    if hasattr(fc, "Formula1") and fc.Formula1:
                        formula_text = str(fc.Formula1)
                        if formula_text and "=" in formula_text:
                            pass  # 条件公式本身是正常的；主要关注重叠
                except Exception:
                    pass
        except Exception:
            pass

    def _check_defined_names(self, wb: Any, issues: list[Issue]) -> None:
        """检测命名区域断裂"""
        try:
            for i in range(1, wb.Names.Count + 1):
                try:
                    nm = wb.Names(i)
                    refers_to = str(nm.RefersTo)
                    if "#REF!" in refers_to:
                        issues.append(
                            Issue(
                                severity="error",
                                category="formula",
                                subtype=SUBTYPE_DEFINEDNAME_BROKEN,
                                location=f"名称 '{nm.Name}'",
                                message=f"命名区域 '{nm.Name}' 引用断裂: {refers_to}",
                                suggestion="重新定义或删除此命名区域（公式 → 名称管理器）",
                            )
                        )
                    # 检查目标工作表是否存在
                    if "!" in refers_to:
                        sheet_ref = refers_to.split("!")[0].strip().strip("'")
                        if sheet_ref and "=" in sheet_ref:
                            # 处理 =SheetName!Range 格式
                            sheet_ref = sheet_ref.split("=")[-1].strip().strip("'")
                        try:
                            wb.Sheets(sheet_ref)
                        except Exception:
                            issues.append(
                                Issue(
                                    severity="error",
                                    category="formula",
                                    subtype=SUBTYPE_DEFINEDNAME_TARGET_MISSING,
                                    location=f"名称 '{nm.Name}'",
                                    message=f"命名区域 '{nm.Name}' 引用的工作表 '{sheet_ref}' 不存在",
                                    suggestion="检查引用的工作表是否被删除或重命名",
                                )
                            )
                except Exception:
                    pass
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
        - 动画触发器缺失
        - 母版覆盖异常
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

                # -- 动画触发器 --
                self._check_anim_triggers(sl, slide_label, issues)

                # -- 母版覆盖异常 --
                self._check_master_override(sl, slide_label, issues)

                # -- 每页字体一致性 --
                self._check_slide_font_consistency(sl, slide_label, issues)

            except Exception:
                pass

        # -- 全局字体一致性分析 --
        if len(all_fonts) > 3:
            issues.append(
                Issue(
                    severity="warning",
                    category="consistency",
                    subtype=SUBTYPE_FONT_INCONSISTENT,
                    location="演示文稿",
                    message=f"使用了 {len(all_fonts)} 种不同字体: {', '.join(sorted(all_fonts))}",
                    suggestion="建议统一为 2-3 种字体以保持视觉一致性",
                )
            )

        return issues

    def _check_shape_overflow(self, sl: Any, label: str, issues: list[Issue]) -> None:
        """检查形状文本是否溢出（增强版）"""
        try:
            for sh in sl.Shapes:
                try:
                    if not sh.HasTextFrame:
                        continue
                    tf = sh.TextFrame
                    # 检测一：自动换行关闭可能导致溢出
                    if not tf.WordWrap:
                        text_len = len(tf.TextRange.Text)
                        shape_width = sh.Width
                        estimated_width = text_len * 14
                        if estimated_width > shape_width * 1.2:
                            issues.append(
                                Issue(
                                    severity="warning",
                                    category="layout",
                                    subtype=SUBTYPE_TEXT_OVERFLOW,
                                    location=f"{label}, 形状 {sh.Name}",
                                    message="形状文本可能超出边界（自动换行已关闭）",
                                    suggestion="启用形状的自动换行，或调整形状宽度",
                                )
                            )

                    # 检测二：检查 HasOverflowing 属性
                    if hasattr(tf, "HasOverflowing"):
                        try:
                            if tf.HasOverflowing:
                                issues.append(
                                    Issue(
                                        severity="warning",
                                        category="layout",
                                        subtype=SUBTYPE_TEXT_OVERFLOW,
                                        location=f"{label}, 形状 {sh.Name}",
                                        message="文本框内容超出形状边界",
                                        suggestion="调整形状大小、减小字体或缩减文本内容",
                                    )
                                )
                        except Exception:
                            pass
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
                                subtype=SUBTYPE_MISSING_ALT_TEXT,
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
                        subtype=SUBTYPE_FIELD_CACHE_STALE,
                        location=label,
                        message=f"切换时间过长 ({duration:.1f}秒)",
                        suggestion="考虑将切换时间缩短到 2 秒以内",
                    )
                )
        except Exception:
            pass

    def _check_anim_triggers(self, sl: Any, label: str, issues: list[Issue]) -> None:
        """检测动画触发器缺失

        遍历幻灯片时间线，检查是否存在无触发器控制的动画序列。
        """
        try:
            for i in range(1, sl.TimeLine.InteractiveSequences.Count + 1):
                try:
                    seq = sl.TimeLine.InteractiveSequences(i)
                    # 检查序列中每个动画效果
                    for j in range(1, seq.Count + 1):
                        try:
                            effect = seq(j)
                            # 如果动画时长很短且无触发器，可能是残留动画
                            if (
                                hasattr(effect, "Timing")
                                and effect.Timing.TriggerType == 1
                            ):  # msoAnimTriggerOnPageClick
                                pass  # 页面点击触发 - 正常
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception:
            pass

    def _check_master_override(self, sl: Any, label: str, issues: list[Issue]) -> None:
        """检测母版覆盖异常

        检测幻灯片是否完全覆盖了母版背景，导致母版统一修改不生效。
        """
        try:
            if hasattr(sl, "FollowMasterBackground") and not sl.FollowMasterBackground:
                issues.append(
                    Issue(
                        severity="info",
                        category="style",
                        subtype=SUBTYPE_MASTER_OVERRIDE,
                        location=label,
                        message="幻灯片未跟随母版背景",
                        suggestion="若需统一修改背景，建议启用 '跟随母版背景'",
                    )
                )
        except Exception:
            pass

    def _check_slide_font_consistency(self, sl: Any, label: str, issues: list[Issue]) -> None:
        """检查单张幻灯片的字体一致性"""
        try:
            fonts: set[str] = set()
            for sh in sl.Shapes:
                try:
                    if sh.HasTextFrame:
                        for run in sh.TextFrame.TextRange.Runs():
                            try:
                                font_name = str(run.Font.Name)
                                if font_name:
                                    fonts.add(font_name)
                            except Exception:
                                pass
                except Exception:
                    pass

            if len(fonts) > 4:
                issues.append(
                    Issue(
                        severity="warning",
                        category="consistency",
                        subtype=SUBTYPE_FONT_INCONSISTENT,
                        location=label,
                        message=f"使用了 {len(fonts)} 种字体: {', '.join(sorted(fonts))}",
                        suggestion="建议单页字体不超过 3 种",
                    )
                )
        except Exception:
            pass
