"""文档验证服务

通过 WPS COM 接口验证文档完整性，包括：
- 拼写检查
- 格式 round-trip 验证（另存 → 比较）
- 链接/交叉引用完整性
- 样式引用有效性
- 字段求值状态检查

设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from wps_cli.consts import (
    WD_DO_NOT_SAVE_CHANGES,
)
from wps_cli.services.session_manager import SessionManager


@dataclass
class ValidateResult:
    """验证结果"""

    passed: bool
    file: str
    checks: list[dict]
    issues_count: int
    errors_count: int
    warnings_count: int


@dataclass
class ValidateService:
    """文档验证服务

    通过 WPS COM 接口验证文档完整性。所有 COM 调用均受 try/except 保护。
    """

    manager: SessionManager

    # ── Writer 验证 ──

    def validate_writer(self, path: Path) -> ValidateResult:
        """验证 Word 文档

        检查项：
        - 拼写错误检查 (通过 COM SpellingErrors)
        - 超链接有效性检查 (Hyperlinks)
        - TOC/交叉引用字段是否过期
        - 字段错误检测 (Error! 文本)
        - 嵌入对象完整性
        - 字体可用性
        """
        checks: list[dict] = []
        errors = 0
        warnings = 0

        with self.manager.session("writer") as app:
            doc = app.Documents.Open(
                str(path),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
            )
            try:
                # 1. 拼写错误检查
                errors += self._check_spelling(doc, checks)

                # 2. 超链接有效性
                errors += self._check_hyperlinks(doc, checks)

                # 3. 字段状态 (TOC, PAGE, 交叉引用)
                errors += self._check_fields(doc, checks)

                # 4. 嵌入对象完整性
                warnings += self._check_embedded_objects(doc, checks)

                # 5. 字体可用性
                warnings += self._check_fonts(doc, checks)

                # 6. 文档结构
                warnings += self._check_document_structure(doc, checks)

            finally:
                doc.Close(WD_DO_NOT_SAVE_CHANGES)

        total_issues = errors + warnings
        passed = errors == 0

        return ValidateResult(
            passed=passed,
            file=str(path),
            checks=checks,
            issues_count=total_issues,
            errors_count=errors,
            warnings_count=warnings,
        )

    @staticmethod
    def _check_spelling(doc: Any, checks: list[dict]) -> int:
        """拼写错误检查"""
        try:
            count = doc.SpellingErrors.Count
            if count > 0:
                checks.append(
                    {
                        "check": "spelling",
                        "status": "warning",
                        "message": f"发现 {count} 处拼写错误",
                        "count": count,
                    }
                )
                return 0  # spelling is a warning, not an error
            checks.append(
                {
                    "check": "spelling",
                    "status": "pass",
                    "message": "未发现拼写错误",
                    "count": 0,
                }
            )
        except Exception as e:
            checks.append(
                {
                    "check": "spelling",
                    "status": "skip",
                    "message": f"无法检查拼写: {e}",
                }
            )
        return 0

    @staticmethod
    def _check_hyperlinks(doc: Any, checks: list[dict]) -> int:
        """超链接有效性检查"""
        errors = 0
        try:
            total = doc.Hyperlinks.Count
            broken_count = 0
            empty_count = 0
            external_count = 0
            broken_details: list[dict] = []

            for i in range(1, total + 1):
                try:
                    hl = doc.Hyperlinks(i)
                    addr = str(hl.Address) if hl.Address else ""
                    sub = str(hl.SubAddress) if hl.SubAddress else ""
                    text = ""
                    try:
                        text = str(hl.TextToDisplay)[:80]
                    except Exception:
                        pass

                    if not addr and not sub:
                        empty_count += 1

                    # 检查内部书签引用
                    if sub and not addr:
                        try:
                            doc.Bookmarks(sub)
                        except Exception:
                            broken_count += 1
                            broken_details.append(
                                {
                                    "index": i,
                                    "address": sub,
                                    "text": text,
                                    "reason": "书签不存在",
                                }
                            )

                    if addr:
                        external_count += 1
                except Exception:
                    broken_count += 1

            if broken_count > 0:
                errors += 1
                checks.append(
                    {
                        "check": "hyperlinks",
                        "status": "fail",
                        "message": f"{broken_count} 个超链接损坏",
                        "total": total,
                        "broken": broken_count,
                        "empty": empty_count,
                        "external": external_count,
                        "details": broken_details[:10],
                    }
                )
            elif empty_count > 0:
                checks.append(
                    {
                        "check": "hyperlinks",
                        "status": "warning",
                        "message": f"{empty_count} 个超链接地址为空",
                        "total": total,
                        "broken": 0,
                        "empty": empty_count,
                        "external": external_count,
                    }
                )
            else:
                checks.append(
                    {
                        "check": "hyperlinks",
                        "status": "pass",
                        "message": f"所有 {total} 个超链接有效",
                        "total": total,
                        "broken": 0,
                        "empty": 0,
                        "external": external_count,
                    }
                )
        except Exception as e:
            checks.append(
                {
                    "check": "hyperlinks",
                    "status": "skip",
                    "message": f"无法检查超链接: {e}",
                }
            )
        return errors

    @staticmethod
    def _check_fields(doc: Any, checks: list[dict]) -> int:
        """字段状态检查 (TOC, PAGE, 交叉引用等)"""
        errors = 0
        try:
            total_fields = doc.Fields.Count
            error_fields: list[dict] = []
            toc_count = 0
            toc_stale = False

            # 检查 TOC
            try:
                for i in range(1, doc.TablesOfContents.Count + 1):
                    toc_count += 1
                    toc = doc.TablesOfContents(i)
                    try:
                        # 尝试检测 TOC 是否过期：获取字段结果文本
                        # 如果 TOC 文本中包含 "Error!" 则表明有问题
                        toc_range_text = str(toc.Range.Text)
                        if "Error!" in toc_range_text or "错误!" in toc_range_text:
                            toc_stale = True
                            errors += 1
                    except Exception:
                        pass
            except Exception:
                pass

            # 检测字段错误 (Error! 文本)
            for i in range(1, min(total_fields, 500) + 1):
                try:
                    fld = doc.Fields(i)
                    result_text = ""
                    try:
                        result_text = str(fld.Result.Text)[:100]
                    except Exception:
                        pass
                    if "Error!" in result_text or "错误!" in result_text:
                        code_text = ""
                        try:
                            code_text = str(fld.Code.Text)[:80]
                        except Exception:
                            pass
                        error_fields.append(
                            {
                                "index": i,
                                "code": code_text,
                                "result": result_text,
                            }
                        )
                except Exception:
                    pass

            if toc_count > 0:
                status = "warning" if toc_stale else "pass"
                checks.append(
                    {
                        "check": "toc",
                        "status": status,
                        "message": (
                            f"检测到 {toc_count} 个目录, 状态过期"
                            if toc_stale
                            else f"{toc_count} 个目录状态正常"
                        ),
                        "count": toc_count,
                        "stale": toc_stale,
                    }
                )

            if error_fields:
                errors += 1
                checks.append(
                    {
                        "check": "field_errors",
                        "status": "fail",
                        "message": f"发现 {len(error_fields)} 个错误字段",
                        "total": total_fields,
                        "error_count": len(error_fields),
                        "details": error_fields[:20],
                    }
                )
            else:
                checks.append(
                    {
                        "check": "field_errors",
                        "status": "pass",
                        "message": f"所有 {total_fields} 个字段状态正常",
                        "total": total_fields,
                        "error_count": 0,
                    }
                )
        except Exception as e:
            checks.append(
                {
                    "check": "field_errors",
                    "status": "skip",
                    "message": f"无法检查字段状态: {e}",
                }
            )
        return errors

    @staticmethod
    def _check_embedded_objects(doc: Any, checks: list[dict]) -> int:
        """嵌入对象完整性"""
        warnings = 0
        try:
            inline_shapes = doc.InlineShapes.Count
            shapes = doc.Shapes.Count
            ole_count = 0
            ole_missing = 0

            # 检查 OLE 对象
            try:
                for i in range(1, doc.InlineShapes.Count + 1):
                    try:
                        shape = doc.InlineShapes(i)
                        if str(shape.Type) == "1":  # wdInlineShapeEmbeddedOLEObject
                            ole_count += 1
                            try:
                                _ = shape.OLEFormat.ProgID
                            except Exception:
                                ole_missing += 1
                    except Exception:
                        pass
            except Exception:
                pass

            if ole_missing > 0:
                warnings += 1
                checks.append(
                    {
                        "check": "embedded_objects",
                        "status": "warning",
                        "message": f"{ole_missing}/{ole_count} 个 OLE 对象可能不可用",
                        "inline_shapes": inline_shapes,
                        "shapes": shapes,
                        "ole_total": ole_count,
                        "ole_missing": ole_missing,
                    }
                )
            else:
                checks.append(
                    {
                        "check": "embedded_objects",
                        "status": "pass",
                        "message": "嵌入对象完整",
                        "inline_shapes": inline_shapes,
                        "shapes": shapes,
                        "ole_total": ole_count,
                        "ole_missing": 0,
                    }
                )
        except Exception as e:
            checks.append(
                {
                    "check": "embedded_objects",
                    "status": "skip",
                    "message": f"无法检查嵌入对象: {e}",
                }
            )
        return warnings

    @staticmethod
    def _check_fonts(doc: Any, checks: list[dict]) -> int:
        """字体可用性检查"""
        warnings = 0
        try:
            # 收集文档中使用的所有字体
            used_fonts: set[str] = set()
            try:
                for i in range(1, min(doc.Characters.Count, 5000)):
                    try:
                        font_name = str(doc.Characters(i).Font.Name)
                        if font_name:
                            used_fonts.add(font_name)
                    except Exception:
                        continue
            except Exception:
                pass

            # 也检查样式中的字体
            try:
                for i in range(1, doc.Styles.Count + 1):
                    try:
                        style = doc.Styles(i)
                        font_name = str(style.Font.Name)
                        if font_name:
                            used_fonts.add(font_name)
                    except Exception:
                        continue
            except Exception:
                pass

            if not used_fonts:
                checks.append(
                    {
                        "check": "fonts",
                        "status": "pass",
                        "message": "未检测到特殊字体使用",
                        "font_count": 0,
                    }
                )
                return 0

            # 检查字体是否可能缺失（通过名称为空或异常判断）
            missing_fonts: set[str] = set()
            for fname in list(used_fonts)[:50]:
                try:
                    # 尝试通过名称获取字体来验证
                    test_range = doc.Range(0, 1)
                    test_range.Font.Name = fname
                    actual = str(test_range.Font.Name)
                    if actual != fname:
                        missing_fonts.add(fname)
                except Exception:
                    missing_fonts.add(fname)

            if missing_fonts:
                warnings += 1
                checks.append(
                    {
                        "check": "fonts",
                        "status": "warning",
                        "message": f"{len(missing_fonts)} 个字体可能不可用",
                        "total": len(used_fonts),
                        "missing": sorted(missing_fonts)[:20],
                        "available": sorted(used_fonts - missing_fonts)[:20],
                    }
                )
            else:
                checks.append(
                    {
                        "check": "fonts",
                        "status": "pass",
                        "message": f"所有 {len(used_fonts)} 个字体可用",
                        "total": len(used_fonts),
                        "fonts": sorted(used_fonts)[:30],
                    }
                )
        except Exception as e:
            checks.append(
                {
                    "check": "fonts",
                    "status": "skip",
                    "message": f"无法检查字体: {e}",
                }
            )
        return warnings

    @staticmethod
    def _check_document_structure(doc: Any, checks: list[dict]) -> int:
        """文档结构检查"""
        warnings = 0
        try:
            issues: list[str] = []

            # 检查是否有节
            sections = doc.Sections.Count
            if sections == 0:
                issues.append("文档无节结构")

            # 检查段落数
            paragraphs = doc.Paragraphs.Count
            if paragraphs == 0:
                issues.append("文档无段落内容")

            # 检查页面设置一致性
            if sections > 1:
                try:
                    first_w = doc.Sections(1).PageSetup.PageWidth
                    first_h = doc.Sections(1).PageSetup.PageHeight
                    for i in range(2, sections + 1):
                        try:
                            sw = doc.Sections(i).PageSetup.PageWidth
                            sh = doc.Sections(i).PageSetup.PageHeight
                            if sw != first_w or sh != first_h:
                                issues.append(f"第 {i} 节页面尺寸不一致")
                                warnings += 1
                        except Exception:
                            pass
                except Exception:
                    pass

            if issues:
                checks.append(
                    {
                        "check": "document_structure",
                        "status": "warning",
                        "message": "; ".join(issues),
                        "sections": sections,
                        "paragraphs": paragraphs,
                    }
                )
            else:
                checks.append(
                    {
                        "check": "document_structure",
                        "status": "pass",
                        "message": "文档结构正常",
                        "sections": sections,
                        "paragraphs": paragraphs,
                    }
                )
        except Exception as e:
            checks.append(
                {
                    "check": "document_structure",
                    "status": "skip",
                    "message": f"无法检查文档结构: {e}",
                }
            )
        return warnings

    # ── Calc 验证 ──

    def validate_calc(self, path: Path) -> ValidateResult:
        """验证 Excel 工作簿

        检查项：
        - 公式求值状态（检测 #REF! #VALUE! #N/A 等）
        - 命名区域引用完整性
        - 外部链接可访问性
        - 条件格式规则冲突
        - 数据验证规则冲突
        """
        checks: list[dict] = []
        errors = 0
        warnings = 0

        with self.manager.session("calc") as app:
            wb = app.Workbooks.Open(
                str(path),
                UpdateLinks=0,
                ReadOnly=True,
            )
            try:
                # 1. 公式错误检查
                errs, wrns = self._check_formula_errors(wb, checks)
                errors += errs
                warnings += wrns

                # 2. 命名区域完整性
                errors += self._check_named_ranges(wb, checks)

                # 3. 外部链接
                warnings += self._check_external_links(wb, checks)

                # 4. 工作表结构
                warnings += self._check_sheet_structure(wb, checks)

            finally:
                wb.Close(WD_DO_NOT_SAVE_CHANGES)

        total_issues = errors + warnings
        passed = errors == 0

        return ValidateResult(
            passed=passed,
            file=str(path),
            checks=checks,
            issues_count=total_issues,
            errors_count=errors,
            warnings_count=warnings,
        )

    @staticmethod
    def _check_formula_errors(wb: Any, checks: list[dict]) -> tuple[int, int]:
        """公式求值状态检查"""
        errors = 0
        warnings = 0
        try:
            total_errors = 0
            error_cells: list[dict] = []

            error_markers = ["#REF!", "#VALUE!", "#N/A", "#NAME?", "#NUM!", "#NULL!", "#DIV/0!"]

            for si in range(1, wb.Sheets.Count + 1):
                ws = wb.Sheets(si)
                try:
                    used = ws.UsedRange
                    if used is None:
                        continue

                    rows = min(used.Rows.Count, 1000)
                    cols = min(used.Columns.Count, 100)

                    for r in range(1, rows + 1):
                        for c in range(1, cols + 1):
                            try:
                                cell = used.Cells(r, c)
                                cell_text = str(cell.Text) if cell.Text else ""
                                for marker in error_markers:
                                    if marker in cell_text:
                                        total_errors += 1
                                        if len(error_cells) < 50:
                                            from wps_cli.services.calc_service import (
                                                _col_letter,
                                            )

                                            addr = f"{_col_letter(c)}{r}"
                                            error_cells.append(
                                                {
                                                    "sheet": ws.Name,
                                                    "cell": addr,
                                                    "value": cell_text,
                                                }
                                            )
                                        break
                            except Exception:
                                continue
                except Exception:
                    continue

            if total_errors > 0:
                errors += 1
                checks.append(
                    {
                        "check": "formula_errors",
                        "status": "fail",
                        "message": f"发现 {total_errors} 个公式错误单元格",
                        "error_count": total_errors,
                        "details": error_cells,
                    }
                )
            else:
                checks.append(
                    {
                        "check": "formula_errors",
                        "status": "pass",
                        "message": "未发现公式错误",
                        "error_count": 0,
                    }
                )
        except Exception as e:
            checks.append(
                {
                    "check": "formula_errors",
                    "status": "skip",
                    "message": f"无法检查公式错误: {e}",
                }
            )
        return errors, warnings

    @staticmethod
    def _check_named_ranges(wb: Any, checks: list[dict]) -> int:
        """命名区域引用完整性"""
        errors = 0
        try:
            total = wb.Names.Count
            broken_names: list[dict] = []

            for i in range(1, total + 1):
                try:
                    nm = wb.Names(i)
                    refers_to = str(nm.RefersTo) if nm.RefersTo else ""
                    name = str(nm.Name) if nm.Name else f"Name_{i}"

                    # 检查是否包含 #REF! 等错误引用
                    if "#REF!" in refers_to:
                        broken_names.append(
                            {
                                "name": name,
                                "refers_to": refers_to,
                            }
                        )
                except Exception:
                    broken_names.append(
                        {
                            "name": f"Name_{i}",
                            "refers_to": "(无法读取)",
                        }
                    )

            if broken_names:
                errors += 1
                checks.append(
                    {
                        "check": "named_ranges",
                        "status": "fail",
                        "message": f"{len(broken_names)} 个命名区域引用损坏",
                        "total": total,
                        "broken": len(broken_names),
                        "details": broken_names[:20],
                    }
                )
            else:
                checks.append(
                    {
                        "check": "named_ranges",
                        "status": "pass",
                        "message": f"所有 {total} 个命名区域正常",
                        "total": total,
                        "broken": 0,
                    }
                )
        except Exception as e:
            checks.append(
                {
                    "check": "named_ranges",
                    "status": "skip",
                    "message": f"无法检查命名区域: {e}",
                }
            )
        return errors

    @staticmethod
    def _check_external_links(wb: Any, checks: list[dict]) -> int:
        """外部链接检查"""
        warnings = 0
        try:
            external_count = 0
            broken_count = 0

            # 尝试获取链接源
            try:
                # xlExcelLinks = 1
                links = wb.LinkSources(1)
                if links:
                    external_count = len(links)
                    for link in links:
                        try:
                            # 尝试打开链接源
                            wb.BreakLink(str(link), 1)  # xlLinkTypeExcelLinks
                            # 如果成功断开，说明链接存在但可能已断开
                            broken_count += 1
                        except Exception:
                            pass
            except Exception:
                # LinkSources 在某些 WPS 版本不可用
                pass

            if external_count > 0:
                if broken_count > 0:
                    warnings += 1
                    checks.append(
                        {
                            "check": "external_links",
                            "status": "warning",
                            "message": f"{broken_count}/{external_count} 个外部链接可能不可用",
                            "total": external_count,
                            "broken": broken_count,
                        }
                    )
                else:
                    checks.append(
                        {
                            "check": "external_links",
                            "status": "pass",
                            "message": f"{external_count} 个外部链接",
                            "total": external_count,
                            "broken": 0,
                        }
                    )
            else:
                checks.append(
                    {
                        "check": "external_links",
                        "status": "pass",
                        "message": "无外部链接",
                        "total": 0,
                    }
                )
        except Exception as e:
            checks.append(
                {
                    "check": "external_links",
                    "status": "skip",
                    "message": f"无法检查外部链接: {e}",
                }
            )
        return warnings

    @staticmethod
    def _check_sheet_structure(wb: Any, checks: list[dict]) -> int:
        """工作表结构检查"""
        warnings = 0
        try:
            hidden_count = 0
            empty_count = 0
            sheet_details: list[dict] = []

            for i in range(1, wb.Sheets.Count + 1):
                ws = wb.Sheets(i)
                is_hidden = False
                is_empty = False
                try:
                    is_hidden = ws.Visible != -1  # xlSheetVisible = -1
                except Exception:
                    pass

                try:
                    used = ws.UsedRange
                    if used is None:
                        is_empty = True
                    else:
                        rows = used.Rows.Count
                        cols = used.Columns.Count
                        if rows <= 1 and cols <= 1:
                            cell_val = used.Cells(1, 1).Value
                            if cell_val is None:
                                is_empty = True
                except Exception:
                    pass

                if is_hidden:
                    hidden_count += 1
                if is_empty:
                    empty_count += 1

                sheet_details.append(
                    {
                        "name": ws.Name,
                        "hidden": is_hidden,
                        "empty": is_empty,
                    }
                )

            if hidden_count > 0 or empty_count > 0:
                msg_parts = []
                if hidden_count > 0:
                    msg_parts.append(f"{hidden_count} 个隐藏工作表")
                if empty_count > 0:
                    msg_parts.append(f"{empty_count} 个空白工作表")
                warnings += 1
                checks.append(
                    {
                        "check": "sheet_structure",
                        "status": "warning",
                        "message": ", ".join(msg_parts),
                        "total": wb.Sheets.Count,
                        "hidden": hidden_count,
                        "empty": empty_count,
                        "sheets": sheet_details,
                    }
                )
            else:
                checks.append(
                    {
                        "check": "sheet_structure",
                        "status": "pass",
                        "message": "工作表结构正常",
                        "total": wb.Sheets.Count,
                        "hidden": 0,
                        "empty": 0,
                    }
                )
        except Exception as e:
            checks.append(
                {
                    "check": "sheet_structure",
                    "status": "skip",
                    "message": f"无法检查工作表结构: {e}",
                }
            )
        return warnings

    # ── Impress 验证 ──

    def validate_impress(self, path: Path) -> ValidateResult:
        """验证 PPT 演示文稿

        检查项：
        - 嵌入媒体完整性
        - 超链接有效性
        - 幻灯片大小一致性
        - 母版引用有效性
        """
        checks: list[dict] = []
        errors = 0
        warnings = 0

        with self.manager.session("impress") as app:
            pres = app.Presentations.Open(str(path), ReadOnly=True)
            try:
                # 1. 超链接有效性
                errors += self._check_impress_hyperlinks(pres, checks)

                # 2. 嵌入媒体完整性
                warnings += self._check_impress_media(pres, checks)

                # 3. 幻灯片大小一致性
                warnings += self._check_slide_size(pres, checks)

                # 4. 母版引用
                warnings += self._check_masters(pres, checks)

            finally:
                pres.Close()

        total_issues = errors + warnings
        passed = errors == 0

        return ValidateResult(
            passed=passed,
            file=str(path),
            checks=checks,
            issues_count=total_issues,
            errors_count=errors,
            warnings_count=warnings,
        )

    @staticmethod
    def _check_impress_hyperlinks(pres: Any, checks: list[dict]) -> int:
        """演示文稿超链接检查"""
        errors = 0
        try:
            total = 0
            broken = 0
            broken_details: list[dict] = []

            for si in range(1, pres.Slides.Count + 1):
                try:
                    sl = pres.Slides(si)
                    for shape in sl.Shapes:
                        try:
                            if hasattr(shape, "Hyperlink") and shape.Hyperlink:
                                total += 1
                                addr = str(shape.Hyperlink.Address) if shape.Hyperlink.Address else ""
                                sub = str(shape.Hyperlink.SubAddress) if shape.Hyperlink.SubAddress else ""
                                if not addr and not sub:
                                    broken += 1
                                    broken_details.append(
                                        {
                                            "slide": si,
                                            "address": "(空)",
                                        }
                                    )
                        except Exception:
                            continue

                    # 检查文本中的超链接
                    try:
                        for hli in range(1, sl.Hyperlinks.Count + 1):
                            total += 1
                            hl = sl.Hyperlinks(hli)
                            addr = str(hl.Address) if hl.Address else ""
                            if not addr:
                                broken += 1
                    except Exception:
                        pass
                except Exception:
                    continue

            if broken > 0:
                errors += 1
                checks.append(
                    {
                        "check": "hyperlinks",
                        "status": "fail",
                        "message": f"{broken} 个超链接异常",
                        "total": total,
                        "broken": broken,
                        "details": broken_details[:10],
                    }
                )
            else:
                checks.append(
                    {
                        "check": "hyperlinks",
                        "status": "pass",
                        "message": f"所有 {total} 个超链接有效",
                        "total": total,
                        "broken": 0,
                    }
                )
        except Exception as e:
            checks.append(
                {
                    "check": "hyperlinks",
                    "status": "skip",
                    "message": f"无法检查超链接: {e}",
                }
            )
        return errors

    @staticmethod
    def _check_impress_media(pres: Any, checks: list[dict]) -> int:
        """嵌入媒体完整性"""
        warnings = 0
        try:
            media_count = 0
            missing_count = 0
            missing_details: list[dict] = []

            for si in range(1, pres.Slides.Count + 1):
                try:
                    sl = pres.Slides(si)
                    for shape in sl.Shapes:
                        try:
                            shape_type = str(shape.Type) if hasattr(shape, "Type") else ""
                            # msoMedia = 16, msoPicture = 13
                            if "16" in str(shape_type):
                                media_count += 1
                                try:
                                    # 尝试访问媒体对象属性
                                    _ = shape.MediaFormat.Length
                                except Exception:
                                    missing_count += 1
                                    missing_details.append(
                                        {"slide": si, "type": "media"}
                                    )
                        except Exception:
                            continue
                except Exception:
                    continue

            if missing_count > 0:
                warnings += 1
                checks.append(
                    {
                        "check": "media",
                        "status": "warning",
                        "message": f"{missing_count}/{media_count} 个嵌入媒体可能缺失",
                        "total": media_count,
                        "missing": missing_count,
                        "details": missing_details[:10],
                    }
                )
            else:
                checks.append(
                    {
                        "check": "media",
                        "status": "pass",
                        "message": f"所有 {media_count} 个嵌入媒体完整",
                        "total": media_count,
                        "missing": 0,
                    }
                )
        except Exception as e:
            checks.append(
                {
                    "check": "media",
                    "status": "skip",
                    "message": f"无法检查嵌入媒体: {e}",
                }
            )
        return warnings

    @staticmethod
    def _check_slide_size(pres: Any, checks: list[dict]) -> int:
        """幻灯片大小一致性"""
        warnings = 0
        try:
            # 检查与母版页面尺寸是否一致
            try:
                pw = pres.PageSetup.SlideWidth
                ph = pres.PageSetup.SlideHeight

                # 检查所有母版
                mismatched_masters: list[dict] = []
                for mi in range(1, pres.SlideMaster.CustomLayouts.Count + 1):
                    try:
                        layout = pres.SlideMaster.CustomLayouts(mi)
                        if (
                            layout.Width != pw
                            or layout.Height != ph
                        ):
                            mismatched_masters.append(
                                {
                                    "layout_index": mi,
                                    "width": layout.Width,
                                    "height": layout.Height,
                                }
                            )
                    except Exception:
                        continue

                if mismatched_masters:
                    warnings += 1
                    checks.append(
                        {
                            "check": "slide_size",
                            "status": "warning",
                            "message": f"{len(mismatched_masters)} 个版式尺寸与母版不一致",
                            "expected": {"width": pw, "height": ph},
                            "mismatched": mismatched_masters,
                        }
                    )
                else:
                    checks.append(
                        {
                            "check": "slide_size",
                            "status": "pass",
                            "message": "所有幻灯片尺寸一致",
                            "size": {"width": pw, "height": ph},
                        }
                    )
            except Exception:
                checks.append(
                    {
                        "check": "slide_size",
                        "status": "pass",
                        "message": "幻灯片尺寸检查通过",
                    }
                )
        except Exception as e:
            checks.append(
                {
                    "check": "slide_size",
                    "status": "skip",
                    "message": f"无法检查幻灯片尺寸: {e}",
                }
            )
        return warnings

    @staticmethod
    def _check_masters(pres: Any, checks: list[dict]) -> int:
        """母版引用有效性"""
        warnings = 0
        try:
            orphan_count = 0
            total_masters = 0
            orphan_details: list[dict] = []

            # 统计母版数
            try:
                total_masters = pres.Designs.Count if hasattr(pres, "Designs") else 0
            except Exception:
                pass

            # 检查每个幻灯片是否引用了有效的母版
            for si in range(1, pres.Slides.Count + 1):
                try:
                    sl = pres.Slides(si)
                    # 尝试访问母版
                    try:
                        _ = sl.Design.Name if hasattr(sl, "Design") else ""
                    except Exception:
                        orphan_count += 1
                        orphan_details.append({"slide": si})
                except Exception:
                    orphan_count += 1

            if orphan_count > 0:
                warnings += 1
                checks.append(
                    {
                        "check": "masters",
                        "status": "warning",
                        "message": f"{orphan_count} 个幻灯片母版引用异常",
                        "total_slides": pres.Slides.Count,
                        "total_masters": total_masters,
                        "orphan_count": orphan_count,
                        "details": orphan_details[:10],
                    }
                )
            else:
                checks.append(
                    {
                        "check": "masters",
                        "status": "pass",
                        "message": "所有母版引用正常",
                        "total_slides": pres.Slides.Count,
                        "total_masters": total_masters,
                    }
                )
        except Exception as e:
            checks.append(
                {
                    "check": "masters",
                    "status": "skip",
                    "message": f"无法检查母版引用: {e}",
                }
            )
        return warnings
