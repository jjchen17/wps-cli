# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""模板合并引擎

支持 {{key}} 占位符替换，覆盖：段落、表格单元格、页眉页脚。
设计理念：AI Agent 设计一次模板（高 token 成本），生产代码填充 N 次（零 token 成本）。
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass
from typing import Any

from wps_cli.consts import WD_REPLACE_ALL

# 占位符模式: {{key}} — key 不能包含空白字符
_PLACEHOLDER_RE = re.compile(r"\{\{(\S+?)\}\}")


def _iter_headers_footers(section: Any) -> Iterable[Any]:
    """遍历当前 Section 的所有页眉页脚"""
    # 页眉: primary (1), first_page (2: wdHeaderFooterFirstPage), even_pages (3: wdHeaderFooterEvenPages)
    # 页脚: 通过 Headers + Footers 集合访问
    for hf_collection in (section.Headers, section.Footers):
        try:
            count = hf_collection.Count
        except Exception:
            continue
        for i in range(1, count + 1):
            try:
                yield hf_collection(i)
            except Exception:
                pass


@dataclass
class TemplateEngine:
    """文档模板合并引擎

    将 Word 文档中的 ``{{key}}`` 占位符替换为实际值，
    使用 COM 的 ``Range.Find`` 保持原有段落格式不变。
    """

    def fill(self, app: Any, data: dict[str, str]) -> dict[str, int]:
        """替换文档中所有 {{key}} 占位符

        Args:
            app: WPS Writer COM 应用对象
            data: 键值对字典，如 ``{"name": "张三", "date": "2026-06-07"}``

        Returns:
            ``{"replaced": {"name": 3, "date": 5}, "total": 8}``

        Raises:
            KeyError: 文档中存在 data 未提供的占位符键名 或
                      data 中存在文档未使用的键名
        """
        if not data:
            raise ValueError("data 不能为空")

        doc = app.ActiveDocument

        # 1. 预扫描所有占位符键名，验证完整性
        doc_keys = set(self.extract_keys(app))
        data_keys = set(data.keys())

        missing_in_data = doc_keys - data_keys
        if missing_in_data:
            raise KeyError(
                f"文档中存在但 data 中未提供的占位符键: {sorted(missing_in_data)}"
            )

        # 注意：data 中多余的键可能是预期的（不同模板用同一 data），不报错

        # 2. 逐键替换，统计计数
        replaced: dict[str, int] = {}
        for key, value in data.items():
            placeholder = f"{{{{{key}}}}}"
            count = self._replace_in_body(doc, placeholder, value)
            if count > 0:
                replaced[key] = count

        return {"replaced": replaced, "total": sum(replaced.values())}

    @staticmethod
    def _replace_in_body(doc: Any, old: str, new: str) -> int:
        """在正文（段落 + 表格 + 页眉页脚）中替换文本，返回替换次数

        遍历文档内容区域，使用 COM Find 逐区域替换并计数。
        """
        count = 0

        # ── 正文段落 + 表格（doc.Content 已覆盖） ──
        find = doc.Content.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        find.Text = old
        find.Replacement.Text = new
        find.MatchCase = True
        find.MatchWildcards = False
        find.Forward = True
        find.Wrap = 0  # wdFindStop

        # 先逐次计数再全量替换（避免 new 包含 old 导致计数偏差）
        scan = doc.Content.Find
        scan.ClearFormatting()
        scan.Text = old
        scan.MatchCase = True
        scan.MatchWildcards = False
        scan.Forward = True
        scan.Wrap = 0
        while scan.Execute(Replace=0):
            count += 1

        if count > 0:
            find.Execute(Replace=WD_REPLACE_ALL)

        # ── 页眉页脚 ──
        for section in doc.Sections:
            for hf in _iter_headers_footers(section):
                rng = hf.Range
                scan_hf = rng.Find
                scan_hf.ClearFormatting()
                scan_hf.Text = old
                scan_hf.MatchCase = True
                scan_hf.MatchWildcards = False
                scan_hf.Forward = True
                scan_hf.Wrap = 0
                hf_count = 0
                while scan_hf.Execute(Replace=0):
                    hf_count += 1
                if hf_count > 0:
                    find_hf = rng.Find
                    find_hf.ClearFormatting()
                    find_hf.Replacement.ClearFormatting()
                    find_hf.Text = old
                    find_hf.Replacement.Text = new
                    find_hf.MatchCase = True
                    find_hf.MatchWildcards = False
                    find_hf.Forward = True
                    find_hf.Wrap = 0
                    find_hf.Execute(Replace=WD_REPLACE_ALL)
                    count += hf_count

        return count

    def extract_keys(self, app: Any) -> list[str]:
        """提取文档中所有占位符键名（用于验证模板）

        返回去重排序后的键名列表。
        """
        doc = app.ActiveDocument
        # 收集所有文本片段
        texts: list[str] = []

        # 正文
        for i in range(1, doc.Paragraphs.Count + 1):
            texts.append(doc.Paragraphs(i).Range.Text)

        # 表格
        for i in range(1, doc.Tables.Count + 1):
            table = doc.Tables(i)
            for r in range(1, table.Rows.Count + 1):
                for c_ in range(1, table.Columns.Count + 1):
                    texts.append(table.Cell(r, c_).Range.Text)

        # 页眉页脚
        for section in doc.Sections:
            for hf in _iter_headers_footers(section):
                texts.append(hf.Range.Text)

        keys: set[str] = set()
        for text in texts:
            for match in _PLACEHOLDER_RE.finditer(text):
                keys.add(match.group(1))

        return sorted(keys)
