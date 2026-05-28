"""WPS COM 常量定义

集中管理所有 COM 互操作中使用的魔法数字，避免裸数字散布在业务代码中。
"""

# ── WdSaveOptions ──────────────────────────────────────────────
WD_DO_NOT_SAVE_CHANGES = 0
WD_SAVE_CHANGES = -1

# ── WdStatistic ────────────────────────────────────────────────
WD_STATISTIC_WORDS = 0
WD_STATISTIC_PAGES = 2
WD_STATISTIC_CHARACTERS = 3

# ── WdUnits / WdGoToItem ──────────────────────────────────────
WD_STORY = 6
WD_PAGE = 1

# ── WdReplace ──────────────────────────────────────────────────
WD_REPLACE_ALL = 2

# ── WdBreakType ────────────────────────────────────────────────
WD_PAGE_BREAK = 7

# ── WdLineSpacing ──────────────────────────────────────────────
WD_LINE_SPACE_MULTIPLE = 4

# ── WdExportFormat / WdSaveFormat (Writer) ─────────────────────
WD_FORMAT_DOCUMENT = 0
WD_FORMAT_TEXT = 2
WD_FORMAT_RTF = 6
WD_FORMAT_HTML = 8
WD_FORMAT_DOCUMENT_DEFAULT = 16
WD_FORMAT_PDF = 17

# ── XlSaveAsAccessMode / XlFileFormat (Calc) ───────────────────
XL_CSV = 6
XL_OPEN_XML_WORKBOOK = 51

# ── XlYesNo ────────────────────────────────────────────────────
XL_YES = 1

# ── XlSortOrder ────────────────────────────────────────────────
XL_ASCENDING = 1
XL_DESCENDING = 2

# ── XlAutoFilterOperator ───────────────────────────────────────
XL_FILTER_EQUAL = 1
XL_FILTER_NOT_EQUAL = 2
XL_FILTER_GREATER = 5
XL_FILTER_LESS = 6
XL_FILTER_GREATER_EQUAL = 7
XL_FILTER_LESS_EQUAL = 8
XL_FILTER_NO_OP = 0

# ── XlChartType ────────────────────────────────────────────────
XL_COLUMN_CLUSTERED = 51
XL_LINE = 4
XL_PIE = 5
XL_XY_SCATTER = -4169
XL_AREA = 76

# ── PpSaveAsFileType (Impress) ─────────────────────────────────
PP_SAVE_AS_PRESENTATION = 0
PP_SAVE_AS_OPEN_XML_PRESENTATION = 24
PP_SAVE_AS_PDF = 32

# ── MsoTextOrientation ─────────────────────────────────────────
MSO_TEXT_ORIENTATION_HORIZONTAL = 1

# ── PpPlaceholderType ──────────────────────────────────────────
PP_PLACEHOLDER_TITLE = 1
PP_PLACEHOLDER_BODY = 2
PP_PLACEHOLDER_SUBTITLE = 3

# ── MsoTextEffect ──────────────────────────────────────────────
MSO_TEXT_EFFECT_1 = 1

# ── WdHeaderFooterIndex ────────────────────────────────────────
WD_HEADER_FOOTER_PRIMARY = 1

# ── WdParagraphAlignment (shared by Writer & Impress) ──────────
ALIGN_LEFT = 0
ALIGN_CENTER = 1
ALIGN_RIGHT = 2
ALIGN_JUSTIFY = 3

# ── Writer 格式映射 ────────────────────────────────────────────
WRITER_FORMATS: dict[str, int] = {
    "docx": WD_FORMAT_DOCUMENT_DEFAULT,
    "doc": WD_FORMAT_DOCUMENT,
    "rtf": WD_FORMAT_RTF,
    "txt": WD_FORMAT_TEXT,
    "html": WD_FORMAT_HTML,
    "pdf": WD_FORMAT_PDF,
}

# ── Calc 格式映射 ──────────────────────────────────────────────
CALC_FORMATS: dict[str, int] = {
    "xlsx": XL_OPEN_XML_WORKBOOK,
    "csv": XL_CSV,
}

# ── Impress 格式映射 ───────────────────────────────────────────
IMPRESS_FORMATS: dict[str, int] = {
    "pptx": PP_SAVE_AS_OPEN_XML_PRESENTATION,
    "ppt": PP_SAVE_AS_PRESENTATION,
    "pdf": PP_SAVE_AS_PDF,
}

# PowerPoint 切换效果
PP_TRANSITION_RANDOM = 3844

# ── MsoAutomationSecurity ──────────────────────────────────────
# 用于禁用宏自动执行，防止打开恶意文档时触发宏代码
MSO_AUTOMATION_SECURITY_LOW = 1
MSO_AUTOMATION_SECURITY_BY_UI = 2
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

# ── 安全限制 ────────────────────────────────────────────────────
# 公式中禁止出现的危险函数（COM 注入防护）
DANGEROUS_FORMULA_TOKENS: tuple[str, ...] = (
    # 命令执行类
    "SHELL(",
    "DDE(",
    "DDEAUTO(",
    "EXEC(",
    "EXECUTE(",
    "CALL(",
    "REGISTER(",
    # 外联/数据获取类（数据外泄、SSRF）
    "HYPERLINK(",
    "WEBSERVICE(",
    "ENCODEURL(",
    "FILTERXML(",
    "RTD(",
    "IMPORTDATA(",
    "IMPORTHTML(",
    "IMPORTRANGE(",
    "IMPORTXML(",
    "IMPORTFEED(",
    # _xlfn. 兼容前缀写法（Excel 兼容性命名空间）
    "_XLFN.WEBSERVICE(",
    "_XLFN.FILTERXML(",
    "_XLFN.ENCODEURL(",
    "_XLFN.RTD(",
)

# PDF 页码解析的硬上限，防止 "1-999999" 内存炸弹
MAX_PDF_PAGE_NUMBER = 9999
MAX_PDF_PAGE_RANGE_SIZE = 1000

# glob 匹配结果数量上限（防止 ** 触发大量 COM 操作）
MAX_GLOB_RESULTS = 200

# 文本替换长度上限（防止超长输入或反向引用爆炸）
MAX_REPLACE_TEXT_LEN = 1000

# 受支持的文件扩展名白名单
WRITER_INPUT_EXTENSIONS: frozenset[str] = frozenset(
    {".doc", ".docx", ".wps", ".rtf", ".txt", ".html", ".htm", ".xml"}
)
CALC_INPUT_EXTENSIONS: frozenset[str] = frozenset(
    {".xls", ".xlsx", ".xlsm", ".et", ".ett", ".csv", ".tsv"}
)
IMPRESS_INPUT_EXTENSIONS: frozenset[str] = frozenset(
    {".ppt", ".pptx", ".pps", ".ppsx", ".dps", ".dpt", ".pot", ".potx"}
)
PDF_INPUT_EXTENSIONS: frozenset[str] = frozenset({".pdf"})
