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
    # 间接引用（可用于混淆绕过）
    "INDIRECT(",
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

# ── XlFormatConditionType ───────────────────────────────────────
# Range.FormatConditions.Add(Type, ...) 的类型参数
# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0)
XL_CF_CELL_VALUE = 1
XL_CF_EXPRESSION = 2
XL_CF_COLOR_SCALE = 3
XL_CF_DATA_BAR = 4
XL_CF_TOP_10 = 5
XL_CF_ICON_SET = 6
XL_CF_UNIQUE_VALUES = 8
XL_CF_TEXT_STRING = 9
XL_CF_ABOVE_AVERAGE = 12

# ── XlFormatConditionOperator ────────────────────────────────────
XL_CF_OP_BETWEEN = 1
XL_CF_OP_NOT_BETWEEN = 2
XL_CF_OP_EQUAL = 3
XL_CF_OP_NOT_EQUAL = 4
XL_CF_OP_GREATER = 5
XL_CF_OP_LESS = 6
XL_CF_OP_GREATER_EQUAL = 7
XL_CF_OP_LESS_EQUAL = 8

# ── XlContainsOperator (文本条件运算符) ─────────────────────────
XL_CONTAINS = 0
XL_DOES_NOT_CONTAIN = 1
XL_BEGINS_WITH = 2
XL_ENDS_WITH = 3

# ── XlDVType (数据验证类型) ──────────────────────────────────────
XL_DV_WHOLE = 1
XL_DV_DECIMAL = 2
XL_DV_LIST = 3
XL_DV_DATE = 4
XL_DV_TIME = 5
XL_DV_TEXT_LENGTH = 6
XL_DV_CUSTOM = 7

# ── XlDVAlertStyle ───────────────────────────────────────────────
XL_DV_ALERT_STOP = 1
XL_DV_ALERT_WARNING = 2
XL_DV_ALERT_INFO = 3

# ── Sparkline 类型 ───────────────────────────────────────────────
XL_SPARK_LINE = 1
XL_SPARK_COLUMN = 2
XL_SPARK_COLUMN_STACKED100 = 3

# ── MsoEncoding (文档 dump 用) ────────────────────────────────────
MSO_ENCODING_UTF8 = 65001

# ── COM 诊断与多版本 ProgID ─────────────────────────────────────
# 各应用类型的候选 ProgID 列表（按优先级排列）
# WPS 12.x 可能不再注册 K 前缀 ProgID，需回退到非 K 前缀
COM_PROGID_CANDIDATES: dict[str, list[str]] = {
    "writer": [
        "KWPS.Application",
        "WPS.Application",
        "Kingsoft.WPS.Application",
    ],
    "calc": [
        "KET.Application",
        "ET.Application",
        "Kingsoft.ET.Application",
    ],
    "impress": [
        "KWPP.Application",
        "WPP.Application",
        "Kingsoft.WPP.Application",
    ],
}
