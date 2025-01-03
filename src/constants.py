"""Excel file format limitations and constants"""

class ExcelLimits:
    MAX_STRING_LENGTH = 32767
    MAX_ROWS = 1048576
    MAX_COLUMNS = 16384  # XFD
    MAX_COLUMN_WIDTH = 255
    MAX_ROW_HEIGHT = 409
    MAX_FONT_SIZE = 409
    MAX_FORMULA_LENGTH = 8192
    MAX_SHEET_NAME_LENGTH = 31
    MAX_HYPERLINK_LENGTH = 2079
    MAX_CELL_STYLES = 64000
    MAX_COLORS = 16777216  # RGB colors
    MAX_CONDITIONAL_FORMATS = 64
    MAX_FILTER_CONDITIONS = 2
    MAX_SORT_CONDITIONS = 64
    MAX_NESTED_FUNCTIONS = 64
    MAX_ARGUMENTS = 255
    MAX_SHEETS = 255

class XMLNamespaces:
    MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    RELATIONSHIPS = 'http://schemas.openxmlformats.org/package/2006/relationships'

INVALID_SHEET_CHARS = ['\\', '/', '?', '*', '[', ']']
ZERO_WIDTH_CHARS = '\u200B\u200C\u200D\uFEFF' 