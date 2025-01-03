"""Validation functions for Excel constraints"""
from typing import List, Optional
from ..models import CellError
from ..constants import ExcelLimits, INVALID_SHEET_CHARS, ZERO_WIDTH_CHARS

def validate_sheet_name(name: str) -> List[CellError]:
    """Validate sheet name constraints"""
    errors = []
    
    if len(name) > ExcelLimits.MAX_SHEET_NAME_LENGTH:
        errors.append(CellError(
            sheet_name=name,
            row=0,
            column="",
            error_type="Sheet name too long",
            details=f"Sheet name length ({len(name)}) exceeds Excel limit ({ExcelLimits.MAX_SHEET_NAME_LENGTH})"
        ))
    
    found_chars = [c for c in INVALID_SHEET_CHARS if c in name]
    if found_chars:
        errors.append(CellError(
            sheet_name=name,
            row=0,
            column="",
            error_type="Invalid sheet name",
            details=f"Sheet name contains invalid characters: {', '.join(found_chars)}"
        ))
    
    if name.startswith("'") or name.endswith("'"):
        errors.append(CellError(
            sheet_name=name,
            row=0,
            column="",
            error_type="Invalid sheet name",
            details="Sheet name cannot start or end with apostrophe"
        ))
    
    return errors

def validate_formula(formula: str, sheet_name: str, cell_ref: str) -> List[CellError]:
    """Validate formula constraints"""
    errors = []
    from ..utils.xml_utils import parse_cell_reference
    
    if len(formula) > ExcelLimits.MAX_FORMULA_LENGTH:
        col, row = parse_cell_reference(cell_ref)
        errors.append(CellError(
            sheet_name=sheet_name,
            row=row,
            column=col,
            error_type="Formula too long",
            details=f"Formula length ({len(formula)}) exceeds Excel limit ({ExcelLimits.MAX_FORMULA_LENGTH})"
        ))
    
    open_parens = formula.count('(')
    if open_parens > ExcelLimits.MAX_NESTED_FUNCTIONS:
        col, row = parse_cell_reference(cell_ref)
        errors.append(CellError(
            sheet_name=sheet_name,
            row=row,
            column=col,
            error_type="Excessive function nesting",
            details=f"Formula contains too many nested functions ({open_parens})"
        ))
    
    return errors

def validate_hyperlink(url: str, sheet_name: str, cell_ref: str) -> List[CellError]:
    """Validate hyperlink constraints"""
    errors = []
    from ..utils.xml_utils import parse_cell_reference
    
    if len(url) > ExcelLimits.MAX_HYPERLINK_LENGTH:
        col, row = parse_cell_reference(cell_ref)
        errors.append(CellError(
            sheet_name=sheet_name,
            row=row,
            column=col,
            error_type="Hyperlink too long",
            details=f"Hyperlink length ({len(url)}) exceeds Excel limit ({ExcelLimits.MAX_HYPERLINK_LENGTH})"
        ))
    return errors

def validate_style(style_type: str, value: float, sheet_name: str, cell_ref: Optional[str] = None) -> List[CellError]:
    """Validate style constraints"""
    errors = []
    
    if style_type == "font_size" and value > ExcelLimits.MAX_FONT_SIZE:
        errors.append(CellError(
            sheet_name=sheet_name,
            row=0,
            column="",
            error_type="Font size too large",
            details=f"Font size {value} exceeds Excel limit ({ExcelLimits.MAX_FONT_SIZE})"
        ))
    elif style_type == "column_width" and value > ExcelLimits.MAX_COLUMN_WIDTH:
        errors.append(CellError(
            sheet_name=sheet_name,
            row=0,
            column="",
            error_type="Column width too large",
            details=f"Column width {value} exceeds Excel limit ({ExcelLimits.MAX_COLUMN_WIDTH})"
        ))
    elif style_type == "row_height" and value > ExcelLimits.MAX_ROW_HEIGHT:
        errors.append(CellError(
            sheet_name=sheet_name,
            row=0,
            column="",
            error_type="Row height too large",
            details=f"Row height {value} exceeds Excel limit ({ExcelLimits.MAX_ROW_HEIGHT})"
        ))
    
    return errors 