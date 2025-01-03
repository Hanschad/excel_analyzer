"""Excel Analyzer main implementation

This module contains the core functionality for analyzing Excel files.
It checks for various potential issues that could cause Excel files
to become corrupted or incompatible.

Key features:
- File structure validation
- String length checking
- Special character detection
- Sheet name validation
- Style analysis
"""
import os
import logging
from zipfile import ZipFile, BadZipFile
from openpyxl.utils.exceptions import InvalidFileException
import xml.etree.ElementTree as ET
from typing import List, Optional

from .models import CellError, AnalysisContext, ErrorSeverity
from .constants import ExcelLimits, XMLNamespaces, ZERO_WIDTH_CHARS
from .utils import xml_utils, validators

class ExcelAnalyzer:
    def __init__(self):
        self.errors: List[CellError] = []
        self.context = AnalysisContext(verbose=False)
        self._setup_logging()

    def _setup_logging(self):
        logging.basicConfig(level=logging.INFO, format='%(message)s')
        self.logger = logging.getLogger(__name__)

    def analyze_file(self, file_path: str, verbose: bool = False) -> List[CellError]:
        """Analyze Excel file and locate errors"""
        self.context.verbose = verbose
        self.errors = []
        
        if self.context.verbose:
            print(f"\n📝 Analyzing: {os.path.basename(file_path)}")
        
        try:
            # First check if file exists
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"File does not exist: {file_path}")

            # Then check file corruption
            self._check_file_corruption(file_path)
            
            # If file is valid, proceed with analysis
            with ZipFile(file_path, 'r') as zf:
                if self.context.verbose:
                    print("\n🔍 Checking file structure...")
                
                self.context.long_string_index = self._analyze_shared_strings(zf)
                self._analyze_styles(zf)
                self._analyze_worksheets(zf)
                self._analyze_data_validations(zf)
                
                if self.context.verbose:
                    print(f"\n✅ Analysis complete. Found {len(self.errors)} issues.")
                
        except FileNotFoundError as e:
            print(f"\n❌ {str(e)}")
            raise
        except InvalidFileException as e:
            print(f"\n❌ {str(e)}")
            raise
        except Exception as e:
            print(f"\n❌ Error during analysis: {str(e)}")
            raise
            
        return self.errors

    def _check_file_corruption(self, file_path: str):
        """Check if the file is corrupted"""
        try:
            with open(file_path, 'rb') as f:
                header = f.read(4)
                if header != b'PK\x03\x04':
                    raise InvalidFileException("Invalid file header, not a valid XLSX file")
                
                # Try to open as ZIP file to verify structure
                try:
                    with ZipFile(file_path) as zf:
                        if self.context.verbose:
                            print("✓ File structure is valid")
                except BadZipFile:
                    raise InvalidFileException("File ZIP structure is corrupted")
                    
        except (OSError, IOError) as e:
            raise InvalidFileException(f"File read error: {str(e)}")

    def _analyze_shared_strings(self, zf: ZipFile) -> Optional[int]:
        """Analyze shared strings table"""
        if 'xl/sharedStrings.xml' not in zf.namelist():
            return None
            
        try:
            shared_strings_xml = zf.read('xl/sharedStrings.xml')
            tree = ET.fromstring(shared_strings_xml)
            
            long_string_index = None
            for i, si in enumerate(xml_utils.find_elements(tree, './/main:si')):
                t = si.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                if t is not None:
                    text = t.text or ""
                    if len(text) > ExcelLimits.MAX_STRING_LENGTH:
                        long_string_index = i
                        self.errors.append(CellError(
                            sheet_name="Shared strings",
                            row=0,
                            column="",
                            error_type="Long string",
                            details=f"String index {i} length ({len(text)}) exceeds Excel limit ({ExcelLimits.MAX_STRING_LENGTH})",
                            severity=ErrorSeverity.ERROR,
                            fix_suggestion="Split the string into multiple cells or store in external resource"
                        ))
                    
                    if any(c in text for c in ZERO_WIDTH_CHARS):
                        self.errors.append(CellError(
                            sheet_name="Shared strings",
                            row=0,
                            column="",
                            error_type="Special character",
                            details=f"String index {i} contains zero-width character",
                            severity=ErrorSeverity.WARNING,
                            fix_suggestion="Remove or replace zero-width characters"
                        ))
            
            return long_string_index
        except ET.ParseError as e:
            self.errors.append(CellError(
                sheet_name="Shared strings",
                row=0,
                column="",
                error_type="XML parsing error",
                details=f"Shared strings table XML parsing failed: {str(e)}",
                severity=ErrorSeverity.CRITICAL,
                fix_suggestion="The file may be corrupted. Try recreating it or recovering from backup"
            ))
            return None

    def _analyze_styles(self, zf: ZipFile):
        """Analyze styles table"""
        pass  # Implement style analysis

    def _analyze_worksheets(self, zf: ZipFile):
        """Analyze worksheets"""
        sheet_files = [f for f in zf.namelist() if f.startswith('xl/worksheets/sheet')]
        
        if self.context.verbose:
            self.logger.info(f"Found worksheet files: {sheet_files}")
        
        for sheet_file in sheet_files:
            try:
                sheet_xml = zf.read(sheet_file)
                tree = ET.fromstring(sheet_xml)
                
                # Get sheet name from workbook.xml
                sheet_number = int(sheet_file.split('sheet')[-1].split('.')[0])
                sheet_name = self._get_sheet_name(zf, sheet_number)
                
                if self.context.verbose:
                    self.logger.info(f"\nAnalyzing sheet {sheet_name or f'Sheet{sheet_number}'}")
                
                if sheet_name:
                    self._check_sheet_name(sheet_name, sheet_number)
                
                # Analyze cells
                cells = xml_utils.find_elements(tree, './/main:c')
                if self.context.verbose:
                    self.logger.info(f"Found {len(cells)} cells to analyze")
                
                for cell in cells:
                    cell_ref = cell.get('r', '')
                    cell_type = cell.get('t', '')
                    
                    if self.context.verbose:
                        self.logger.info(f"\nAnalyzing cell {cell_ref} (type: {cell_type})")
                        self.logger.info(f"Cell XML: {ET.tostring(cell, encoding='unicode')}")
                    
                    # Check all possible string values
                    # 1. Check inline strings
                    is_elem = cell.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is')
                    if is_elem is not None:
                        t_elem = is_elem.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                        if t_elem is not None and t_elem.text:
                            self._check_string_content(t_elem.text, cell_ref, sheet_name or f"Sheet{sheet_number}")
                    
                    # 2. Check direct string values
                    value_elem = cell.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                    if value_elem is not None and value_elem.text:
                        if cell_type in ('str', 's', '') or cell_type is None:
                            self._check_string_content(value_elem.text, cell_ref, sheet_name or f"Sheet{sheet_number}")
                
            except ET.ParseError as e:
                self.errors.append(CellError(
                    sheet_name=f"Sheet{sheet_number}",
                    row=0,
                    column="",
                    error_type="XML parsing error",
                    details=f"Worksheet XML parsing failed: {str(e)}",
                    severity=ErrorSeverity.CRITICAL,
                    fix_suggestion="The worksheet may be corrupted. Try recreating it"
                ))

    def _get_sheet_name(self, zf: ZipFile, sheet_number: int) -> Optional[str]:
        """Get sheet name from workbook.xml"""
        try:
            workbook_xml = zf.read('xl/workbook.xml')
            tree = ET.fromstring(workbook_xml)
            sheets = tree.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet')
            
            for sheet in sheets:
                if sheet.get('sheetId') == str(sheet_number):
                    return sheet.get('name')
            
            return None
        except Exception:
            return None

    def _analyze_data_validations(self, zf: ZipFile):
        """Analyze data validations"""
        pass  # Implement data validation analysis 

    def _check_sheet_name(self, name: str, sheet_number: int):
        """Check sheet name constraints"""
        if len(name) > ExcelLimits.MAX_SHEET_NAME_LENGTH:
            self.errors.append(CellError(
                sheet_name=name,
                row=0,
                column="",
                error_type="Sheet name too long",
                details=f"Sheet name length ({len(name)}) exceeds Excel limit ({ExcelLimits.MAX_SHEET_NAME_LENGTH})",
                severity=ErrorSeverity.ERROR,
                fix_suggestion=f"Rename the sheet to use fewer than {ExcelLimits.MAX_SHEET_NAME_LENGTH} characters"
            )) 

    def _check_string_content(self, text: str, cell_ref: str, sheet_name: str):
        """Check string content for various issues"""
        col, row = xml_utils.parse_cell_reference(cell_ref)
        
        if len(text) > ExcelLimits.MAX_STRING_LENGTH:
            self.errors.append(CellError(
                sheet_name=sheet_name,
                row=row,
                column=col,
                error_type="Long string",
                details=f"Cell string length ({len(text)}) exceeds Excel limit ({ExcelLimits.MAX_STRING_LENGTH})",
                severity=ErrorSeverity.ERROR,
                fix_suggestion="Split the string into multiple cells or store in external resource"
            ))
        
        if any(c in text for c in ZERO_WIDTH_CHARS):
            self.errors.append(CellError(
                sheet_name=sheet_name,
                row=row,
                column=col,
                error_type="Special character",
                details="Cell contains zero-width character",
                severity=ErrorSeverity.WARNING,
                fix_suggestion="Remove or replace zero-width characters"
            )) 