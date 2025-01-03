"""Data models for Excel analysis

This module defines the data structures used throughout the analyzer.
It includes error representations, analysis context, and report formats.

Key classes:
- CellError: Represents an issue found in a specific cell
- AnalysisContext: Holds the current analysis state
- AnalysisReport: Contains the complete analysis results"""
from dataclasses import dataclass
from enum import Enum
from typing import Optional, Dict, List

class ErrorSeverity(Enum):
    """Error severity levels"""
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    CRITICAL = "critical"

@dataclass
class CellError:
    sheet_name: str
    row: int
    column: str
    error_type: str
    details: str
    severity: ErrorSeverity = ErrorSeverity.ERROR
    fix_suggestion: Optional[str] = None

@dataclass
class AnalysisContext:
    verbose: bool
    long_string_index: int = None

@dataclass
class AnalysisReport:
    """Analysis report with categorized errors"""
    file_name: str
    total_errors: int
    errors_by_severity: Dict[ErrorSeverity, List[CellError]]
    errors_by_sheet: Dict[str, List[CellError]]
    
    def to_dict(self) -> dict:
        """Convert report to dictionary format"""
        return {
            "file_name": self.file_name,
            "total_errors": self.total_errors,
            "errors_by_severity": {
                sev.value: [self._error_to_dict(err) for err in errs]
                for sev, errs in self.errors_by_severity.items()
            },
            "errors_by_sheet": {
                sheet: [self._error_to_dict(err) for err in errs]
                for sheet, errs in self.errors_by_sheet.items()
            }
        }
    
    @staticmethod
    def _error_to_dict(error: CellError) -> dict:
        """Convert CellError to dictionary"""
        return {
            "sheet_name": error.sheet_name,
            "row": error.row,
            "column": error.column,
            "error_type": error.error_type,
            "details": error.details,
            "severity": error.severity.value,
            "fix_suggestion": error.fix_suggestion
        } 