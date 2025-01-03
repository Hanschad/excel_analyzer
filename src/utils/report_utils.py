"""Utilities for generating analysis reports"""
import json
from typing import List
from ..models import CellError, AnalysisReport, ErrorSeverity

def generate_report(file_name: str, errors: List[CellError]) -> AnalysisReport:
    """Generate analysis report from errors"""
    errors_by_severity = {sev: [] for sev in ErrorSeverity}
    errors_by_sheet = {}
    
    for error in errors:
        # Group by severity
        errors_by_severity[error.severity].append(error)
        
        # Group by sheet
        if error.sheet_name not in errors_by_sheet:
            errors_by_sheet[error.sheet_name] = []
        errors_by_sheet[error.sheet_name].append(error)
    
    return AnalysisReport(
        file_name=file_name,
        total_errors=len(errors),
        errors_by_severity=errors_by_severity,
        errors_by_sheet=errors_by_sheet
    )

def export_report_json(report: AnalysisReport, output_file: str):
    """Export report to JSON file"""
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(report.to_dict(), f, indent=2)

def export_report_html(report: AnalysisReport, output_file: str):
    """Export report to HTML file"""
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Excel Analysis Report - {report.file_name}</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 20px; }}
            .error {{ color: red; }}
            .warning {{ color: orange; }}
            .info {{ color: blue; }}
        </style>
    </head>
    <body>
        <h1>Excel Analysis Report</h1>
        <h2>File: {report.file_name}</h2>
        <p>Total errors found: {report.total_errors}</p>
        
        <h3>Errors by Severity</h3>
        {_generate_severity_section(report)}
        
        <h3>Errors by Worksheet</h3>
        {_generate_sheet_section(report)}
    </body>
    </html>
    """
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html_content)

def _generate_severity_section(report: AnalysisReport) -> str:
    sections = []
    for severity in ErrorSeverity:
        errors = report.errors_by_severity[severity]
        if errors:
            sections.append(f"""
            <div class="{severity.value}">
                <h4>{severity.value.title()} ({len(errors)})</h4>
                <ul>
                    {''.join(f'<li>{_format_error(err)}</li>' for err in errors)}
                </ul>
            </div>
            """)
    return '\n'.join(sections)

def _generate_sheet_section(report: AnalysisReport) -> str:
    sections = []
    for sheet, errors in report.errors_by_sheet.items():
        sections.append(f"""
        <div>
            <h4>{sheet} ({len(errors)})</h4>
            <ul>
                {''.join(f'<li>{_format_error(err)}</li>' for err in errors)}
            </ul>
        </div>
        """)
    return '\n'.join(sections)

def _format_error(error: CellError) -> str:
    location = f"Cell {error.column}{error.row}" if error.column else "Sheet level"
    fix = f"<br>Suggestion: {error.fix_suggestion}" if error.fix_suggestion else ""
    return f"{location} - {error.error_type}: {error.details}{fix}" 