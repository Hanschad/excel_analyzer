#!/usr/bin/env python3
"""Excel Analyzer CLI entry point

This module provides the command-line interface for the Excel Analyzer tool.
It handles command-line arguments and outputs the analysis results.

Usage:
    excel-analyzer [-v] [--json REPORT.json] [--html REPORT.html] EXCEL_FILE

Options:
    -v, --verbose            Show detailed information during analysis
    --json REPORT.json      Export report in JSON format
    --html REPORT.html      Export report in HTML format

Example:
    excel-analyzer -v example.xlsx --json report.json
"""
import argparse
import sys
import os
from .analyzer import ExcelAnalyzer
from .models import ErrorSeverity
from .utils.report_utils import generate_report, export_report_json, export_report_html

def _get_severity_icon(severity: ErrorSeverity) -> str:
    """Get appropriate icon for severity level"""
    icons = {
        ErrorSeverity.CRITICAL: "🔴",
        ErrorSeverity.ERROR: "🟡",
        ErrorSeverity.WARNING: "🟠",
        ErrorSeverity.INFO: "🔵"
    }
    return icons.get(severity, "•")

def main():
    parser = argparse.ArgumentParser(description='Excel File Structure Analyzer')
    parser.add_argument('file', help='Path to Excel file to analyze')
    parser.add_argument('-v', '--verbose', action='store_true', help='Show detailed information')
    parser.add_argument('--json', help='Export report to JSON file')
    parser.add_argument('--html', help='Export report to HTML file')
    args = parser.parse_args()

    analyzer = ExcelAnalyzer()
    try:
        errors = analyzer.analyze_file(args.file, args.verbose)
        
        # Generate report
        report = generate_report(os.path.basename(args.file), errors)
        
        # Export reports if requested
        if args.json:
            export_report_json(report, args.json)
            if args.verbose:
                print(f"\n💾 JSON report saved to: {args.json}")
                
        if args.html:
            export_report_html(report, args.html)
            if args.verbose:
                print(f"\n💾 HTML report saved to: {args.html}")
        
        # Print results summary
        if not errors:
            print("\n✅ No issues found")
        else:
            print(f"\n⚠️  Found {len(errors)} issues:")
            
            # Group errors by severity
            by_severity = {}
            for error in errors:
                if error.severity not in by_severity:
                    by_severity[error.severity] = []
                by_severity[error.severity].append(error)
            
            # Print errors by severity
            for severity in sorted(by_severity.keys(), key=lambda x: x.value):
                severity_errors = by_severity[severity]
                print(f"\n{_get_severity_icon(severity)} {severity.value.upper()} ({len(severity_errors)}):")
                
                for error in severity_errors:
                    location = f"'{error.sheet_name}' at {error.column}{error.row}" if error.column else f"'{error.sheet_name}'"
                    print(f"  • {error.error_type} in {location}")
                    print(f"    {error.details}")
                    if error.fix_suggestion and args.verbose:
                        print(f"    💡 {error.fix_suggestion}")
                    
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main() 