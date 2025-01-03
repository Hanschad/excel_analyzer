import unittest
import os
import json
from src.models import CellError, ErrorSeverity
from src.utils.report_utils import generate_report, export_report_json, export_report_html

class TestReports(unittest.TestCase):
    def setUp(self):
        self.test_files_dir = os.path.join(os.path.dirname(__file__), 'test_files')
        os.makedirs(self.test_files_dir, exist_ok=True)
        
        # Create sample errors
        self.errors = [
            CellError(
                sheet_name="Sheet1",
                row=1,
                column="A",
                error_type="Long string",
                details="String too long",
                severity=ErrorSeverity.ERROR,
                fix_suggestion="Split the string into multiple cells"
            ),
            CellError(
                sheet_name="Sheet1",
                row=2,
                column="B",
                error_type="Special character",
                details="Contains zero-width space",
                severity=ErrorSeverity.WARNING,
                fix_suggestion="Remove special characters"
            ),
            CellError(
                sheet_name="Sheet2",
                row=0,
                column="",
                error_type="Sheet name",
                details="Invalid sheet name",
                severity=ErrorSeverity.INFO,
                fix_suggestion="Rename the sheet"
            )
        ]

    def test_generate_report(self):
        """Test report generation"""
        report = generate_report("test.xlsx", self.errors)
        
        self.assertEqual(report.total_errors, 3)
        self.assertEqual(len(report.errors_by_sheet), 2)
        self.assertEqual(len(report.errors_by_severity[ErrorSeverity.ERROR]), 1)
        self.assertEqual(len(report.errors_by_severity[ErrorSeverity.WARNING]), 1)
        self.assertEqual(len(report.errors_by_severity[ErrorSeverity.INFO]), 1)

    def test_export_json(self):
        """Test JSON export"""
        report = generate_report("test.xlsx", self.errors)
        json_file = os.path.join(self.test_files_dir, "report.json")
        
        export_report_json(report, json_file)
        
        # Verify JSON file
        self.assertTrue(os.path.exists(json_file))
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
            self.assertEqual(data["total_errors"], 3)
            self.assertEqual(len(data["errors_by_sheet"]), 2)
            
            # Verify error severity is properly serialized
            self.assertIn("error", data["errors_by_severity"])
            self.assertIn("warning", data["errors_by_severity"])
            self.assertIn("info", data["errors_by_severity"])
            
            # Verify error content
            sheet1_errors = data["errors_by_sheet"]["Sheet1"]
            self.assertEqual(len(sheet1_errors), 2)
            self.assertEqual(sheet1_errors[0]["severity"], "error")
            self.assertEqual(sheet1_errors[1]["severity"], "warning")

    def test_export_html(self):
        """Test HTML export"""
        report = generate_report("test.xlsx", self.errors)
        html_file = os.path.join(self.test_files_dir, "report.html")
        
        export_report_html(report, html_file)
        
        # Verify HTML file
        self.assertTrue(os.path.exists(html_file))
        with open(html_file, 'r', encoding='utf-8') as f:
            content = f.read()
            self.assertIn("Excel Analysis Report", content)
            self.assertIn("test.xlsx", content)
            self.assertIn("Total errors found: 3", content)
            self.assertIn("Sheet1", content)
            self.assertIn("Sheet2", content)

    def tearDown(self):
        # Clean up test files
        if os.path.exists(self.test_files_dir):
            for file in os.listdir(self.test_files_dir):
                os.remove(os.path.join(self.test_files_dir, file))
            os.rmdir(self.test_files_dir) 