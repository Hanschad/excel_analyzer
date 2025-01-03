# Excel Analyzer

A Python tool for analyzing Excel files to detect issues that could cause corruption or compatibility problems.

## Features

- Detects strings exceeding Excel limits
- Identifies special characters (e.g., zero-width characters)
- Validates worksheet names
- Verifies file structure integrity
- Generates detailed analysis reports
- Provides fix suggestions

## Installation

### From Source

```bash
# Clone repository
git clone https://github.com/yourusername/excel_analyzer.git
cd excel_analyzer

# Create virtual environment (optional but recommended)
python -m venv venv
source venv/bin/activate  # Linux/Mac
# or
venv\Scripts\activate  # Windows

# Install dependencies
pip install -r requirements.txt

# Install in development mode
pip install -e .
```

### Using pip (if published to PyPI)

```bash
pip install excel-analyzer
```

## Usage

### Command Line

```bash
# Basic analysis
excel-analyzer path/to/excel_file.xlsx

# Show detailed information
excel-analyzer -v path/to/excel_file.xlsx

# Export JSON report
excel-analyzer path/to/excel_file.xlsx --json report.json

# Export HTML report
excel-analyzer path/to/excel_file.xlsx --html report.html

# Use multiple options
excel-analyzer -v path/to/excel_file.xlsx --json report.json --html report.html
```

### As a Library

```python
from excel_analyzer import ExcelAnalyzer

# Create analyzer instance
analyzer = ExcelAnalyzer()

# Analyze file
errors = analyzer.analyze_file("example.xlsx", verbose=True)

# Process results
for error in errors:
    print(f"Found error in {error.sheet_name}: {error.details}")
    if error.fix_suggestion:
        print(f"Suggestion: {error.fix_suggestion}")
```

## Error Types

- **Long string**: Cell string exceeds Excel limit (32,767 characters)
- **Special character**: Contains zero-width characters
- **Sheet name too long**: Worksheet name exceeds 31 characters
- **XML parsing error**: XML structure is corrupted
- **Invalid file**: File format is invalid or corrupted

## Development

### Project Structure

```
excel_analyzer/
├── src/
│   ├── __init__.py
│   ├── analyzer.py      # Main analysis logic
│   ├── constants.py     # Constants definitions
│   ├── models.py        # Data models
│   └── utils/
│       ├── __init__.py
│       ├── xml_utils.py    # XML processing utilities
│       ├── validators.py   # Validation functions
│       └── report_utils.py # Report generation utilities
├── tests/
│   ├── __init__.py
│   ├── test_analyzer.py
│   └── test_reports.py
├── main.py             # CLI entry point
├── setup.py           # Installation config
└── requirements.txt   # Dependencies
```

### Running Tests

```bash
# Run all tests
python -m unittest discover tests

# Run specific test file
python -m unittest tests/test_analyzer.py

# Run tests with verbose output
python -m unittest discover tests -v
```

### Building Documentation

```bash
# If using Sphinx (needs to be installed)
cd docs
make html
```

## Contributing

Pull Requests are welcome! Please ensure:

1. Update test cases
2. Follow existing code style
3. Update relevant documentation

## License

MIT License - see LICENSE file

## Author

Your Name <your.email@example.com>
