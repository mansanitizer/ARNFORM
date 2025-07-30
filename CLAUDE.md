# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an ARN (Asset Reconstruction Number) Change Form Generator - a web application that automatically populates Word documents from Excel data. The application reads structured Excel files and generates properly formatted Word documents with bold field labels and underlined values while preserving exact document layout.

## Architecture

The codebase follows a dual-mode architecture:

1. **Web Application Mode** (`app.py`): Flask-based web server with drag-and-drop file upload interface
2. **Command-Line Mode** (`populate_arn_form.py`): Standalone script for batch processing

### Core Processing Pipeline
Both modes share the same multi-page document processing logic:
- Excel parsing using `openpyxl` with multi-row support (Row 1=Headers, Row 2+=Data)
- Processes ALL data rows automatically (skips empty rows)
- Word document generation creates one page per Excel data row
- Each page is a complete form with proper page breaks between pages
- Formatting preservation through run-based text styling (bold labels, underlined values)

### Critical File Handling Patterns
The application handles Windows file locking issues through:
- `read_only=True` mode for Excel file access
- `tempfile.mkstemp()` instead of `NamedTemporaryFile` for proper file descriptor management
- Explicit workbook closure in try/finally blocks
- Graceful cleanup with PermissionError handling

## Common Development Commands

### Setup and Environment
```bash
# Cross-platform setup (automatic dependency installation)
# Windows: setup.bat
# Unix/Linux/Mac: ./setup.sh

# Manual virtual environment setup
python3 -m venv venv
source venv/bin/activate  # Unix/Mac
# or
venv\Scripts\activate.bat  # Windows

pip install -r requirements.txt
```

### Running the Application
```bash
# Web application (recommended)
python app.py
# Access at http://localhost:8000

# Command-line processing
python populate_arn_form.py

# Platform-specific launchers
# Windows: run.bat
# Unix/Linux/Mac: ./run.sh
```

### Development and Testing
```bash
# Run web app in debug mode (default in app.py)
python app.py

# Test command-line functionality
python populate_arn_form.py

# Manual file testing (ensure template files exist)
# Required: "Request for Change of Broker.docx" (Word template)
# Required: "Format for ARN change.xlsx" (Excel data source)
```

## Excel Data Format Requirements

The application expects a rigid Excel structure with multi-row support:
- **Row 1**: Headers (`Mutual Fund`, `Folio No`, `PAN`, `Investor [First Holder only]`)
- **Row 2+**: Data values (one form page generated per row)
- **Columns**: A=Mutual Fund, B=Folio No, C=PAN, D=Investor
- **Multi-Page Output**: 3 data rows = 1 Word document with 3 pages
- **File formats**: .xlsx or .xls only
- **Empty Row Handling**: Completely empty rows are automatically skipped

## Word Document Processing Logic

The application creates multi-page documents through template duplication:
- **Template Loading**: Loads Word template once per page generation
- **Page Duplication**: For each Excel row, creates a complete template copy
- **Paragraph Processing**: Uses exact text matching (e.g., `'Mutual Fund:'`)
- **Content Replacement**: Clears paragraphs and rebuilds with formatted runs
- **Formatting**: Applies labels=bold, values=underlined with exact spacing
- **Page Breaks**: Adds page breaks between forms (except after last page)
- **Output**: Single Word file containing multiple complete forms

## Critical Dependencies and Constraints

### Template File Dependencies
- `Request for Change of Broker.docx`: Must exist in root directory
- `Format for ARN change.xlsx`: Required for command-line mode
- Templates contain specific paragraph structures that the code expects

### Platform-Specific Considerations
- Windows file locking requires careful handle management
- Cross-platform path handling for temporary files
- Browser auto-opening varies by platform (macOS: `open`, Linux: `xdg-open`)

### Security and File Handling
- 16MB file size limit enforced by Flask
- Secure filename handling with `werkzeug.utils.secure_filename`
- Temporary file cleanup with fallback for permission errors
- File type validation limited to Excel formats

## Web Interface Architecture

The Flask application provides:
- Single-page interface with drag-and-drop upload (`templates/index.html`)
- Professional styling with gradient backgrounds (`static/style.css`)
- Client-side file validation and progress feedback
- Server-side processing with flash message error handling
- Automatic file download with timestamped filenames

## Error Handling Patterns

When modifying the application, maintain these error handling patterns:
- Excel reading: Return `None` on failure, check before proceeding
- Word processing: Return `True/False` success indicators
- File operations: Use try/finally blocks with explicit cleanup
- Web uploads: Flash user-friendly messages, redirect on errors
- Temporary file management: Handle PermissionError gracefully