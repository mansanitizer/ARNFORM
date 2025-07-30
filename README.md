# ARN Change Form Generator - Web Application

A web-based application that automatically populates ARN (Asset Reconstruction Number) Change Request forms using Excel data.

## Features

- **Drag & Drop Interface**: Simply drag your Excel file onto the web page
- **Automatic Processing**: Reads Excel data and populates the Word document
- **Professional Formatting**: Bold field labels and underlined values
- **Instant Download**: Get your populated form immediately

## Quick Start

### Windows Users:
1. **First-time setup:** Double-click `setup.bat`
2. **Run the application:** Double-click `run.bat`

### Mac/Linux Users:
1. **First-time setup:** `./setup.sh`
2. **Run the application:** `./run.sh`

### Manual Setup (All Platforms):
1. **Install dependencies:**
   ```bash
   # Create virtual environment
   python3 -m venv venv
   
   # Activate virtual environment
   # Windows:
   venv\Scripts\activate.bat
   # Mac/Linux:
   source venv/bin/activate
   
   # Install requirements
   pip install -r requirements.txt
   ```

2. **Start the application:**
   ```bash
   python app.py
   ```

3. **Open your browser and go to:**
   ```
   http://localhost:8000
   ```

## Excel File Requirements

Your Excel file must have the following structure:

| Column A | Column B | Column C | Column D |
|----------|----------|----------|----------|
| Mutual Fund | Folio No | PAN | Investor [First Holder only] |
| Your Data | Your Data | Your Data | Your Data |

- **Row 1**: Headers (as shown above)
- **Row 2**: Your actual data
- **File Format**: .xlsx or .xls

## How to Use

1. Open the web application in your browser
2. Drag and drop your Excel file onto the upload area (or click to browse)
3. Click "Generate ARN Form"
4. The populated Word document will be automatically downloaded

## Files Included

- `app.py` - Main Flask web application
- `templates/index.html` - Web interface
- `static/style.css` - Styling
- `Request for Change of Broker.docx` - Template document
- `populate_arn_form.py` - Original command-line script (backup)

## Technical Details

- **Backend**: Python Flask
- **Frontend**: HTML5, CSS3, JavaScript
- **Dependencies**: Flask, openpyxl, python-docx
- **File Upload**: Secure file handling with validation
- **Processing**: Preserves exact Word document formatting

## Security Features

- File type validation (only Excel files accepted)
- Secure filename handling
- Temporary file cleanup
- File size limits (16MB max)

## Support

If you encounter any issues:
1. Ensure your Excel file follows the required format
2. Check that the Word template is in the same directory
3. Verify all dependencies are installed in the virtual environment