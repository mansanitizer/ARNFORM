#!/usr/bin/env python3

from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import openpyxl
from docx import Document
import os
import tempfile
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # Change this to a random secret key
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
TEMPLATE_DOCX = "Request for Change of Broker.docx"

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def read_excel_data(excel_file_path):
    """Read data from Excel file and return as dictionary"""
    workbook = None
    try:
        workbook = openpyxl.load_workbook(excel_file_path, read_only=True)
        sheet = workbook.active
        
        # Get data from row 2 (row 1 contains headers)
        data = {
            'mutual_fund': sheet['A2'].value,
            'folio_no': sheet['B2'].value,
            'pan': sheet['C2'].value,
            'investor': sheet['D2'].value
        }
        
        return data
    except Exception as e:
        print(f"Error reading Excel file: {str(e)}")
        return None
    finally:
        # Ensure workbook is properly closed
        if workbook:
            workbook.close()

def populate_word_document(template_path, data, output_path):
    """Populate Word document with Excel data while preserving exact formatting"""
    try:
        doc = Document(template_path)
        
        # Process paragraphs with precise formatting preservation
        for paragraph in doc.paragraphs:
            original_text = paragraph.text
            
            # Handle "Mutual Fund: " line (Paragraph 3)
            if original_text.strip() == 'Mutual Fund:':
                paragraph.clear()
                # Add bold label
                run1 = paragraph.add_run("  Mutual Fund: ")
                run1.bold = True
                # Add underlined value
                run2 = paragraph.add_run(data['mutual_fund'])
                run2.underline = True
            
            # Handle "Folio No:* ... PAN:* " line (Paragraph 4)
            elif 'Folio No:*' in original_text and 'PAN:*' in original_text:
                paragraph.clear()
                # Add bold "Folio No:*" label
                run1 = paragraph.add_run("      Folio No:* ")
                run1.bold = True
                # Add underlined folio number
                run2 = paragraph.add_run(str(data['folio_no']))
                run2.underline = True
                # Add spacing
                paragraph.add_run("                                                                                                          ")
                # Add bold "PAN:*" label
                run3 = paragraph.add_run("PAN:* ")
                run3.bold = True
                # Add underlined PAN
                run4 = paragraph.add_run(data['pan'])
                run4.underline = True
            
            # Handle "Investor [First Holder only]:  " line (Paragraph 5)
            elif original_text.strip() == 'Investor [First Holder only]:':
                paragraph.clear()
                # Add bold label
                run1 = paragraph.add_run("  Investor [First Holder only]: ")
                run1.bold = True
                # Add underlined value
                run2 = paragraph.add_run(data['investor'].strip())
                run2.underline = True
            
            # Handle acknowledgement slip fields
            elif original_text.strip() == 'Mutual Fund :':
                paragraph.clear()
                # Add bold label
                run1 = paragraph.add_run("Mutual Fund : ")
                run1.bold = True
                # Add underlined value
                run2 = paragraph.add_run(data['mutual_fund'])
                run2.underline = True
            elif 'Folio No :' in original_text and 'Date of Receipt:' in original_text:
                paragraph.clear()
                # Add bold "Folio No :" label
                run1 = paragraph.add_run("Folio No : ")
                run1.bold = True
                # Add underlined folio number
                run2 = paragraph.add_run(str(data['folio_no']))
                run2.underline = True
                # Add spacing and Date of Receipt
                paragraph.add_run("                              		                                       Date of Receipt:	")
        
        # Save the populated document
        doc.save(output_path)
        return True
    except Exception as e:
        print(f"Error populating Word document: {str(e)}")
        return False

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file selected')
        return redirect(url_for('index'))
    
    file = request.files['file']
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        temp_excel_path = None
        
        try:
            # Create temporary file for uploaded Excel
            temp_excel_fd, temp_excel_path = tempfile.mkstemp(suffix='.xlsx')
            
            # Close the file descriptor and save the uploaded file
            os.close(temp_excel_fd)
            file.save(temp_excel_path)
            
            # Read data from Excel
            excel_data = read_excel_data(temp_excel_path)
            
            if excel_data is None:
                flash('Error reading Excel file. Please check the file format.')
                return redirect(url_for('index'))
            
            # Create output file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"Populated_ARN_Form_{timestamp}.docx"
            output_path = os.path.join(tempfile.gettempdir(), output_filename)
            
            # Populate Word document
            if populate_word_document(TEMPLATE_DOCX, excel_data, output_path):
                # Send the populated document
                return send_file(output_path, as_attachment=True, 
                               download_name=output_filename,
                               mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            else:
                flash('Error processing the document. Please try again.')
                return redirect(url_for('index'))
                
        except Exception as e:
            flash(f'Error processing file: {str(e)}')
            return redirect(url_for('index'))
        finally:
            # Clean up temporary Excel file
            if temp_excel_path and os.path.exists(temp_excel_path):
                try:
                    os.unlink(temp_excel_path)
                except PermissionError:
                    # If we can't delete it immediately, it will be cleaned up by the OS eventually
                    pass
    else:
        flash('Please upload a valid Excel file (.xlsx or .xls)')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8000)