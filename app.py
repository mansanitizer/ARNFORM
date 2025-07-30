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
    """Read data from Excel file and return as list of dictionaries (one per row)"""
    workbook = None
    try:
        workbook = openpyxl.load_workbook(excel_file_path, read_only=True)
        sheet = workbook.active
        
        # Get all data rows (row 1 contains headers, data starts from row 2)
        data_rows = []
        max_row = sheet.max_row
        
        for row_num in range(2, max_row + 1):  # Start from row 2, go to last row
            # Check if row has any data (skip completely empty rows)
            mutual_fund = sheet[f'A{row_num}'].value
            folio_no = sheet[f'B{row_num}'].value
            pan = sheet[f'C{row_num}'].value
            investor = sheet[f'D{row_num}'].value
            
            # Convert to strings and strip whitespace for better empty detection
            mutual_fund_str = str(mutual_fund).strip() if mutual_fund is not None else ''
            folio_no_str = str(folio_no).strip() if folio_no is not None else ''
            pan_str = str(pan).strip() if pan is not None else ''
            investor_str = str(investor).strip() if investor is not None else ''
            
            # Skip row if all cells are empty or contain only 'None'
            if not any([mutual_fund_str and mutual_fund_str != 'None', 
                       folio_no_str and folio_no_str != 'None',
                       pan_str and pan_str != 'None', 
                       investor_str and investor_str != 'None']):
                continue
                
            data = {
                'mutual_fund': mutual_fund_str,
                'folio_no': folio_no_str,
                'pan': pan_str,
                'investor': investor_str
            }
            data_rows.append(data)
        
        return data_rows if data_rows else None
    except Exception as e:
        print(f"Error reading Excel file: {str(e)}")
        return None
    finally:
        # Ensure workbook is properly closed
        if workbook:
            workbook.close()

def populate_single_page(doc, data):
    """Helper function to populate a single page with data"""
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
            run2 = paragraph.add_run(str(data['mutual_fund']))
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
            run4 = paragraph.add_run(str(data['pan']))
            run4.underline = True
        
        # Handle "Investor [First Holder only]:  " line (Paragraph 5)
        elif original_text.strip() == 'Investor [First Holder only]:':
            paragraph.clear()
            # Add bold label
            run1 = paragraph.add_run("  Investor [First Holder only]: ")
            run1.bold = True
            # Add underlined value
            run2 = paragraph.add_run(str(data['investor']).strip())
            run2.underline = True
        
        # Handle acknowledgement slip fields
        elif original_text.strip() == 'Mutual Fund :':
            paragraph.clear()
            # Add bold label
            run1 = paragraph.add_run("Mutual Fund : ")
            run1.bold = True
            # Add underlined value
            run2 = paragraph.add_run(str(data['mutual_fund']))
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

def populate_word_document(template_path, data_list, output_path):
    """Populate Word document with multiple pages of Excel data"""
    try:
        # For single page, use simpler approach
        if len(data_list) == 1:
            doc = Document(template_path)
            populate_single_page(doc, data_list[0])
            doc.save(output_path)
            return 1
        
        # For multiple pages, create new document and copy content properly
        output_doc = Document()
        
        # Clear the default empty paragraph
        if output_doc.paragraphs:
            p = output_doc.paragraphs[0]
            p._element.getparent().remove(p._element)
        
        for page_index, data in enumerate(data_list):
            # Load a fresh template for each page
            template_doc = Document(template_path)
            
            # Populate this template with the current row's data
            populate_single_page(template_doc, data)
            
            # Copy all paragraphs and elements from template to output
            for element in template_doc.element.body:
                output_doc.element.body.append(element)
            
            # Add page break after each page except the last one
            if page_index < len(data_list) - 1:
                output_doc.add_page_break()
        
        # Save the multi-page document
        output_doc.save(output_path)
        return len(data_list)  # Return number of pages created
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
            
            # Read data from Excel (now returns list of dictionaries)
            excel_data = read_excel_data(temp_excel_path)
            
            if excel_data is None or len(excel_data) == 0:
                flash('Error reading Excel file or no data found. Please check the file format.')
                return redirect(url_for('index'))
            
            # Create output file with page count
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            page_count = len(excel_data)
            output_filename = f"Populated_ARN_Form_{page_count}pages_{timestamp}.docx"
            output_path = os.path.join(tempfile.gettempdir(), output_filename)
            
            # Populate Word document (now handles multiple pages)
            result = populate_word_document(TEMPLATE_DOCX, excel_data, output_path)
            if result:
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