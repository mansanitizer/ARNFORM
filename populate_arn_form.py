#!/usr/bin/env python3

import openpyxl
from docx import Document
import os
from datetime import datetime

def read_excel_data(excel_file_path):
    """Read data from Excel file and return as list of dictionaries (one per row)"""
    workbook = None
    try:
        workbook = openpyxl.load_workbook(excel_file_path)
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

def main():
    # File paths
    excel_file = "Format for ARN change.xlsx"
    docx_file = "Request for Change of Broker.docx"
    
    # Check if files exist
    if not os.path.exists(excel_file):
        print(f"Error: Excel file '{excel_file}' not found!")
        return
    
    if not os.path.exists(docx_file):
        print(f"Error: Word document '{docx_file}' not found!")
        return
    
    try:
        # Read data from Excel (now returns list of dictionaries)
        print("Reading data from Excel file...")
        excel_data = read_excel_data(excel_file)
        
        if excel_data is None or len(excel_data) == 0:
            print("Error: No data found in Excel file or file is empty.")
            return
        
        print(f"Excel data loaded - {len(excel_data)} row(s) found:")
        for i, data in enumerate(excel_data, 1):
            print(f"\nRow {i}:")
            for key, value in data.items():
                print(f"  {key}: {value}")
        
        # Create output file with page count
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        page_count = len(excel_data)
        output_file = f"Populated_ARN_Form_{page_count}pages_{timestamp}.docx"
        
        # Populate Word document (now handles multiple pages)
        print(f"\nPopulating Word document with {page_count} page(s)...")
        result = populate_word_document(docx_file, excel_data, output_file)
        
        if result:
            print(f"\nSuccess! Generated {result} page(s) from {len(excel_data)} Excel row(s).")
            print(f"Output file: {output_file}")
        else:
            print("Error: Failed to populate Word document.")
        
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()