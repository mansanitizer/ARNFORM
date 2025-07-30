#!/usr/bin/env python3

import openpyxl
from docx import Document
import os
from datetime import datetime

def read_excel_data(excel_file_path):
    """Read data from Excel file and return as dictionary"""
    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook.active
    
    # Get data from row 2 (row 1 contains headers)
    data = {
        'mutual_fund': sheet['A2'].value,
        'folio_no': sheet['B2'].value,
        'pan': sheet['C2'].value,
        'investor': sheet['D2'].value
    }
    
    workbook.close()
    return data

def populate_word_document(docx_file_path, data, output_file_path):
    """Populate Word document with Excel data while preserving exact formatting"""
    doc = Document(docx_file_path)
    
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
    doc.save(output_file_path)
    print(f"Populated document saved as: {output_file_path}")

def main():
    # File paths
    excel_file = "Format for ARN change.xlsx"
    docx_file = "Request for Change of Broker.docx"
    output_file = "Populated_Request_for_Change_of_Broker.docx"
    
    # Check if files exist
    if not os.path.exists(excel_file):
        print(f"Error: Excel file '{excel_file}' not found!")
        return
    
    if not os.path.exists(docx_file):
        print(f"Error: Word document '{docx_file}' not found!")
        return
    
    try:
        # Read data from Excel
        print("Reading data from Excel file...")
        excel_data = read_excel_data(excel_file)
        
        print("Excel data loaded:")
        for key, value in excel_data.items():
            print(f"  {key}: {value}")
        
        # Populate Word document
        print("\nPopulating Word document...")
        populate_word_document(docx_file, excel_data, output_file)
        
        print(f"\nSuccess! The form has been populated with data from the Excel file.")
        print(f"Output file: {output_file}")
        
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()