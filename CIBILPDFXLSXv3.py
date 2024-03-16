# -*- coding: utf-8 -*-
"""
Project: Unstructured PDF Data into organised Spreadsheet
Author: Prashant L. Mitba
Version: 2.6.0
Date: 10-03-2024
Description: Converting CIBIL PDF into Excel
"""
import re
import pandas as pd
import pdfplumber
import spacy
import logging
import sys
import xlwings

def extract_text_from_pdf(pdf_path):
        text = ''
        with pdfplumber.open(pdf_path) as pdf_document:
                for page_num in range(len(pdf_document.pages)):
                        page = pdf_document.pages[page_num]
                        text += page.extract_text()
        return text

def process_pdf_text_with_spacy(text):
    nlp = spacy.load("en_core_web_sm")
    sections = re.findall(r'ACCOUNT DATES AMOUNTS STATUS\n([\s\S]*?)(?=\nACCOUNT DATES AMOUNTS STATUS|$)', text)    
    
    
    data_list = []

    for i, section in enumerate(sections):
        #if re.findall(r'ACCOUNT DATES AMOUNTS STATUS\n([\s\S]*?)\nACCOUNT DATES AMOUNTS STATUS|$', sent.text):
            doc = nlp(section)
            #logging.info(f"\nProcessing Section {i + 1}:\n{sections}")
            # Define patterns for extraction
            patterns = {
                    'Member Name': r'MEMBER NAME: (.+?)(?=\b|$)',
                    'Account Number': r'ACCOUNT NUMBER: (.+?)(?=\b|$)',
                    'Opened Date': r'OPENED: (\S+)(?=\b|$)',
                    'Sanctioned Amount': r'SANCTIONED: ([\d,]+)(?=\b|$)',
                    #'Status': r'CREDIT FACILITY STATUS: (.+?)(?=\b|$)',
                    'Last Payment Date': r'LAST PAYMENT: (\S+)(?=\b|$)',
                    'Current Balance': r'CURRENT BALANCE: ([\d,-]+)(?=\b|$)',
                    'Closed Date': r'CLOSED: (\S+)(?=\b|$)',
                    'Loan Type': r'TYPE: (.+?)(?=\b|$)',
                    'EMI': r'EMI: ([\d,]+)(?=\b|$)',
                    #'Certified Date': r'REPORTED AND CERTIFIED: (.+?)(?=\b|$)',
                    'Overdue': r'OVERDUE: ([\d,]+)(?=\b|$)',
                    'Ownership': r'OWNERSHIP: (.+?)(?=\b|$)',
                    #'History Start': r'PMT HIST START: (\S+)(?=\b|$)',
                    #'History End': r'PMT HIST END: (\S+)(?=\b|$)',
                    'DPD': r'DAYS PAST DUE/ASSET CLASSIFICATION \(UP TO 36 MONTHS; LEFT TO RIGHT\)\n([\s\S]*?)(?=\nACCOUNT DATES AMOUNTS STATUS|$)'
                # Add more patterns as needed...
            }
            # Initialize fields
            fields = {field: None for field in patterns}

            # Extract relevant information using regular expressions
            for field, pattern in patterns.items():
                match = re.search(pattern, section)
                if match:
                    fields[field] = match.group(1).strip()
                    #logging.info(f"Extracted {field}: {fields[field]}")
            # Create a data dictionary
            data = fields.copy()
            data_list.append(data)
            
    data_author = {'Author':['Author: Prashant L. Mitba']}
    data_list.append(data_author)            
    return data_list
 
def save_to_excel(data_list, output_file='extracted_data.xlsx'):
    #data_author = {'Author':['Prashant L. Mitba']}
    #data_list = data_list.append(data_author)
    df = pd.DataFrame(data_list)
    df.to_excel(output_file, index=False)
    #logging.info(f"\nData saved to {output_file}")
 
# Add logging configuration
#logging.basicConfig(filename='extraction_log.txt', level=logging.DEBUG)
 
# Example usage
pdf_path = sys.argv[1]
pdf_text = extract_text_from_pdf(pdf_path)
result = process_pdf_text_with_spacy(pdf_text)
save_to_excel(result)
