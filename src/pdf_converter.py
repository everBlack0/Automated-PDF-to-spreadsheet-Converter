#!/usr/bin/env python3

"""
PDF-to-Spreadsheet Converter
Extracts structured data from PDF forms and converts to Excel
"""

import os
import re
import pandas as pd
import pdfplumber
from pathlib import Path
from datetime import datetime

class PDFConverter:
    def __init__(self):
        self.extraction_patterns = {
            # Personal Information
            'name': [
                r'Name:\s*([^\n]+)',
                r'Full Name:\s*([^\n]+)',
                r'Applicant Name:\s*([^\n]+)'
            ],
            'email': [
                r'Email(?:\s+Address)?:\s*([^\s\n]+@[^\s\n]+)',
                r'E-mail:\s*([^\s\n]+@[^\s\n]+)'
            ],
            'phone': [
                r'Phone(?:\s+Number)?:\s*([^\n]+)',
                r'Telephone:\s*([^\n]+)',
                r'Contact Number:\s*([^\n]+)'
            ],
            'address': [
                r'Address:\s*([^\n]+(?:\n[^\n:]+)*?)(?=\n[A-Z][a-z]*:|$)',
                r'Mailing Address:\s*([^\n]+)'
            ],
            'date_of_birth': [
                r'Date of Birth:\s*([^\n]+)',
                r'DOB:\s*([^\n]+)',
                r'Birth Date:\s*([^\n]+)'
            ],
            'ssn': [
                r'Social Security Number:\s*([^\n]+)',
                r'SSN:\s*([^\n]+)'
            ],
            
            # Position Information
            'position': [
                r'Position Applied For:\s*([^\n]+)',
                r'Job Title:\s*([^\n]+)',
                r'Position:\s*([^\n]+)'
            ],
            'salary': [
                r'Desired Salary:\s*([^\n]+)',
                r'Expected Salary:\s*([^\n]+)',
                r'Salary Expectation:\s*([^\n]+)'
            ],
            'experience': [
                r'Years of Experience:\s*([^\n]+)',
                r'Experience:\s*([^\n]+)',
                r'Work Experience:\s*([^\n]+)'
            ],
            'availability': [
                r'Availability:\s*([^\n]+)',
                r'Start Date:\s*([^\n]+)'
            ],
            'employment_type': [
                r'Employment Type:\s*([^\n]+)',
                r'Position Type:\s*([^\n]+)'
            ],
            
            # Education
            'degree': [
                r'Highest Degree:\s*([^\n]+)',
                r'Degree:\s*([^\n]+)',
                r'Education Level:\s*([^\n]+)'
            ],
            'major': [
                r'Major/Field of Study:\s*([^\n]+)',
                r'Major:\s*([^\n]+)',
                r'Field of Study:\s*([^\n]+)'
            ],
            'institution': [
                r'Institution:\s*([^\n]+)',
                r'University:\s*([^\n]+)',
                r'School:\s*([^\n]+)'
            ],
            'graduation_year': [
                r'Graduation Year:\s*([^\n]+)',
                r'Grad Year:\s*([^\n]+)'
            ],
            'gpa': [
                r'GPA:\s*([^\n]+)',
                r'Grade Point Average:\s*([^\n]+)'
            ],
            
            # Survey specific
            'customer_name': [
                r'Customer Name:\s*([^\n]+)'
            ],
            'purchase_date': [
                r'Purchase Date:\s*([^\n]+)'
            ],
            'product': [
                r'Product(?:\s+Purchased)?:\s*([^\n]+)'
            ],
            'rating': [
                r'(?:Overall\s+)?(?:Satisfaction\s+)?Rating:\s*([^\n]+)',
                r'Rate.*?:\s*([^\n]+)'
            ],
            'recommend': [
                r'Would Recommend:\s*([^\n]+)',
                r'Recommendation:\s*([^\n]+)'
            ],
            'customer_id': [
                r'Customer ID:\s*([^\n]+)',
                r'ID:\s*([^\n]+)'
            ]
        }
    
    def extract_text_from_pdf(self, pdf_path):
        """Extract all text from PDF file"""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() + "\n"
                return text
        except Exception as e:
            print(f"Error extracting text from {pdf_path}: {e}")
            return ""
    
    def extract_data_from_text(self, text):
        """Extract structured data using regex patterns"""
        extracted_data = {}
        
        for field, patterns in self.extraction_patterns.items():
            value = None
            
            for pattern in patterns:
                match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
                if match:
                    value = match.group(1).strip()
                    value = re.sub(r'\s+', ' ', value)
                    break
            
            extracted_data[field] = value if value else "Not Found"
        
        return extracted_data
    
    def process_single_pdf(self, pdf_path):
        """Process a single PDF file and extract data"""
        text = self.extract_text_from_pdf(pdf_path)
        
        if not text.strip():
            print(f"No text extracted from {pdf_path}")
            return None
        
        data = self.extract_data_from_text(text)
        
        data['source_file'] = os.path.basename(pdf_path)
        data['processed_date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        return data
    
    def save_to_excel(self, data, output_file="extracted_data.xlsx"):
        """Save extracted data to Excel file"""
        if not data:
            return False
        
        try:
            df = pd.DataFrame(data)
            
            priority_columns = [
                'source_file', 'name', 'email', 'phone', 'position', 
                'customer_name', 'product', 'rating', 'processed_date'
            ]
            
            all_columns = list(df.columns)
            ordered_columns = []
            
            for col in priority_columns:
                if col in all_columns:
                    ordered_columns.append(col)
                    all_columns.remove(col)
            
            ordered_columns.extend(all_columns)
            df = df[ordered_columns]
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Extracted_Data', index=False)
                
                workbook = writer.book
                worksheet = writer.sheets['Extracted_Data']
                
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            return True
            
        except Exception as e:
            print(f"Error saving Excel file: {e}")
            return False

def main():
    """Main execution function"""
    print("PDF-to-Spreadsheet Converter")
    
    converter = PDFConverter()
    
    pdf_dir = "data/input"
    output_dir = "data/output"
    
    if not os.path.exists(pdf_dir):
        print(f"Directory '{pdf_dir}' not found!")
        print("Please create the directory and add PDF files.")
        return
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    pdf_files = [f for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf')]
    if not pdf_files:
        print(f"No PDF files found in '{pdf_dir}'")
        return
    
    print("\nAvailable PDF files:")
    for idx, fname in enumerate(pdf_files, 1):
        print(f"  {idx}. {fname}")
    
    try:
        choice = int(input(f"\nSelect a PDF to convert (1-{len(pdf_files)}): "))
        if choice < 1 or choice > len(pdf_files):
            print("Invalid selection.")
            return
    except ValueError:
        print("Please enter a valid number.")
        return
    
    selected_pdf = pdf_files[choice - 1]
    pdf_path = os.path.join(pdf_dir, selected_pdf)
    
    print(f"Processing: {selected_pdf}")
    
    data = converter.process_single_pdf(pdf_path)
    if not data:
        print("Failed to extract data from PDF")
        return
    
    output_file = os.path.join(output_dir, f"{os.path.splitext(selected_pdf)[0]}_extracted.xlsx")
    
    success = converter.save_to_excel([data], output_file)
    if success:
        print(f"\nConversion completed successfully!")
        print(f"Output file: {output_file}")
        print(f"Source PDF: {pdf_path}")
    else:
        print("Failed to save Excel file")

if __name__ == "__main__":
    main()