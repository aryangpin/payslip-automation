#!/usr/bin/env python3
"""
Payslip Batch Processor
Extracts data from payslip images and fills Excel template
"""

import re
import os
from typing import Dict, Optional
from PIL import Image
import openpyxl
from pathlib import Path


class PayslipExtractor:
    """Extract data from payslip images using OCR"""

    def __init__(self, ocr_engine='easyocr'):
        """
        Initialize the extractor

        Args:
            ocr_engine: 'easyocr' or 'tesseract'
        """
        self.ocr_engine = ocr_engine

        if ocr_engine == 'easyocr':
            try:
                import easyocr
                self.reader = easyocr.Reader(['en'])
                print("Using EasyOCR engine")
            except ImportError:
                print("EasyOCR not available, falling back to Tesseract")
                self.ocr_engine = 'tesseract'

        if self.ocr_engine == 'tesseract':
            try:
                import pytesseract
                self.pytesseract = pytesseract
                print("Using Tesseract engine")
            except ImportError:
                raise ImportError("Neither EasyOCR nor Tesseract is available. Please install one.")

    def extract_text(self, image_path: str) -> str:
        """Extract text from image using OCR"""

        if self.ocr_engine == 'easyocr':
            result = self.reader.readtext(image_path)
            # Combine all detected text with positions
            text_lines = []
            for detection in result:
                bbox, text, confidence = detection
                text_lines.append(text)
            return '\n'.join(text_lines)
        else:
            img = Image.open(image_path)
            return self.pytesseract.image_to_string(img)

    def parse_payslip(self, text: str) -> Dict[str, any]:
        """Parse payslip text and extract relevant fields"""

        data = {
            'company_name': '',
            'employee_no': '',
            'employee_name': '',
            'ic_no': '',
            'period': '',
            'date': '',
            'basic_rate': 0.0,
            'working_days': 0.0,
            'basic_pay': 0.0,
            'allowances': {},
            'overtime': [],
            'monthly_gross': 0.0,
            'epf_employer': 0.0,
            'socso_employer': 0.0,
            'eis_employer': 0.0,
            'ytd_al': 0.0,
            'ytd_mc': 0.0,
            'deduction': 0.0,
            'epf_employee': 0.0,
            'socso_employee': 0.0,
            'eis_employee': 0.0,
            'nett_pay': 0.0
        }

        lines = text.split('\n')

        # Extract company name
        for line in lines:
            if 'INDUSTRIES' in line or 'SDN BHD' in line:
                data['company_name'] = line.strip()
                break

        # Extract employee information
        for i, line in enumerate(lines):
            # Employee number
            if 'EMPLOYEE' in line and 'LINE NO' in line:
                # Look for Y#### pattern
                match = re.search(r'Y\d+', line)
                if match:
                    data['employee_no'] = match.group()

            # Employee name and IC
            if 'NAME' in line:
                # Look in nearby lines for name
                for j in range(max(0, i-2), min(len(lines), i+3)):
                    name_match = re.search(r'KYAW\s+\w+\s+\w+|[A-Z]{2,}\s+[A-Z]{2,}\s+[A-Z]{2,}', lines[j])
                    if name_match:
                        data['employee_name'] = name_match.group().strip()

            if 'I/C NO' in line or 'IC NO' in line:
                match = re.search(r'MD\d+|[A-Z]{2}\d+', line)
                if match:
                    data['ic_no'] = match.group()

            # Period and date
            if 'PAYROLL' in line and ('SEPTEMBER' in line or re.search(r'\w+\s+\d{4}', line)):
                data['period'] = line.strip()

            if 'MONTHLY' in line and 'BANK' in line:
                date_match = re.search(r'\d{2}/\d{2}/\d{4}', line)
                if date_match:
                    data['date'] = date_match.group()

            # Basic rate and working days
            if 'BASIC RATE' in line:
                match = re.search(r'(\d+\.?\d*)', line)
                if match:
                    data['basic_rate'] = float(match.group(1))

            if 'WORKING DAYS' in line:
                match = re.search(r'(\d+\.?\d*)', line)
                if match:
                    data['working_days'] = float(match.group(1))

            # Basic pay
            if 'BASIC PAY' in line and 'DIRECTOR' not in line:
                match = re.search(r'(\d+\.?\d*)', line)
                if match:
                    data['basic_pay'] = float(match.group(1))

            # Allowances
            if 'LEADER ALLW' in line or 'ALLOWANCE' in line:
                match = re.search(r'(\d+\.?\d*)', line)
                if match and 'LEADER' in line:
                    data['allowances']['LEADER_ALLW'] = float(match.group(1))

            # Monthly gross
            if 'MONTHLY GROSS' in line:
                match = re.search(r'(\d+\.?\d*)', line)
                if match:
                    data['monthly_gross'] = float(match.group(1))

            # EPF, SOCSO, EIS
            if "EPF" in line and "YER" in line:
                match = re.search(r'(\d+\.?\d*)', line)
                if match:
                    data['epf_employer'] = float(match.group(1))

            if "SOCSO" in line and "YER" in line:
                match = re.search(r'(\d+\.?\d*)', line)
                if match:
                    data['socso_employer'] = float(match.group(1))

            if "EIS" in line and "YER" in line:
                match = re.search(r'(\d+\.?\d*)', line)
                if match:
                    data['eis_employer'] = float(match.group(1))

            # YTD AL and MC
            if 'YTD AL' in line:
                match = re.search(r'(\d+\.?\d*)', line)
                if match:
                    data['ytd_al'] = float(match.group(1))

            if 'YTD MC' in line:
                match = re.search(r'(\d+\.?\d*)', line)
                if match:
                    data['ytd_mc'] = float(match.group(1))

            # Overtime
            if '1.5 TIMES' in line or 'OVERTIME' in line:
                # Try to extract overtime details
                numbers = re.findall(r'\d+\.?\d*', line)
                if len(numbers) >= 3:
                    data['overtime'].append({
                        'type': '1.5 TIMES',
                        'rate': float(numbers[0]) if numbers else 0,
                        'hours': float(numbers[1]) if len(numbers) > 1 else 0,
                        'amount': float(numbers[2]) if len(numbers) > 2 else 0
                    })

        # Try to extract summary line (last line with all values)
        for line in reversed(lines):
            if 'NETT' in line or re.findall(r'\d+\.?\d*', line):
                numbers = re.findall(r'\d+\.?\d*', line)
                if len(numbers) >= 8:
                    try:
                        data['basic_pay'] = float(numbers[0])
                        # numbers[1] is director fee
                        data['overtime_total'] = float(numbers[2]) if len(numbers) > 2 else 0
                        data['allowance_total'] = float(numbers[3]) if len(numbers) > 3 else 0
                        data['monthly_gross'] = float(numbers[4]) if len(numbers) > 4 else 0
                        data['deduction'] = float(numbers[5]) if len(numbers) > 5 else 0
                        data['epf_employee'] = float(numbers[6]) if len(numbers) > 6 else 0
                        data['socso_employee'] = float(numbers[7]) if len(numbers) > 7 else 0
                        if len(numbers) > 8:
                            data['eis_employee'] = float(numbers[8])
                        if len(numbers) > 9:
                            data['nett_pay'] = float(numbers[9])
                        break
                    except (ValueError, IndexError):
                        pass

        return data


class ExcelWriter:
    """Write extracted data to Excel template"""

    def __init__(self, template_path: str):
        """Load Excel template"""
        self.template_path = template_path
        self.wb = openpyxl.load_workbook(template_path)

    def fill_data(self, data: Dict[str, any], sheet_name: Optional[str] = None):
        """
        Fill extracted data into Excel template

        Args:
            data: Extracted payslip data
            sheet_name: Target sheet name (uses active sheet if None)
        """
        if sheet_name:
            sheet = self.wb[sheet_name]
        else:
            sheet = self.wb.active

        # This is a template mapping - will be adjusted after seeing actual template structure
        # For now, creating a simple data dump

        # Start from row 2 (assuming row 1 has headers)
        row = 2

        # Add basic information
        mappings = {
            'A': 'employee_no',
            'B': 'employee_name',
            'C': 'ic_no',
            'D': 'period',
            'E': 'basic_rate',
            'F': 'working_days',
            'G': 'basic_pay',
            'H': 'monthly_gross',
            'I': 'epf_employer',
            'J': 'socso_employer',
            'K': 'eis_employer',
            'L': 'epf_employee',
            'M': 'socso_employee',
            'N': 'eis_employee',
            'O': 'nett_pay',
        }

        for col, field in mappings.items():
            if field in data:
                sheet[f'{col}{row}'] = data[field]

    def save(self, output_path: str):
        """Save the filled Excel file"""
        self.wb.save(output_path)
        print(f"Saved to: {output_path}")

    def close(self):
        """Close the workbook"""
        self.wb.close()


def process_payslip(image_path: str, template_path: str, output_path: str, ocr_engine='easyocr'):
    """
    Process a single payslip image

    Args:
        image_path: Path to payslip image
        template_path: Path to Excel template
        output_path: Path for output Excel file
        ocr_engine: OCR engine to use ('easyocr' or 'tesseract')
    """
    print(f"Processing payslip: {image_path}")

    # Extract data
    extractor = PayslipExtractor(ocr_engine=ocr_engine)
    text = extractor.extract_text(image_path)

    print("\n=== Extracted Text ===")
    print(text)
    print("\n=== Parsing Data ===")

    data = extractor.parse_payslip(text)

    print("\nExtracted Data:")
    for key, value in data.items():
        if value and value != 0 and value != {} and value != []:
            print(f"  {key}: {value}")

    # Write to Excel
    print("\n=== Writing to Excel ===")
    excel_writer = ExcelWriter(template_path)
    excel_writer.fill_data(data)
    excel_writer.save(output_path)
    excel_writer.close()

    print("\nProcessing complete!")
    return data


def batch_process(image_dir: str, template_path: str, output_dir: str, ocr_engine='easyocr'):
    """
    Batch process multiple payslip images

    Args:
        image_dir: Directory containing payslip images
        template_path: Path to Excel template
        output_dir: Directory for output Excel files
        ocr_engine: OCR engine to use
    """
    os.makedirs(output_dir, exist_ok=True)

    image_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff'}
    image_files = []

    for file in os.listdir(image_dir):
        if Path(file).suffix.lower() in image_extensions:
            image_files.append(os.path.join(image_dir, file))

    print(f"Found {len(image_files)} images to process")

    for i, image_path in enumerate(image_files, 1):
        print(f"\n{'='*60}")
        print(f"Processing {i}/{len(image_files)}: {os.path.basename(image_path)}")
        print(f"{'='*60}")

        # Generate output filename
        image_name = Path(image_path).stem
        output_path = os.path.join(output_dir, f"{image_name}_output.xlsx")

        try:
            process_payslip(image_path, template_path, output_path, ocr_engine)
        except Exception as e:
            print(f"Error processing {image_path}: {e}")
            import traceback
            traceback.print_exc()


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='Process payslip images and fill Excel template')
    parser.add_argument('--image', type=str, help='Single image file to process')
    parser.add_argument('--batch', type=str, help='Directory of images to batch process')
    parser.add_argument('--template', type=str, default='SA - Empty.xlsx', help='Excel template file')
    parser.add_argument('--output', type=str, help='Output file/directory')
    parser.add_argument('--ocr', type=str, choices=['easyocr', 'tesseract'], default='easyocr',
                        help='OCR engine to use')

    args = parser.parse_args()

    if args.image:
        # Process single image
        output = args.output or args.image.replace('.jpg', '_output.xlsx').replace('.png', '_output.xlsx')
        process_payslip(args.image, args.template, output, args.ocr)

    elif args.batch:
        # Batch process
        output_dir = args.output or 'output'
        batch_process(args.batch, args.template, output_dir, args.ocr)

    else:
        # Default: process the test image
        process_payslip(
            'Image_20251027112221_133_4.jpg',
            'SA - Empty.xlsx',
            'output_test.xlsx',
            args.ocr
        )
