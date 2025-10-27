#!/usr/bin/env python3
"""
Demo script to fill Excel template with payslip data
Based on the actual payslip image data
"""

import openpyxl
from openpyxl.styles import Font, Alignment
import copy


def find_next_empty_row(sheet, start_row=30):
    """Find the next empty row in the worksheet"""
    for row in range(start_row, sheet.max_row + 10):
        if sheet[f'B{row}'].value is None or sheet[f'B{row}'].value == '':
            return row
    return start_row


def add_employee_to_excel(template_path, output_path, employee_data):
    """
    Add employee payslip data to Excel template

    Args:
        template_path: Path to Excel template
        output_path: Path for output Excel file
        employee_data: Dictionary containing employee payslip data
    """
    print(f"Loading template: {template_path}")
    wb = openpyxl.load_workbook(template_path)
    sheet = wb.active

    print(f"Template sheet: {sheet.title}")

    # Find the next empty row (after existing employees)
    # Based on the template structure, we'll add to "Factory - Office & Admin" section
    # which starts at row 30. Let's find the first empty row after that.

    # For demo, let's add to row 31 (right after the section header at row 30)
    target_row = 31

    # Check if row 31 already has data, if so find next empty
    if sheet[f'B{target_row}'].value:
        target_row = find_next_empty_row(sheet, target_row)

    print(f"Adding employee data to row {target_row}")

    # Get the last number used
    last_no = 8  # Based on the template, last number is 8
    for row in range(11, target_row):
        cell_value = sheet[f'A{row}'].value
        if cell_value and isinstance(cell_value, int):
            last_no = max(last_no, cell_value)

    employee_no = last_no + 1

    # Fill basic information
    sheet[f'A{target_row}'] = employee_no
    sheet[f'B{target_row}'] = employee_data.get('employee_code', '')
    sheet[f'C{target_row}'] = employee_data.get('employee_name', '')
    sheet[f'D{target_row}'] = employee_data.get('nric', '')
    sheet[f'E{target_row}'] = employee_data.get('basic_pay', 0)

    # Overtime data
    # 1.5 times overtime
    ot_15_hours = employee_data.get('ot_15_hours', 0)
    ot_15_amount = employee_data.get('ot_15_amount', 0)
    sheet[f'H{target_row}'] = ot_15_hours
    sheet[f'I{target_row}'] = ot_15_amount

    # Initialize other OT columns to 0
    sheet[f'F{target_row}'] = 0  # 1.0 Hrs
    sheet[f'G{target_row}'] = 0  # 1.0 Amount
    sheet[f'J{target_row}'] = 0  # 2.0 Hrs
    sheet[f'K{target_row}'] = 0  # 2.0 Amount
    sheet[f'L{target_row}'] = 0  # 3.0 Hrs
    sheet[f'M{target_row}'] = 0  # 3.0 Amount
    sheet[f'N{target_row}'] = 0  # Rest Day
    sheet[f'O{target_row}'] = 0  # Rest Day Amount
    sheet[f'P{target_row}'] = 0  # P. Holiday
    sheet[f'Q{target_row}'] = 0  # P. Holiday Amount

    # Total OT
    sheet[f'R{target_row}'] = ot_15_amount

    # Allowances
    sheet[f'S{target_row}'] = employee_data.get('child_care', 0)
    sheet[f'T{target_row}'] = employee_data.get('incentive', 0)
    sheet[f'U{target_row}'] = employee_data.get('cond_incent', 0)
    sheet[f'V{target_row}'] = employee_data.get('car_transp', 0)
    sheet[f'W{target_row}'] = employee_data.get('travelling_allw', 0)

    # Total Payable (Basic Pay + OT + Allowances)
    total_payable = employee_data.get('monthly_gross', 0)
    sheet[f'X{target_row}'] = total_payable

    # EPF
    epf_employer = employee_data.get('epf_employer', 0)
    epf_employee = employee_data.get('epf_employee', 0)
    sheet[f'Y{target_row}'] = epf_employer
    sheet[f'Z{target_row}'] = epf_employee
    sheet[f'AA{target_row}'] = epf_employer + epf_employee

    # SOCSO
    socso_employer = employee_data.get('socso_employer', 0)
    socso_employee = employee_data.get('socso_employee', 0)
    sheet[f'AB{target_row}'] = socso_employer
    sheet[f'AC{target_row}'] = socso_employee
    sheet[f'AD{target_row}'] = socso_employer + socso_employee

    # EIS
    eis_employer = employee_data.get('eis_employer', 0)
    eis_employee = employee_data.get('eis_employee', 0)
    sheet[f'AE{target_row}'] = eis_employer
    sheet[f'AF{target_row}'] = eis_employee
    sheet[f'AG{target_row}'] = eis_employer + eis_employee

    # Deductions
    sheet[f'AH{target_row}'] = employee_data.get('staff_loan', 0)
    sheet[f'AI{target_row}'] = employee_data.get('advance', 0)
    sheet[f'AJ{target_row}'] = employee_data.get('pcb', 0)

    # Nett Payable
    nett_payable = employee_data.get('nett_pay', 0)
    sheet[f'AK{target_row}'] = nett_payable

    print(f"\nEmployee data added successfully!")
    print(f"  Employee No: {employee_no}")
    print(f"  Code: {employee_data.get('employee_code', '')}")
    print(f"  Name: {employee_data.get('employee_name', '')}")
    print(f"  Basic Pay: RM {employee_data.get('basic_pay', 0):.2f}")
    print(f"  OT Amount: RM {ot_15_amount:.2f}")
    print(f"  Total Payable: RM {total_payable:.2f}")
    print(f"  Nett Pay: RM {nett_payable:.2f}")

    # Save the file
    wb.save(output_path)
    print(f"\nExcel file saved to: {output_path}")
    wb.close()


def main():
    """Main function - demo with actual payslip data"""

    # Data extracted from the payslip image (Image_20251027112221_133_4.jpg)
    employee_data = {
        'employee_code': 'Y0034',
        'employee_name': 'KYAW SWAR HTET',
        'nric': 'MD630258',
        'basic_pay': 1650.00,
        'working_days': 26.00,

        # Overtime
        'ot_15_hours': 29.00,
        'ot_15_rate': 13.5577,
        'ot_15_amount': 393.17,

        # Allowances
        'child_care': 0,
        'incentive': 0,
        'cond_incent': 0,
        'car_transp': 0,
        'travelling_allw': 230.00,  # LEADER ALLW

        # Totals
        'monthly_gross': 2273.17,

        # EPF
        'epf_employer': 0.00,
        'epf_employee': 0.00,

        # SOCSO
        'socso_employer': 39.35,
        'socso_employee': 11.25,

        # EIS
        'eis_employer': 0.00,
        'eis_employee': 0.00,

        # Deductions
        'staff_loan': 0.00,
        'advance': 0.00,
        'pcb': 0.00,

        # Net pay
        'nett_pay': 2261.92
    }

    print("="*60)
    print("Payslip Excel Filler - Demo")
    print("="*60)
    print(f"\nProcessing employee: {employee_data['employee_name']}")

    template_path = 'SA - Empty.xlsx'
    output_path = 'SA - Filled_Demo.xlsx'

    add_employee_to_excel(template_path, output_path, employee_data)

    print("\n" + "="*60)
    print("Demo completed successfully!")
    print("="*60)


if __name__ == '__main__':
    main()
