#!/usr/bin/env python3
"""
Payslip处理工具 - 根据Staff Code填充Excel
提取数据：Basic Pay, OT, Incentive, EPF, SOCSO, EIS
"""

import openpyxl
from pathlib import Path


def find_employee_row_by_staff_code(sheet, staff_code):
    """
    在Excel中根据Staff Code查找员工所在行

    Args:
        sheet: Excel工作表
        staff_code: 员工代码（如Y0034）

    Returns:
        行号，如果找不到返回None
    """
    print(f"正在查找Staff Code: {staff_code}")

    # 遍历B列查找staff code
    for row in range(1, sheet.max_row + 1):
        cell_value = sheet[f'B{row}'].value
        if cell_value and str(cell_value).strip().upper() == str(staff_code).strip().upper():
            print(f"✓ 找到员工：{staff_code} 在第{row}行")
            return row

    print(f"✗ 未找到Staff Code: {staff_code}")
    return None


def update_payslip_data(template_path, output_path, payslip_data):
    """
    更新Excel中的payslip数据

    Args:
        template_path: Excel模板路径
        output_path: 输出文件路径
        payslip_data: 包含员工薪资数据的字典
    """
    print(f"\n{'='*60}")
    print(f"处理Payslip数据")
    print(f"{'='*60}\n")

    # 加载Excel
    print(f"加载Excel模板: {template_path}")
    wb = openpyxl.load_workbook(template_path)
    sheet = wb.active
    print(f"工作表: {sheet.title}\n")

    # 查找员工行
    staff_code = payslip_data.get('staff_code')
    row = find_employee_row_by_staff_code(sheet, staff_code)

    if row is None:
        print(f"\n错误：未找到Staff Code '{staff_code}'")
        print("请检查：")
        print("1. Staff Code是否正确")
        print("2. Excel模板中B列是否包含该员工")
        wb.close()
        return False

    print(f"\n开始更新数据到第{row}行...")
    print("-" * 60)

    # 更新Basic Pay (E列)
    if 'basic_pay' in payslip_data:
        old_value = sheet[f'E{row}'].value
        sheet[f'E{row}'] = payslip_data['basic_pay']
        print(f"Basic Pay (E{row}):      {old_value} → {payslip_data['basic_pay']}")

    # 更新OT (加班费)
    # 根据Excel模板，OT可能在多列：F-R列用于不同倍率的加班
    # 这里假设总OT金额在R列 (Total OT)
    if 'ot_total' in payslip_data:
        old_value = sheet[f'R{row}'].value
        sheet[f'R{row}'] = payslip_data['ot_total']
        print(f"OT Total (R{row}):       {old_value} → {payslip_data['ot_total']}")

    # 如果有分项OT数据（1.5倍加班）
    if 'ot_15_hours' in payslip_data:
        sheet[f'H{row}'] = payslip_data['ot_15_hours']
        print(f"OT 1.5x Hours (H{row}):  → {payslip_data['ot_15_hours']}")

    if 'ot_15_amount' in payslip_data:
        sheet[f'I{row}'] = payslip_data['ot_15_amount']
        print(f"OT 1.5x Amount (I{row}): → {payslip_data['ot_15_amount']}")

    # 更新Incentive (T列)
    if 'incentive' in payslip_data:
        old_value = sheet[f'T{row}'].value
        sheet[f'T{row}'] = payslip_data['incentive']
        print(f"Incentive (T{row}):      {old_value} → {payslip_data['incentive']}")

    # 更新Allowance (如果有)
    if 'allowance' in payslip_data:
        old_value = sheet[f'W{row}'].value
        sheet[f'W{row}'] = payslip_data['allowance']
        print(f"Allowance (W{row}):      {old_value} → {payslip_data['allowance']}")

    # 更新Total Payable (X列) - 如果提供了
    if 'total_payable' in payslip_data:
        old_value = sheet[f'X{row}'].value
        sheet[f'X{row}'] = payslip_data['total_payable']
        print(f"Total Payable (X{row}):  {old_value} → {payslip_data['total_payable']}")

    # 更新EPF
    if 'epf_employer' in payslip_data:
        old_value = sheet[f'Y{row}'].value
        sheet[f'Y{row}'] = payslip_data['epf_employer']
        print(f"EPF Employer (Y{row}):   {old_value} → {payslip_data['epf_employer']}")

    if 'epf_employee' in payslip_data:
        old_value = sheet[f'Z{row}'].value
        sheet[f'Z{row}'] = payslip_data['epf_employee']
        print(f"EPF Employee (Z{row}):   {old_value} → {payslip_data['epf_employee']}")

    # EPF Total (AA列)
    if 'epf_employer' in payslip_data and 'epf_employee' in payslip_data:
        sheet[f'AA{row}'] = payslip_data['epf_employer'] + payslip_data['epf_employee']
        print(f"EPF Total (AA{row}):     → {sheet[f'AA{row}'].value}")

    # 更新SOCSO
    if 'socso_employer' in payslip_data:
        old_value = sheet[f'AB{row}'].value
        sheet[f'AB{row}'] = payslip_data['socso_employer']
        print(f"SOCSO Employer (AB{row}): {old_value} → {payslip_data['socso_employer']}")

    if 'socso_employee' in payslip_data:
        old_value = sheet[f'AC{row}'].value
        sheet[f'AC{row}'] = payslip_data['socso_employee']
        print(f"SOCSO Employee (AC{row}): {old_value} → {payslip_data['socso_employee']}")

    # SOCSO Total (AD列)
    if 'socso_employer' in payslip_data and 'socso_employee' in payslip_data:
        sheet[f'AD{row}'] = payslip_data['socso_employer'] + payslip_data['socso_employee']
        print(f"SOCSO Total (AD{row}):   → {sheet[f'AD{row}'].value}")

    # 更新EIS
    if 'eis_employer' in payslip_data:
        old_value = sheet[f'AE{row}'].value
        sheet[f'AE{row}'] = payslip_data['eis_employer']
        print(f"EIS Employer (AE{row}):  {old_value} → {payslip_data['eis_employer']}")

    if 'eis_employee' in payslip_data:
        old_value = sheet[f'AF{row}'].value
        sheet[f'AF{row}'] = payslip_data['eis_employee']
        print(f"EIS Employee (AF{row}):  {old_value} → {payslip_data['eis_employee']}")

    # EIS Total (AG列)
    if 'eis_employer' in payslip_data and 'eis_employee' in payslip_data:
        sheet[f'AG{row}'] = payslip_data['eis_employer'] + payslip_data['eis_employee']
        print(f"EIS Total (AG{row}):     → {sheet[f'AG{row}'].value}")

    # 更新Nett Pay (AK列) - 如果提供了
    if 'nett_pay' in payslip_data:
        old_value = sheet[f'AK{row}'].value
        sheet[f'AK{row}'] = payslip_data['nett_pay']
        print(f"Nett Pay (AK{row}):      {old_value} → {payslip_data['nett_pay']}")

    print("-" * 60)

    # 保存文件
    print(f"\n保存文件到: {output_path}")
    wb.save(output_path)
    wb.close()

    print(f"\n{'='*60}")
    print(f"✓ 处理完成！")
    print(f"{'='*60}\n")

    return True


def main():
    """主函数 - 处理payslip数据"""

    # 从图片Image_20251027112221_133_4.jpg提取的数据
    # 你可以修改这里的数据来处理其他员工
    # 注意：请使用Excel模板中实际存在的Staff Code

    # 示例数据 - 使用Excel中存在的员工AF0001
    payslip_data = {
        'staff_code': 'AF0001',  # Staff Code - 必需（Excel中的Hamdan Bin Kassim）

        # Basic Pay
        'basic_pay': 1650.00,

        # OT (加班)
        'ot_15_hours': 29.00,     # 1.5倍加班小时
        'ot_15_amount': 393.17,   # 1.5倍加班费
        'ot_total': 393.17,       # 总加班费

        # Incentive (奖金)
        'incentive': 0.00,

        # Allowance (津贴)
        'allowance': 230.00,      # LEADER ALLW

        # Total
        'total_payable': 2273.17,  # Monthly Gross

        # EPF
        'epf_employer': 0.00,     # EPF 雇主部分
        'epf_employee': 0.00,     # EPF 员工部分

        # SOCSO
        'socso_employer': 39.35,  # SOCSO 雇主部分
        'socso_employee': 11.25,  # SOCSO 员工部分

        # EIS
        'eis_employer': 0.00,     # EIS 雇主部分
        'eis_employee': 0.00,     # EIS 员工部分

        # Nett Pay
        'nett_pay': 2261.92,
    }

    # 文件路径
    template_path = 'SA - Empty.xlsx'
    output_path = 'SA - Updated.xlsx'

    # 检查模板文件是否存在
    if not Path(template_path).exists():
        print(f"错误：找不到模板文件 '{template_path}'")
        return

    # 处理数据
    success = update_payslip_data(template_path, output_path, payslip_data)

    if success:
        print(f"✓ 成功！请打开文件查看: {output_path}")
        print(f"\n更新的数据：")
        print(f"  Staff Code:    {payslip_data['staff_code']}")
        print(f"  Basic Pay:     RM {payslip_data['basic_pay']:.2f}")
        print(f"  OT:            RM {payslip_data['ot_total']:.2f}")
        print(f"  Incentive:     RM {payslip_data['incentive']:.2f}")
        print(f"  EPF (E+E):     RM {payslip_data['epf_employer'] + payslip_data['epf_employee']:.2f}")
        print(f"  SOCSO (E+E):   RM {payslip_data['socso_employer'] + payslip_data['socso_employee']:.2f}")
        print(f"  EIS (E+E):     RM {payslip_data['eis_employer'] + payslip_data['eis_employee']:.2f}")
        print(f"  Nett Pay:      RM {payslip_data['nett_pay']:.2f}")


if __name__ == '__main__':
    main()
