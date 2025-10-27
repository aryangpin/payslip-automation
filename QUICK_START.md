# Payslip批量处理工具 - 快速开始

## 📋 文件说明

- **process_payslip.py** - 主处理脚本（根据Staff Code填充数据）
- **SA - Empty.xlsx** - Excel模板
- **SA - Updated.xlsx** - 生成的输出文件 ✅

## 🚀 快速使用

### 方法1：直接运行（已配置好测试数据）

```bash
python3 process_payslip.py
```

生成文件：`SA - Updated.xlsx`

---

### 方法2：处理你自己的payslip数据

编辑 `process_payslip.py`，修改第186行开始的数据：

```python
payslip_data = {
    'staff_code': 'AF0001',  # 修改为你的Staff Code
    'basic_pay': 1650.00,    # 修改为实际基本工资
    'ot_total': 393.17,      # 修改为加班费
    'incentive': 0.00,       # 修改为奖金
    'epf_employer': 0.00,    # EPF雇主部分
    'epf_employee': 0.00,    # EPF员工部分
    'socso_employer': 39.35, # SOCSO雇主部分
    'socso_employee': 11.25, # SOCSO员工部分
    'eis_employer': 0.00,    # EIS雇主部分
    'eis_employee': 0.00,    # EIS员工部分
    # ... 更多字段
}
```

然后运行：
```bash
python3 process_payslip.py
```

---

## 📊 Excel模板中的员工列表

当前可用的Staff Code（在SA - Empty.xlsx中）：

| Staff Code | 员工姓名 |
|------------|----------|
| A001 | Wong Thim Fatt |
| A003 | Teo Wei Hao |
| A004 | Chee Kang Hwai |
| A007 | Norarsikin Binti Ahian |
| A005 | Lim Yeong Kern |
| AA01 | Lim Weng Khim |
| AA02 | Hing Soo Cheng |
| AA03 | Hing Soo Yee |
| AE02 | Kwek Chee Yang |
| **AF0001** | **Hamdan Bin Kassim** ✅ (已演示) |
| AF0002 | Aisha Bin Kassim |
| AF0003 | Mazri Bin Said |
| AF0005 | Mohamed Anuar Bin Abdullah Sani |
| AF0006 | Chan Yew Kwan |
| AF0007 | Nor Azimah Binti Razali |
| AF0008 | Nureen Batrisyia Binti Nik Khairul Hizam |
| AF0010 | Muhammad Akmal Fariz Bin Mazri |

---

## 📝 填充的数据字段

脚本会自动填充以下Excel列：

| 数据项 | Excel列 | 说明 |
|--------|---------|------|
| Basic Pay | E | 基本工资 |
| OT Hours (1.5x) | H | 1.5倍加班小时 |
| OT Amount (1.5x) | I | 1.5倍加班费 |
| OT Total | R | 总加班费 |
| Incentive | T | 奖金 |
| Allowance | W | 津贴 |
| Total Payable | X | 月总收入 |
| EPF Employer | Y | EPF雇主部分 |
| EPF Employee | Z | EPF员工部分 |
| EPF Total | AA | EPF总计 |
| SOCSO Employer | AB | SOCSO雇主部分 |
| SOCSO Employee | AC | SOCSO员工部分 |
| SOCSO Total | AD | SOCSO总计 |
| EIS Employer | AE | EIS雇主部分 |
| EIS Employee | AF | EIS员工部分 |
| EIS Total | AG | EIS总计 |
| Nett Pay | AK | 净工资 |

---

## ✅ 已完成示例

演示数据（AF0001 - Hamdan Bin Kassim）：
- ✓ Basic Pay: RM 1650.00
- ✓ OT (1.5x): RM 393.17 (29 hours)
- ✓ Incentive: RM 0.00
- ✓ Allowance: RM 230.00
- ✓ EPF: RM 0.00 (Employer + Employee)
- ✓ SOCSO: RM 50.60 (Employer 39.35 + Employee 11.25)
- ✓ EIS: RM 0.00 (Employer + Employee)
- ✓ Nett Pay: RM 2261.92

**生成文件：SA - Updated.xlsx** ✅

---

## 🔄 批量处理多个员工

如果需要处理多个员工，修改脚本主函数：

```python
def main():
    employees_data = [
        {
            'staff_code': 'AF0001',
            'basic_pay': 1650.00,
            # ... 其他数据
        },
        {
            'staff_code': 'AF0002',
            'basic_pay': 1800.00,
            # ... 其他数据
        },
        # 添加更多员工...
    ]

    for emp_data in employees_data:
        update_payslip_data('SA - Empty.xlsx', 'SA - Updated.xlsx', emp_data)
```

---

## 💡 提示

1. **找不到Staff Code？**
   - 运行 `python3 check_template.py` 查看Excel中所有员工
   - 确保Staff Code完全匹配（大小写敏感）

2. **修改输出文件名？**
   - 编辑脚本第223行：`output_path = 'SA - Updated.xlsx'`

3. **需要处理新的payslip图片？**
   - 根据图片中的数据修改 `payslip_data` 字典
   - 确保Staff Code在Excel模板中存在

---

## 📦 下载文件

生成的文件：
- **SA - Updated.xlsx** - 可直接下载使用

---

## 🎯 下一步

1. 打开 `SA - Updated.xlsx` 查看结果
2. 根据你的新payslip图片修改数据
3. 重新运行脚本生成新文件

需要帮助？查看脚本中的详细注释！
