# Payslip批量处理工具 - 使用指南

## 简介

这是一个自动化工资单处理工具，可以：
1. 从工资单图片中提取数据
2. 自动填充到Excel工资单模板中
3. 支持批量处理多张工资单

## 文件说明

### 核心文件
- `payslip_processor.py` - 完整版处理脚本（支持OCR图片识别）
- `demo_fill_excel.py` - 演示脚本（手动数据填充示例）
- `check_template.py` - Excel模板结构查看工具
- `requirements.txt` - Python依赖包列表
- `README.md` - 英文文档
- `USAGE_CN.md` - 中文使用指南（本文件）

### 测试文件
- `Image_20251027112221_133_4.jpg` - 测试用工资单图片
- `SA - Empty.xlsx` - Excel模板文件
- `SA - Filled_Demo.xlsx` - 演示脚本生成的输出文件

## 快速开始

### 1. 安装依赖

**基础依赖（必须）：**
```bash
pip3 install openpyxl Pillow
```

**OCR支持（可选，用于自动图片识别）：**
```bash
# 轻量级方案（推荐新手）
pip3 install pytesseract

# 高精度方案（需要约1.5GB空间）
pip3 install easyocr
```

### 2. 运行演示

最简单的方式是运行演示脚本：

```bash
python3 demo_fill_excel.py
```

这将：
- 使用预定义的测试数据
- 填充到 `SA - Empty.xlsx` 模板
- 生成 `SA - Filled_Demo.xlsx` 输出文件

### 3. 查看结果

打开生成的 `SA - Filled_Demo.xlsx` 文件，您会看到：
- 员工信息：Y0034 - KYAW SWAR HTET
- 基本工资：RM 1650.00
- 加班费：RM 393.17
- 津贴：RM 230.00
- 净工资：RM 2261.92

## 自定义使用

### 方式1：修改演示脚本

编辑 `demo_fill_excel.py`，修改 `employee_data` 字典：

```python
employee_data = {
    'employee_code': 'Y0035',  # 修改员工编号
    'employee_name': '张三',    # 修改员工姓名
    'nric': 'ABC123456',       # 修改身份证号
    'basic_pay': 2000.00,      # 修改基本工资
    # ... 其他字段
}
```

然后运行：
```bash
python3 demo_fill_excel.py
```

### 方式2：使用OCR自动处理（需要安装EasyOCR）

处理单张图片：
```bash
python3 payslip_processor.py --image your_payslip.jpg --output output.xlsx
```

批量处理多张图片：
```bash
# 将所有工资单图片放在 images 文件夹
mkdir images
# 将图片复制到 images/ 文件夹

# 批量处理
python3 payslip_processor.py --batch images --output output_files
```

## 数据字段说明

工资单中可以提取和填充的字段：

| 字段名 | 说明 | 示例 |
|--------|------|------|
| employee_code | 员工编号 | Y0034 |
| employee_name | 员工姓名 | KYAW SWAR HTET |
| nric | 身份证号 | MD630258 |
| basic_pay | 基本工资 | 1650.00 |
| ot_15_hours | 1.5倍加班小时数 | 29.00 |
| ot_15_amount | 1.5倍加班费 | 393.17 |
| travelling_allw | 差旅/津贴 | 230.00 |
| monthly_gross | 月总收入 | 2273.17 |
| epf_employer | EPF雇主部分 | 0.00 |
| epf_employee | EPF员工部分 | 0.00 |
| socso_employer | SOCSO雇主部分 | 39.35 |
| socso_employee | SOCSO员工部分 | 11.25 |
| nett_pay | 净工资 | 2261.92 |

## 常见问题

### Q1: 如何查看Excel模板结构？

运行：
```bash
python3 check_template.py
```

这会显示Excel模板的完整结构，包括所有列和行的信息。

### Q2: OCR识别不准确怎么办？

1. 确保图片清晰，分辨率足够
2. 使用EasyOCR而不是Tesseract（准确度更高）
3. 检查提取的文本，手动调整 `payslip_processor.py` 中的正则表达式

### Q3: 如何批量处理多个工资单？

```bash
# 1. 创建图片文件夹
mkdir payslip_images

# 2. 将所有工资单图片放入文件夹
# cp your_payslips/*.jpg payslip_images/

# 3. 批量处理
python3 payslip_processor.py --batch payslip_images --output processed_files
```

### Q4: Excel文件被占用无法保存？

确保在运行脚本前关闭所有打开的Excel文件。

### Q5: 能否处理不同格式的工资单？

需要修改 `payslip_processor.py` 中的 `parse_payslip()` 函数，根据您的工资单格式调整数据提取规则。

## 工作流程

```
工资单图片 → OCR识别 → 数据提取 → 解析字段 → 填充Excel → 生成输出文件
```

详细步骤：

1. **图片读取**：使用Pillow或EasyOCR读取图片
2. **文本提取**：通过OCR引擎提取文本
3. **数据解析**：使用正则表达式匹配和提取关键字段
4. **Excel操作**：使用openpyxl库打开模板并填充数据
5. **保存文件**：生成新的Excel文件

## 进阶使用

### 自定义字段映射

如果您的Excel模板列不同，修改 `demo_fill_excel.py` 中的列映射：

```python
# 示例：如果您的模板中"姓名"在D列而不是C列
sheet[f'D{target_row}'] = employee_data.get('employee_name', '')
```

### 添加新字段

1. 在 `employee_data` 字典中添加新字段
2. 在 `add_employee_to_excel()` 函数中添加对应的Excel单元格赋值

```python
# 添加新字段：奖金
employee_data['bonus'] = 500.00

# 在函数中填充（假设奖金在AZ列）
sheet[f'AZ{target_row}'] = employee_data.get('bonus', 0)
```

## 技术支持

如有问题：
1. 检查 `README.md` 获取更多技术细节
2. 查看脚本中的注释和文档字符串
3. 确保所有依赖包已正确安装

## 许可证

MIT License - 可自由使用和修改
