# Payslip Batch Processor

自动处理工资单图片，提取数据并填入Excel模板的批量处理工具。

## 功能特性

- 使用OCR技术自动识别工资单图片中的文本
- 智能解析工资单数据（员工信息、工资、津贴、扣除等）
- 自动填充Excel模板
- 支持单个文件或批量处理
- 支持多种OCR引擎（EasyOCR、Tesseract）

## 安装依赖

```bash
pip3 install -r requirements.txt
```

**注意：** EasyOCR需要下载约1GB的依赖包，如果网络较慢可以只安装Tesseract：

```bash
pip3 install openpyxl Pillow pytesseract
```

## 使用方法

### 1. 处理单个图片

```bash
python3 payslip_processor.py --image Image_20251027112221_133_4.jpg --template "SA - Empty.xlsx" --output output.xlsx
```

### 2. 批量处理多个图片

```bash
python3 payslip_processor.py --batch ./images --template "SA - Empty.xlsx" --output ./output
```

### 3. 使用默认测试

```bash
python3 payslip_processor.py
```

这将处理默认的测试图片 `Image_20251027112221_133_4.jpg`

### 4. 选择OCR引擎

```bash
# 使用EasyOCR（推荐，准确度高）
python3 payslip_processor.py --image test.jpg --ocr easyocr

# 使用Tesseract（速度快，体积小）
python3 payslip_processor.py --image test.jpg --ocr tesseract
```

## 参数说明

- `--image`: 单个图片文件路径
- `--batch`: 包含多个图片的文件夹路径
- `--template`: Excel模板文件路径（默认：SA - Empty.xlsx）
- `--output`: 输出文件/文件夹路径
- `--ocr`: OCR引擎选择（easyocr 或 tesseract，默认：easyocr）

## 支持的数据字段

脚本可以提取以下字段：

- 员工信息：员工编号、姓名、身份证号
- 工资信息：基本工资、工作天数、基本薪资
- 津贴：各种津贴（如领导津贴）
- 加班：加班类型、费率、小时数、金额
- 扣除：EPF、SOCSO、EIS（员工和雇主部分）
- 总计：月总收入、净工资

## 文件结构

```
payslip-automation/
├── payslip_processor.py      # 主处理脚本
├── check_template.py          # Excel模板检查工具
├── requirements.txt           # Python依赖
├── README.md                  # 本文件
├── SA - Empty.xlsx            # Excel模板
├── Image_20251027112221_133_4.jpg  # 测试图片
└── output/                    # 输出文件夹（自动创建）
```

## 自定义Excel映射

如果需要调整Excel模板的字段映射，请编辑 `payslip_processor.py` 中 `ExcelWriter.fill_data()` 方法的 `mappings` 字典。

## 常见问题

### Q: OCR识别不准确怎么办？

A:
1. 确保图片清晰，分辨率足够高
2. 尝试使用EasyOCR引擎（准确度更高）
3. 检查提取的文本输出，手动调整解析规则

### Q: Excel模板不匹配？

A:
1. 运行 `python3 check_template.py` 查看模板结构
2. 根据模板调整 `payslip_processor.py` 中的字段映射

## 开发说明

- `PayslipExtractor` 类：负责OCR和数据提取
- `ExcelWriter` 类：负责Excel文件操作
- `process_payslip()` 函数：处理单个文件
- `batch_process()` 函数：批量处理

## 许可证

MIT License
