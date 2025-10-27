# 如何开始新的 Payslip 处理会话

## 🚀 快速启动指令

### 选项 1: 最简单（推荐）
```
请查看 prompt.txt，然后帮我处理这个 payslip 图片
```
然后上传你的 payslip 图片即可。

---

### 选项 2: 更详细
```
你好！请执行以下操作：

1. 读取 /home/user/payslip-automation/prompt.txt 了解项目需求
2. 查看我上传的 payslip 图片
3. 提取 Staff Code 和相关数据（Basic Pay, OT, Incentive, EPF, SOCSO, EIS）
4. 不要修改 Staff Name 和 NRIC
5. 更新 process_payslip.py 中的数据
6. 运行脚本生成 SA - Updated.xlsx
7. Git commit + push
8. 返回 GitHub 下载链接
```

---

### 选项 3: 超简短版
```
按 prompt.txt 处理 payslip
```

---

## 📝 完整会话示例

**你可以这样说：**

> 你好！请查看 prompt.txt，然后帮我处理这个 payslip 图片（Image_xxxxx.jpg）

**AI 会自动：**
1. ✅ 读取 prompt.txt 了解所有规则
2. ✅ 查看图片提取数据
3. ✅ 验证 Staff Code
4. ✅ 填充 Excel（不修改 Staff Name 和 NRIC）
5. ✅ Git commit + push
6. ✅ 返回 GitHub 链接

---

## 🎯 关键词提示

只要你的消息中包含以下关键词，AI 就会知道要做什么：

- **"prompt.txt"** - AI 会读取项目需求
- **"payslip"** - AI 知道这是工资单处理任务
- **"处理"** 或 **"process"** - 开始执行流程

---

## ✅ 推荐用法

### 方式 A：一句话搞定
```
按 prompt.txt 处理这个 payslip
```

### 方式 B：更明确
```
请根据 prompt.txt 的流程，处理我上传的 payslip 图片，
然后生成 Excel 并给我 GitHub 下载链接
```

### 方式 C：简单直接
```
处理 payslip：[图片文件名]
参考：prompt.txt
```

---

## 📂 项目位置

- **项目路径**: `/home/user/payslip-automation/`
- **需求文档**: `prompt.txt`
- **处理脚本**: `process_payslip.py`
- **输出文件**: `SA - Updated.xlsx`

---

## 💡 提示

- 你只需要提到 **"prompt.txt"** 和 **"payslip"**
- AI 会自动读取所有规则和流程
- 无需重复说明 Staff Name 和 NRIC 不修改等细节
- 一切都已经写在 prompt.txt 中了！

---

**记住：只要说 "按 prompt.txt 处理 payslip" 就可以了！** 🎉
