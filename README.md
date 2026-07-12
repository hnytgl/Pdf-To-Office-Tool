# PDF 转 Office 工具

一个轻量级命令行工具，可在 Windows、macOS 和 Linux 上批量把 PDF 转换为 Word、PowerPoint 或 Excel。

## 功能

- 单个文件或整个目录批量转换
- 支持递归扫描、页码范围和覆盖输出
- Word：尽量保留文字、图片、表格和页面布局
- PowerPoint：每页生成一张高保真幻灯片
- Excel：提取可编辑表格；没有表格时按行保存页面文字
- 带红色印章的表格：识别并忽略常见红色矢量印章，避免印章圆环和文字被误判为表格线或业务数据
- Excel 数据保护：PDF 中以 `=` 开头的内容按普通文本保存，不会被 Excel 当作公式执行
- 原子写入：转换完成后才替换目标文件，失败时不留下半成品

## 安装

需要 Python 3.10 或更高版本。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e .
```

## 使用

```powershell
# 转 Word、PowerPoint 或 Excel
pdf2office .\example.pdf --format word
pdf2office .\example.pdf --format ppt
pdf2office .\example.pdf --format excel

# 指定输出文件
pdf2office .\example.pdf --format excel -o .\example.xlsx

# 递归转换目录中的 PDF
pdf2office .\pdfs --format excel -o .\excel输出 --recursive

# 只转换第 1 至第 3 页
pdf2office .\example.pdf --format excel --start 1 --end 3

# 覆盖已有文件
pdf2office .\pdfs --format excel -o .\excel输出 --overwrite
```

## 带章表格的 Excel 转换

Excel 转换会同时尝试严格表格线、容错表格线和无边框文字对齐三种模型，并对原页面与“去除疑似红色印章对象”的页面分别提取。程序根据有效单元格、行列结构、数据密度和碎片数量选择质量最高的结果。

生成的工作簿包含一个隐藏的 `_转换报告` 工作表，记录每页采用的提取策略、忽略的疑似印章对象数和质量评分，便于审计。业务数据仍按 PDF 原文保存，不会把数字字符串擅自改成浮点数、日期或公式。

需要注意：

- 对于带文字层的电子 PDF，图片形式的印章通常不会干扰表格数据；红色矢量印章会被识别并过滤。
- 纯扫描件没有可提取的文字层，需要先进行 OCR。本项目目前不会猜测无法识别的单元格内容。
- 印章完全遮住纸面文字、字体编码损坏、手写内容或极复杂的跨页合并表格，仍可能需要人工复核。
- “质量评分”用于选择候选表格，不代表数据正确率。重要财务或合同数据应与原 PDF 核对。

## 开发与测试

```powershell
python -m pip install -e .[dev]
python -m pytest -q
```

测试包含一个由程序生成的回归样例：红色矢量印章横跨表格边框和数据单元格，转换后会核对表头、金额、文本公式和隐藏审计信息。
