# PDF 转 Office 工具

这是一个轻量级 PDF 转 Office 命令行工具，适合在 Windows、macOS 和 Linux 上批量把 PDF 转为 Word、PPT 和 Excel 文件。

## 功能

- 单个 PDF 转 Word、PPT、Excel。
- 整个目录批量转换。
- 支持递归扫描子目录。
- 支持指定页码范围。
- 支持覆盖已有文件。
- 输出清晰的成功、跳过和失败信息。

不同格式的转换策略：

- Word：使用 `pdf2docx`，尽量保留文字、图片、表格和页面布局。
- PPT：把 PDF 每一页渲染成图片，并按页生成 PPT 幻灯片，适合汇报和展示。
- Excel：优先提取 PDF 表格；如果没有表格，则按页写入文本内容，适合后续整理。

## 安装

建议使用 Python 3.10 或更高版本。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e .
```

## 使用方法

转换为 Word：

```powershell
pdf2office .\example.pdf --format word
```

转换为 PPT：

```powershell
pdf2office .\example.pdf --format ppt
```

转换为 Excel：

```powershell
pdf2office .\example.pdf --format excel
```

指定输出文件：

```powershell
pdf2office .\example.pdf --format word -o .\example.docx
```

批量转换目录中的 PDF：

```powershell
pdf2office .\pdfs --format ppt -o .\ppt输出
```

递归转换子目录：

```powershell
pdf2office .\pdfs --format excel -o .\excel输出 --recursive
```

只转换第 1 到第 3 页：

```powershell
pdf2office .\example.pdf --format word --start 1 --end 3
```

覆盖已有文件：

```powershell
pdf2office .\pdfs --format word -o .\word输出 --overwrite
```

## 注意事项

- 如果 PDF 是扫描件图片，Word 和 Excel 转换通常不能自动识别图片中的文字。扫描件需要 OCR 工具配合。
- PPT 转换会把每一页当作图片放入幻灯片，展示效果稳定，但文字不可直接编辑。
- Excel 转换适合表格型 PDF；复杂版式、跨页表格或扫描件可能需要人工整理。
- 复杂排版、特殊字体或加密 PDF 可能无法完全保持原样。
- 页码参数按普通用户习惯从 1 开始，例如 `--start 1 --end 3` 表示转换第 1 到第 3 页。

## 开发测试

```powershell
python -m pip install -e .[dev]
python -m pytest -q
```
