# PDF 到 Word 转换器 - 智能文字识别版

这是一个功能强大的 Python 程序，可以**智能识别和提取PDF中的文字**并转换为Word文档，支持普通PDF和扫描版PDF，并提供编辑和关键词搜索功能。

## 功能特性

✅ **智能文字识别** - 自动检测PDF类型，智能提取文字内容  
✅ **PDF 转 Word** - 将PDF文件（包括扫描版）转换为可编辑的Word文档  
✅ **OCR文字识别** - 支持扫描版PDF的文字识别（中文、英文）  
✅ **关键词搜索** - 在Word文档中搜索特定关键词  
✅ **关键词高亮** - 自动高亮显示文档中的关键词  
✅ **文本替换** - 批量替换文档中的文本内容  
✅ **添加内容** - 在文档末尾添加新的文本  
✅ **文档信息** - 查看文档的段落数、字符数等统计信息

### 基础功能（必需）

```bash
pip install pdf2docx python-docx PyPDF2
```

### OCR功能（可选 - 用于扫描版PDF）

如果需要识别扫描版PDF中的文字，需要额外安装：

```bash
pip install pdf2image pytesseract Pillow
```

**并且**需要安装Tesseract OCR引擎和Poppler工具。

详细的OCR安装步骤请查看：**[OCR_SETUP.md](OCR_SETUP.md)**

### 一键安装所有依赖

```bash
pip install -r requirements.txt
```

**注意**：即使不安装OCR功能，程序仍可处理普通文本型PDF。 2. 扫描版PDF（图片PDF）
- 特点：扫描文档、图片转的PDF
- 处理方式：使用OCR技术识别图片中的文字
- 准确度：90%+（取决于图片清晰度）
- 速度：较慢（需要逐页识别）
- 语言支持：中文、英文

**程序会自动检测PDF类型并选择最佳处理方式！**

## 安装依赖

在运行程序之前，需要安装必要的 Python 库：

```bash
pip install -r requirements.txt
```

或者手动安装：

```bash
pip install pdf2docx python-docx PyPDF2
```

## 使用方法

### 方式一：图形界面（推荐）

运行图形界面版本，通过可视化界面操作：

```bash
python pdf_to_word_gui.py
```

图形界面提供：
- 📁 可视化文件选择器
- 🎨 清晰的选项卡布局
- 📊 实时操作反馈
- 💡 友好的中文提示

### 方式二：交互式菜单

直接运行命令行程序，按照菜单提示操作：

```bash（自动检测PDF类型）
word_file = converter.convert_pdf_to_word('input.pdf', 'output.docx')

# 强制使用OCR识别文字（适用于扫描版PDF）
word_file = converter.convert_pdf_to_word('scan.pdf', use_ocr=True)

# 检查PDF是否包含文字
has_text = converter.check_pdf_has_text('document.pdf')
if has_text:
    print("这是文本型PDF")
else:
    print("这是扫描版PDF，需要OCR识别"
```

### 方式三：作为模块使用

在您的 Python 代码中导入使用：

```python
from pdf_to_word_converter import PDFToWordConverter

# 创建转换器实例
converter = PDFToWordConverter()

# 1. 转换 PDF 到 Word
word_file = converter.convert_pdf_to_word('input.pdf', 'output.docx')

# 2. 搜索关键词
results = converter.search_keyword('output.docx', '关键词')
for idx, text in results:
    print(f"段落 {idx}: {text}")

# 3. 高亮关键词
converter.highlight_keyword('output.docx', '重要', 'output_highlighted.docx')

# 4. 替换文本
converter.replace_text('output.docx', '旧文本', '新文本', 'output_edited.docx')

# 5. 添加文本（文字识别）
- 支持将PDF文件转换为 .docx 格式
- **智能识别PDF类型**：自动检测是否需要OCR
- **文本型PDF**：快速直接提取文字
- **扫描版PDF**：使用OCR识别图片中的文字
- 支持中文、英文等多语言识别cument('output.docx', '这是新添加的内容')

# 6. 查看文档信息
info = converter.get_document_info('output.docx')
print(info)
```

## 功能说明

### 1. PDF 转 Word
- 支持将 PDF 文件转换为 .docx 格式
- 自动保留原始格式和布局
- 支持自定义输出路径

### 2. 关键词搜索
- 不区分大小写搜索
- 返回包含关键词的所有段落
- 显示段落索引和内容预览

### 3. 关键词高亮
- 将关键词用黄色背景高亮显示
- 便于快速定位重要内容
- 生成新文件保留原文档

### 4. 文本替换
- 批量替换文档中的指定文本
- 支持段落和表格内容
- 显示替换次数统计

### 5. 添加文本
- 在文档末尾添加新段落
- 保持原有格式

### 6. 文档信息
- 查看段落数量
- 查看表格数量
- 查看总字符数类型说明**
- **文本型PDF**：可直接复制文字的PDF，转换快速准确
- **扫描版PDF**：扫描文档或图片生成的PDF，需要OCR识别文字
- 程序会自动检测PDF类型并选择最佳处理方式

⚠️ **OCR 识别准确度**
- 识别准确度取决于图片清晰度（建议300 DPI以上）
- 确保文字清晰、无倾斜、无模糊
- 支持中文简体和英文（可扩展其他语言）
- 复杂排版可能影响识别效果

⚠️ **PDF 格式限制**
- 建议使用文本型 PDF（非扫描版）
- 扫描版 PDF 需要先进行 OCR 处理

⚠️ **文：文字识别和转换
converter = PDFToWordConverter()

# 示例1: 转换普通PDF
word_file = converter.convert_pdf_to_word('报告.pdf')

# 示例2: 转换扫描版PDF（自动OCR）
word_file = converter.convert_pdf_to_word('扫描文档.pdf')

# 示例3: 检查PDF类型
if converter.check_pdf_has_text('document.pdf'):
    print("这是文本型PDF，可以直接提取")
    converter.convert_pdf_to_word('document.pdf', use_ocr=False)
else:
    print("这是扫描版PDF，将使用OCR识别")
    converter.convert_pdf_to_word('document.pdf', use_ocr=True)

# 示例4: 转换后搜索关键词
results = converter.search_keyword(word_file, '结论')
print(f"找到 {len(results)} 处结果")

# 示例5: 高亮显示"重要"
converter.highlight_keyword(word_file, '重要')
**Q: 如何识别扫描版PDF的文字？**  
A: 安装OCR功能后，程序会自动检测并使用OCR识别。详见 [OCR_SETUP.md](OCR_SETUP.md)

**Q: 转换失败怎么办？**  
A: 检查PDF文件是否损坏，确保有读取权限。扫描版PDF需要安装OCR功能。

**Q: OCR识别不准确怎么办？**  
A: 确保PDF图片清晰（300+ DPI），文字清晰无倾斜，已安装对应语言包。

**Q: 不想安装OCR可以用吗？**  
A: 可以！程序仍可处理普通文本型PDF，只是无法识别扫描版PDF。

**Q: 中文乱码怎么办？**  
A: 确保安装了中文语言包（chi_sim），并且PDF包含正确的中文字体。

**Q: 支持哪些语言？**  
A: 默认支持中文简体和英文，可通过安装Tesseract语言包扩展。

**Q: 转换速度慢？**  
A: OCR识别需要时间，特别是多页扫描文档。文本型PDF转换很快。重要"
converter.highlight_keyword(word_file, '重要')

# 替换公司名称
converter.replace_text(word_file, '旧公司名', '新公司名')
```

## 常见问题

**Q: 转换失败怎么办？**  
A: 检查 PDF 文件是否损坏，确保有读取权限

**Q: 中文乱码怎么办？**  
A: 确保 PDF 包含正确的中文字体嵌入

**Q: 格式不完全一致？**  
A: PDF 到 Word 的转换可能无法 100% 保留复杂格式

## 系统要求

- Python 3.7 或更高版本
- Windows / macOS / Linux

## 许可证

MIT License
