# PDF 到 Word 转换器

这是一个功能强大的 Python 程序，可以将 PDF 文件转换为 Word 文档，并提供编辑和关键词搜索功能。

## 功能特性

✅ **PDF 转 Word** - 将 PDF 文件转换为可编辑的 Word (.docx) 文档  
✅ **关键词搜索** - 在 Word 文档中搜索特定关键词  
✅ **关键词高亮** - 自动高亮显示文档中的关键词  
✅ **文本替换** - 批量替换文档中的文本内容  
✅ **添加内容** - 在文档末尾添加新的文本  
✅ **文档信息** - 查看文档的段落数、字符数等统计信息

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

```bash
python pdf_to_word_converter.py
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

# 5. 添加文本
converter.add_text_to_document('output.docx', '这是新添加的内容')

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
- 查看总字符数

## 注意事项

⚠️ **PDF 格式限制**
- 建议使用文本型 PDF（非扫描版）
- 扫描版 PDF 需要先进行 OCR 处理

⚠️ **文件路径**
- 支持相对路径和绝对路径
- Windows 系统建议使用绝对路径

⚠️ **编码问题**
- 程序使用 UTF-8 编码
- 确保 PDF 和 Word 文件使用正确的编码

## 示例

```python
# 完整示例
converter = PDFToWordConverter()

# 转换 PDF
word_file = converter.convert_pdf_to_word('报告.pdf')

# 搜索"结论"关键词
results = converter.search_keyword(word_file, '结论')
print(f"找到 {len(results)} 处结果")

# 高亮显示"重要"
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
