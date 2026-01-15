# OCR 功能安装指南

本程序支持两种PDF转换模式：

## 📝 模式说明

### 1. 标准模式（默认）
- 适用于：文本型PDF（可复制文字的PDF）
- 特点：快速、准确
- 无需额外配置

### 2. OCR模式（文字识别）
- 适用于：扫描版PDF、图片PDF
- 特点：识别图片中的文字
- 支持：中文、英文等多语言

## 🔧 OCR功能安装步骤

### Windows 系统

**1. 安装Python依赖包**
```bash
pip install pdf2image pytesseract Pillow
```

**2. 安装Tesseract OCR引擎**

下载并安装Tesseract：
- 下载地址：https://github.com/UB-Mannheim/tesseract/wiki
- 下载文件：`tesseract-ocr-w64-setup-5.x.x.exe`
- 安装时选择安装路径，例如：`C:\Program Files\Tesseract-OCR`
- **重要**：安装时勾选"中文语言包"

**3. 配置环境变量**

方法一：添加到系统PATH
- 右键"此电脑" → "属性" → "高级系统设置"
- "环境变量" → 系统变量中的"Path"
- 添加：`C:\Program Files\Tesseract-OCR`

方法二：在程序中指定（推荐）
在Python代码开头添加：
```python
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

**4. 安装Poppler（PDF转图片工具）**

下载Poppler：
- 下载地址：https://github.com/oschwartz10612/poppler-windows/releases
- 下载最新版本的zip文件
- 解压到：`C:\Program Files\poppler`
- 添加到PATH：`C:\Program Files\poppler\Library\bin`

### macOS 系统

```bash
# 安装Tesseract
brew install tesseract tesseract-lang

# 安装Poppler
brew install poppler

# 安装Python依赖
pip install pdf2image pytesseract Pillow
```

### Linux 系统

```bash
# Ubuntu/Debian
sudo apt-get update
sudo apt-get install tesseract-ocr tesseract-ocr-chi-sim poppler-utils

# 安装Python依赖
pip install pdf2image pytesseract Pillow
```

## ✅ 验证安装

运行以下Python代码测试：

```python
# 测试Tesseract
import pytesseract
print(pytesseract.get_tesseract_version())

# 测试pdf2image
from pdf2image import convert_from_path
print("pdf2image 已安装")

print("✓ OCR功能准备就绪！")
```

## 🎯 使用方法

程序会**自动检测**PDF类型：
- 如果是文本型PDF → 使用标准模式
- 如果是扫描版PDF → 自动使用OCR模式

手动指定模式：
```python
from pdf_to_word_converter import PDFToWordConverter

converter = PDFToWordConverter()

# 强制使用OCR
converter.convert_pdf_to_word('scan.pdf', use_ocr=True)

# 不使用OCR
converter.convert_pdf_to_word('text.pdf', use_ocr=False)

# 自动检测（默认）
converter.convert_pdf_to_word('document.pdf')
```

## 📊 识别语言

**已支持的语言**：
- 中文简体：chi_sim
- 英文：eng

**添加更多语言**：

Windows - 下载语言包：
```
https://github.com/tesseract-ocr/tessdata
```
将 `.traineddata` 文件放到：
```
C:\Program Files\Tesseract-OCR\tessdata\
```

修改代码支持更多语言：
```python
text = pytesseract.image_to_string(image, lang='chi_sim+eng+jpn')
```

## ❓ 常见问题

**Q: 提示"tesseract not found"**
A: Tesseract未正确安装或未添加到PATH

**Q: OCR识别准确率不高**
A: 
- 确保PDF图片清晰度足够（建议300 DPI）
- 确保安装了对应语言包
- 扫描文档保持文字清晰、无倾斜

**Q: 转换速度慢**
A: OCR识别需要时间，特别是多页文档。标准模式更快。

**Q: 不想安装OCR功能**
A: 程序仍可正常工作，只是不支持扫描版PDF的文字识别

## 📚 更多信息

- Tesseract官方文档：https://tesseract-ocr.github.io/
- pdf2image文档：https://github.com/Belval/pdf2image
- pytesseract文档：https://github.com/madmaze/pytesseract
