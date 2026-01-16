"""
OCR功能安装辅助工具
自动下载和配置Tesseract OCR和Poppler
"""
import os
import sys
import urllib.request
import zipfile
import subprocess
from pathlib import Path

def download_file(url, filename):
    """下载文件"""
    print(f"正在下载 {filename}...")
    try:
        urllib.request.urlretrieve(url, filename, reporthook=lambda b, bs, s: print(f'\r进度: {b*bs/s*100:.1f}%', end=''))
        print(f"\n✓ 下载完成: {filename}")
        return True
    except Exception as e:
        print(f"\n✗ 下载失败: {e}")
        return False

def main():
    print("=" * 70)
    print("OCR功能安装向导")
    print("=" * 70)
    print()
    
    print("OCR功能需要以下外部工具：")
    print("1. Tesseract OCR引擎 - 文字识别")
    print("2. Poppler - PDF转图片")
    print()
    
    # 检查是否有管理员权限
    import ctypes
    is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
    
    if not is_admin:
        print("⚠️  警告: 建议以管理员身份运行此脚本以便自动安装")
        print()
    
    print("安装选项：")
    print()
    print("【推荐】手动安装（更可靠）：")
    print()
    print("1️⃣  安装 Tesseract OCR:")
    print("   下载地址: https://github.com/UB-Mannheim/tesseract/wiki")
    print("   下载文件: tesseract-ocr-w64-setup-5.x.x.exe")
    print("   安装路径: C:\\Program Files\\Tesseract-OCR")
    print("   重要: 安装时勾选【中文语言包 chi_sim】")
    print()
    print("2️⃣  安装 Poppler:")
    print("   下载地址: https://github.com/oschwartz10612/poppler-windows/releases")
    print("   下载文件: Release-XX.XX.X-0.zip")
    print("   解压到: C:\\Program Files\\poppler")
    print()
    print("3️⃣  添加到系统PATH:")
    print("   - 右键'此电脑' → 属性 → 高级系统设置")
    print("   - 环境变量 → 系统变量 → Path → 编辑")
    print("   - 添加两个路径:")
    print("     * C:\\Program Files\\Tesseract-OCR")
    print("     * C:\\Program Files\\poppler\\Library\\bin")
    print()
    print("-" * 70)
    print()
    
    # 提供直接下载链接
    print("快速下载链接：")
    print()
    print("Tesseract (64位):")
    print("https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.3.20231005.exe")
    print()
    print("Poppler (最新版):")
    print("https://github.com/oschwartz10612/poppler-windows/releases/latest")
    print()
    print("=" * 70)
    
    # 检查是否已安装
    print()
    print("检查安装状态...")
    print()
    
    # 检查Tesseract
    tesseract_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    
    tesseract_found = False
    for path in tesseract_paths:
        if os.path.exists(path):
            print(f"✓ Tesseract 已安装: {path}")
            tesseract_found = True
            
            # 测试Tesseract
            try:
                result = subprocess.run([path, '--version'], capture_output=True, text=True)
                print(f"  版本: {result.stdout.split()[1]}")
            except:
                pass
            break
    
    if not tesseract_found:
        print("✗ Tesseract 未找到")
        print("  请按照上述说明安装Tesseract")
    
    # 检查Poppler
    poppler_paths = [
        r"C:\Program Files\poppler\Library\bin\pdfinfo.exe",
        r"C:\poppler\Library\bin\pdfinfo.exe",
    ]
    
    poppler_found = False
    for path in poppler_paths:
        if os.path.exists(path):
            print(f"✓ Poppler 已安装: {os.path.dirname(path)}")
            poppler_found = True
            break
    
    if not poppler_found:
        print("✗ Poppler 未找到")
        print("  请按照上述说明安装Poppler")
    
    print()
    print("=" * 70)
    
    if tesseract_found and poppler_found:
        print()
        print("🎉 恭喜！OCR功能已准备就绪！")
        print()
        print("您现在可以：")
        print("1. 运行 GUI: py pdf_to_word_gui.py")
        print("2. 转换扫描版PDF并自动识别文字")
        print()
    else:
        print()
        print("⚠️  OCR功能尚未完全安装")
        print()
        print("完成安装后，程序将自动支持扫描版PDF的文字识别")
        print("即使不安装OCR，程序仍可处理普通文本型PDF")
        print()
    
    input("按任意键退出...")

if __name__ == "__main__":
    main()
