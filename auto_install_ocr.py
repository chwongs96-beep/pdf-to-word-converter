"""
自动安装OCR工具（Tesseract和Poppler）
"""
import os
import sys
import urllib.request
import zipfile
import subprocess
from pathlib import Path
import shutil

def download_with_progress(url, filename):
    """带进度条下载文件"""
    print(f"\n正在下载: {filename}")
    print(f"URL: {url}")
    
    def progress(block_num, block_size, total_size):
        downloaded = block_num * block_size
        percent = min(downloaded * 100 / total_size, 100)
        bar_length = 50
        filled = int(bar_length * percent / 100)
        bar = '█' * filled + '░' * (bar_length - filled)
        print(f'\r[{bar}] {percent:.1f}% ({downloaded/1024/1024:.1f}MB/{total_size/1024/1024:.1f}MB)', end='')
    
    try:
        urllib.request.urlretrieve(url, filename, reporthook=progress)
        print(f"\n✓ 下载完成: {filename}\n")
        return True
    except Exception as e:
        print(f"\n✗ 下载失败: {e}\n")
        return False

def install_tesseract():
    """下载并引导安装Tesseract"""
    print("=" * 70)
    print("步骤 1/2: 安装 Tesseract OCR引擎")
    print("=" * 70)
    
    # Tesseract 下载URL
    tesseract_url = "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.3.20231005.exe"
    tesseract_file = "tesseract-installer.exe"
    
    # 检查是否已安装
    tesseract_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.exists(tesseract_path):
        print(f"✓ Tesseract 已安装: {tesseract_path}")
        return True
    
    # 下载
    if not os.path.exists(tesseract_file):
        print("正在下载Tesseract安装程序...")
        if not download_with_progress(tesseract_url, tesseract_file):
            print("下载失败。请手动下载:")
            print(tesseract_url)
            return False
    
    # 运行安装程序
    print("\n" + "=" * 70)
    print("启动Tesseract安装程序...")
    print("=" * 70)
    print()
    print("⚠️  重要提示:")
    print("1. 在安装过程中，请勾选【Additional language data】")
    print("2. 展开后勾选【Chinese - Simplified】(chi_sim)")
    print("3. 建议安装路径: C:\\Program Files\\Tesseract-OCR")
    print()
    input("准备好后按回车键开始安装...")
    
    try:
        # 启动安装程序
        subprocess.Popen([tesseract_file])
        print("\n安装程序已启动，请按照提示完成安装")
        print("安装完成后请关闭安装窗口，然后按回车继续...")
        input()
        
        # 检查是否安装成功
        if os.path.exists(tesseract_path):
            print("✓ Tesseract 安装成功！")
            return True
        else:
            print("⚠️  请确保已完成Tesseract安装")
            return False
    except Exception as e:
        print(f"启动安装程序失败: {e}")
        return False

def install_poppler():
    """下载并安装Poppler"""
    print("\n" + "=" * 70)
    print("步骤 2/2: 安装 Poppler PDF工具")
    print("=" * 70)
    
    # Poppler 下载URL (使用备用下载源)
    poppler_url = "https://github.com/oschwartz10612/poppler-windows/releases/download/v24.08.0-0/Release-24.08.0-0.zip"
    poppler_file = "poppler.zip"
    poppler_dir = r"C:\Program Files\poppler"
    
    # 检查是否已安装
    if os.path.exists(os.path.join(poppler_dir, "Library", "bin", "pdfinfo.exe")):
        print(f"✓ Poppler 已安装: {poppler_dir}")
        return True
    
    # 下载
    if not os.path.exists(poppler_file):
        print("\n正在下载Poppler...")
        if not download_with_progress(poppler_url, poppler_file):
            print("下载失败。请手动下载:")
            print("https://github.com/oschwartz10612/poppler-windows/releases/latest")
            return False
    
    # 解压
    print(f"\n正在解压到: {poppler_dir}")
    try:
        # 创建目录
        os.makedirs(poppler_dir, exist_ok=True)
        
        # 解压
        with zipfile.ZipFile(poppler_file, 'r') as zip_ref:
            # 解压所有文件
            for member in zip_ref.namelist():
                # 移除顶层目录
                target_path = os.path.join(poppler_dir, *member.split('/')[1:])
                if member.endswith('/'):
                    os.makedirs(target_path, exist_ok=True)
                else:
                    os.makedirs(os.path.dirname(target_path), exist_ok=True)
                    with zip_ref.open(member) as source, open(target_path, 'wb') as target:
                        shutil.copyfileobj(source, target)
        
        print("✓ Poppler 解压完成！")
        return True
    except Exception as e:
        print(f"✗ 解压失败: {e}")
        return False

def configure_path():
    """配置环境变量"""
    print("\n" + "=" * 70)
    print("配置环境变量")
    print("=" * 70)
    
    tesseract_path = r"C:\Program Files\Tesseract-OCR"
    poppler_path = r"C:\Program Files\poppler\Library\bin"
    
    print("\n需要将以下路径添加到系统PATH环境变量:")
    print(f"1. {tesseract_path}")
    print(f"2. {poppler_path}")
    print()
    print("自动配置方法:")
    print()
    
    # 生成PowerShell命令
    ps_command = f"""
$paths = @(
    "{tesseract_path}",
    "{poppler_path}"
)

$currentPath = [Environment]::GetEnvironmentVariable("Path", "User")
foreach ($path in $paths) {{
    if ($currentPath -notlike "*$path*") {{
        [Environment]::SetEnvironmentVariable(
            "Path",
            $currentPath + ";" + $path,
            "User"
        )
        Write-Host "✓ 已添加: $path"
    }} else {{
        Write-Host "○ 已存在: $path"
    }}
}}
Write-Host "`n环境变量配置完成！"
Write-Host "请重启终端或程序以使更改生效"
"""
    
    # 保存PowerShell脚本
    ps_file = "configure_path.ps1"
    with open(ps_file, 'w', encoding='utf-8') as f:
        f.write(ps_command)
    
    print("选项1: 自动配置（推荐）")
    choice = input("是否自动配置环境变量? (y/n): ").lower()
    
    if choice == 'y':
        try:
            result = subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_file],
                capture_output=True,
                text=True
            )
            print(result.stdout)
            if result.returncode == 0:
                print("\n✓ 环境变量配置成功！")
                return True
        except Exception as e:
            print(f"自动配置失败: {e}")
    
    print("\n选项2: 手动配置")
    print("步骤:")
    print("1. 按 Win + R，输入: sysdm.cpl")
    print("2. 高级 → 环境变量")
    print("3. 用户变量 → Path → 编辑")
    print("4. 新建，添加以下两个路径:")
    print(f"   - {tesseract_path}")
    print(f"   - {poppler_path}")
    print("5. 确定保存")
    return False

def verify_installation():
    """验证安装"""
    print("\n" + "=" * 70)
    print("验证安装")
    print("=" * 70)
    print()
    
    success = True
    
    # 检查Tesseract
    tesseract_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.exists(tesseract_path):
        print(f"✓ Tesseract: {tesseract_path}")
        try:
            result = subprocess.run([tesseract_path, '--version'], 
                                  capture_output=True, text=True)
            version = result.stdout.split('\n')[0]
            print(f"  {version}")
            
            # 检查中文语言包
            tessdata_dir = r"C:\Program Files\Tesseract-OCR\tessdata"
            chi_sim = os.path.join(tessdata_dir, "chi_sim.traineddata")
            if os.path.exists(chi_sim):
                print("  ✓ 中文语言包已安装")
            else:
                print("  ⚠️  中文语言包未找到")
                print(f"     请从以下地址下载 chi_sim.traineddata:")
                print(f"     https://github.com/tesseract-ocr/tessdata/raw/main/chi_sim.traineddata")
                print(f"     并放到: {tessdata_dir}")
                success = False
        except:
            print("  ⚠️  无法验证版本")
    else:
        print("✗ Tesseract: 未安装")
        success = False
    
    # 检查Poppler
    poppler_bin = r"C:\Program Files\poppler\Library\bin\pdfinfo.exe"
    if os.path.exists(poppler_bin):
        print(f"✓ Poppler: {os.path.dirname(poppler_bin)}")
    else:
        print("✗ Poppler: 未安装")
        success = False
    
    # 测试Python导入
    print("\n测试Python包:")
    try:
        from pdf2image import convert_from_path
        print("✓ pdf2image 导入成功")
    except ImportError as e:
        print(f"✗ pdf2image 导入失败: {e}")
        success = False
    
    try:
        import pytesseract
        print("✓ pytesseract 导入成功")
        
        # 配置tesseract路径
        if os.path.exists(tesseract_path):
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
            print(f"  已配置路径: {tesseract_path}")
    except ImportError as e:
        print(f"✗ pytesseract 导入失败: {e}")
        success = False
    
    return success

def main():
    print("=" * 70)
    print("OCR功能自动安装程序")
    print("=" * 70)
    print()
    print("此程序将自动安装:")
    print("1. Tesseract OCR引擎 (文字识别)")
    print("2. Poppler PDF工具 (PDF转图片)")
    print()
    
    input("按回车键开始安装...")
    
    # 安装Tesseract
    tesseract_ok = install_tesseract()
    
    # 安装Poppler
    poppler_ok = install_poppler()
    
    # 配置环境变量
    if tesseract_ok or poppler_ok:
        path_ok = configure_path()
    
    # 验证安装
    print()
    if verify_installation():
        print("\n" + "=" * 70)
        print("🎉 安装完成！")
        print("=" * 70)
        print()
        print("OCR功能已准备就绪！")
        print()
        print("现在您可以:")
        print("1. 重启终端或IDE")
        print("2. 运行: py pdf_to_word_gui.py")
        print("3. 转换扫描版PDF，程序将自动识别文字")
        print()
    else:
        print("\n" + "=" * 70)
        print("⚠️  安装未完全完成")
        print("=" * 70)
        print()
        print("请查看上述错误信息并手动完成安装")
        print("详细说明请参考: OCR_SETUP.md")
        print()
    
    input("\n按回车键退出...")

if __name__ == "__main__":
    main()
