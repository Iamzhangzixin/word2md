#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word2MD 打包脚本
自动安装依赖并打包成exe文件
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def run_command(cmd, description):
    """运行命令并显示进度"""
    print(f"\n🔄 {description}...")
    try:
        result = subprocess.run(cmd, shell=True, check=True, capture_output=True, text=True)
        print(f"✅ {description}完成")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {description}失败: {e}")
        if e.stdout:
            print(f"输出: {e.stdout}")
        if e.stderr:
            print(f"错误: {e.stderr}")
        return False

def check_and_install_pyinstaller():
    """检查并安装PyInstaller"""
    try:
        import PyInstaller
        print("✅ PyInstaller已安装")
        return True
    except ImportError:
        print("📦 安装PyInstaller...")
        return run_command("pip install pyinstaller", "安装PyInstaller")

def create_version_file():
    """创建版本信息文件"""
    version_content = """# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1, 0, 0, 0),
    prodvers=(1, 0, 0, 0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        u'080404B0',
        [StringStruct(u'CompanyName', u'Word2MD'),
        StringStruct(u'FileDescription', u'Word转Markdown转换器'),
        StringStruct(u'FileVersion', u'1.0.0.0'),
        StringStruct(u'InternalName', u'Word2MD'),
        StringStruct(u'LegalCopyright', u'Copyright © 2024'),
        StringStruct(u'OriginalFilename', u'Word2MD.exe'),
        StringStruct(u'ProductName', u'Word2MD Converter'),
        StringStruct(u'ProductVersion', u'1.0.0.0')])
      ]), 
    VarFileInfo([VarStruct(u'Translation', [2052, 1200])])
  ]
)"""
    
    with open('version_info.txt', 'w', encoding='utf-8') as f:
        f.write(version_content)
    print("✅ 版本信息文件已创建")

def build_exe():
    """使用PyInstaller打包exe"""
    print("\n🚀 开始打包Word2MD...")
    
    # 确保图标文件存在
    if not os.path.exists('word2md_icon.ico'):
        print("⚠️  图标文件不存在，创建默认图标...")
        run_command("python create_simple_icon.py", "创建图标")
    
    # 创建版本信息文件
    create_version_file()
    
    # PyInstaller命令
    cmd = [
        "pyinstaller",
        "--onefile",  # 打包成单个exe文件
        "--windowed",  # 无控制台窗口
        "--name=Word2MD",  # exe文件名
        "--icon=word2md_icon.ico",  # 图标
        "--version-file=version_info.txt",  # 版本信息
        "--add-data=word2md_icon.ico;.",  # 包含图标文件
        "--add-data=requirements.txt;.",  # 包含依赖文件
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.ttk",
        "--hidden-import=tkinter.filedialog",
        "--hidden-import=tkinter.messagebox",
        "--hidden-import=tkinter.scrolledtext",
        "--hidden-import=docx",
        "--hidden-import=mammoth",
        "--hidden-import=pypandoc",
        "--hidden-import=PIL",
        "--hidden-import=PIL.Image",
        "--hidden-import=PIL.ImageDraw",
        "--hidden-import=lxml",
        "--hidden-import=lxml.etree",
        "--exclude-module=matplotlib",
        "--exclude-module=numpy",
        "--exclude-module=scipy",
        "--exclude-module=pandas",
        "--exclude-module=jupyter",
        "--exclude-module=IPython",
        "--exclude-module=pytest",
        "--exclude-module=setuptools",
        "word2md_enhanced.py"
    ]
    
    cmd_str = " ".join(cmd)
    return run_command(cmd_str, "打包exe文件")

def create_portable_package():
    """创建便携版包"""
    print("\n📦 创建便携版包...")
    
    # 创建发布目录
    release_dir = Path("release")
    if release_dir.exists():
        shutil.rmtree(release_dir)
    release_dir.mkdir()
    
    # 复制exe文件
    exe_path = Path("dist/Word2MD.exe")
    if exe_path.exists():
        shutil.copy2(exe_path, release_dir / "Word2MD.exe")
        print("✅ exe文件已复制到release目录")
    else:
        print("❌ 找不到exe文件")
        return False
    
    # 创建使用说明
    readme_content = """# Word2MD - Word转Markdown转换器

## 功能特点
- 支持单文件和批量转换
- 自动处理图片、公式、表格
- 支持.docx、.doc格式
- 完全独立运行，无需安装Python

## 使用方法
1. 双击Word2MD.exe启动程序
2. 选择转换模式（单文件或批量）
3. 选择Word文档和输出位置
4. 点击"开始转换"

## 注意事项
- 首次运行可能需要几秒钟启动时间
- 转换大文件时请耐心等待
- 图片会自动保存到images文件夹

## 技术支持
如遇问题，请检查：
1. Word文档是否完整
2. 输出目录是否有写入权限
3. 磁盘空间是否充足

版本: 1.0.0
"""
    
    with open(release_dir / "使用说明.txt", 'w', encoding='utf-8') as f:
        f.write(readme_content)
    
    # 复制示例文件
    if Path("test").exists():
        example_dir = release_dir / "示例文件"
        example_dir.mkdir()
        for file in Path("test").glob("*.docx"):
            if file.name != "X射线脉冲星光子到达时间建模.docx":
                continue
            shutil.copy2(file, example_dir / file.name)
            print(f"✅ 示例文件已复制: {file.name}")
    
    print(f"✅ 便携版包已创建在: {release_dir.absolute()}")
    return True

def main():
    """主函数"""
    print("🎯 Word2MD 自动打包工具")
    print("=" * 50)
    
    # 检查当前目录
    if not os.path.exists('word2md_enhanced.py'):
        print("❌ 找不到word2md_enhanced.py文件")
        print("请在包含源代码的目录中运行此脚本")
        return False
    
    # 检查并安装PyInstaller
    if not check_and_install_pyinstaller():
        return False
    
    # 清理之前的构建
    for dir_name in ['build', 'dist', '__pycache__']:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"🧹 已清理: {dir_name}")
    
    # 打包exe
    if not build_exe():
        return False
    
    # 创建便携版包
    if not create_portable_package():
        return False
    
    print("\n🎉 打包完成！")
    print("📁 exe文件位置: dist/Word2MD.exe")
    print("📦 便携版包位置: release/")
    print("\n✨ 现在您可以直接运行Word2MD.exe，无需Python环境！")
    
    return True

if __name__ == "__main__":
    try:
        success = main()
        if not success:
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n⚠️  用户取消操作")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 打包过程中出现错误: {e}")
        sys.exit(1)
