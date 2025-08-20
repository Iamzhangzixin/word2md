@echo off
echo 正在打包Word2MD应用...

REM 激活conda环境
call conda activate word2md

REM 安装PyInstaller
pip install pyinstaller

REM 清理之前的构建
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM 使用PyInstaller打包
pyinstaller --onefile --windowed --name=Word2MD --icon=word2md_icon.ico --add-data="word2md_icon.ico;." word2md_enhanced.py

echo 打包完成！exe文件位于 dist\Word2MD.exe
pause
