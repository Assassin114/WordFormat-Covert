@echo off
REM ==========================================
REM DocFormatter 打包脚本 (Windows)
REM 执行: build.bat
REM 输出: dist/DocFormatter/ 文件夹
REM 发布: 将 dist/DocFormatter/ 打包成 zip
REM ==========================================

echo.
echo === DocFormatter 打包工具 ===
echo.

REM 检查 PyInstaller
python -m PyInstaller --version >/dev/null 2>&1
if %errorlevel% neq 0 (
    echo [安装 PyInstaller...]
    pip install pyinstaller
)

echo [开始打包...]
python -m PyInstaller ^
    --noconfirm ^
    --name=DocFormatter ^
    --windowed ^
    --add-data="docformatter/templates/builtin;docformatter/templates/builtin" ^
    docformatter/main.py

echo.
echo === 打包完成 ===
echo 输出目录: dist\DocFormatter\
echo 发布方式: 将 dist\DocFormatter\ 整个文件夹打包成 zip 发给用户
echo 用户使用: 解压后双击 DocFormatter.exe 即可运行
echo.

pause
