@echo off
chcp 65001 >nul
echo ========================================
echo DocFormatter 构建脚本
echo ========================================
echo.

echo [1/3] 检查 Python 版本...
python --version
echo.

echo [2/3] 安装依赖（请等待）...
pip install PyQt6 python-docx lxml
if errorlevel 1 (
    echo PyQt6 等安装失败，请以管理员身份运行
    pause
    exit /b 1
)

pip install pyinstaller
if errorlevel 1 (
    echo pyinstaller 安装失败
    pause
    exit /b 1
)

echo.
echo [3/3] 打包（请等待，可能需要几分钟）...
pyinstaller build.spec --clean

if errorlevel 1 (
    echo 打包失败
    pause
    exit /b 1
)

echo.
echo ========================================
echo 构建完成！
echo 文件位置: dist\DocFormatter.exe
echo ========================================
pause
