@echo off
echo Installing dependencies...
pip install PyQt6 python-docx lxml pyinstaller

echo.
echo Building...
pyinstaller build.spec --clean

echo.
echo Done! Check dist\ folder for DocFormatter.exe
pause
