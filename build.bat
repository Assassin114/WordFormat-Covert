@echo off
REM DocFormatter Windows Build Script

echo Building DocFormatter...
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python not found
    exit /b 1
)

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt -q

REM Build with PyInstaller
echo Building executable...
pyinstaller build.spec --clean

if exist "dist\DocFormatter.exe" (
    echo.
    echo Build successful!
    echo Output: dist\DocFormatter.exe
    echo.
    set /p OPEN=Open output folder? (Y/N):
    if /i "%OPEN%"=="Y" explorer dist
) else (
    echo.
    echo Build failed!
    exit /b 1
)
