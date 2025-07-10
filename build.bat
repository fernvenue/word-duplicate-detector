@echo off
chcp 65001 > nul
echo Word Duplicate Content Detection Tool - Build Script
echo =====================================================

echo Checking Python environment...
python --version > nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed. Please install Python 3.7+
    pause
    exit /b 1
)
echo Python environment check passed

echo.
echo Installing/updating dependencies...
pip install --upgrade pip
pip install pyinstaller lxml python-docx
if errorlevel 1 (
    echo Failed to install dependencies
    pause
    exit /b 1
)
echo Dependencies installed successfully

echo.
echo Building executable...
pyinstaller --onefile --windowed --name="WordDuplicateDetector" --distpath=dist --workpath=build --specpath=. --add-data "locales;locales" main.py
if errorlevel 1 (
    echo Build failed, please check error messages
    pause
    exit /b 1
)

echo.
if exist "dist\WordDuplicateDetector.exe" (
    echo Build successful!
    echo Executable location: %cd%\dist\
    echo Executable file: WordDuplicateDetector.exe
    echo.
    echo Build completed successfully!
    pause
) else (
    echo Build failed - executable not found
    pause
    exit /b 1
)
