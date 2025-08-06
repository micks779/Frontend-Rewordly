@echo off
REM Icon Generation Script for Windows
REM Generates PNG icons from SVG source for the Rewordly Outlook Add-in

echo 🤖 Rewordly Icon Generator
echo ==========================

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python is not installed or not in PATH
    echo Please install Python and try again
    pause
    exit /b 1
)

REM Run the Python script
echo 🎨 Generating PNG icons...
python generate-icons.py

if errorlevel 1 (
    echo ❌ Icon generation failed!
    pause
    exit /b 1
)

echo.
echo 🎉 Icon generation complete!
echo.
echo Next steps:
echo 1. Review the generated PNG files in this directory
echo 2. The manifest.xml has been updated automatically
echo 3. For production, host these files on a CDN
echo.
pause 