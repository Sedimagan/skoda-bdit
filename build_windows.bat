@echo off
title Skoda Best Dealer in Town App - Builder
echo.
echo =====================================================
echo   Skoda Best Dealer in Town App  -  Windows Builder
echo =====================================================
echo.

:: Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

echo [1/3] Installing required packages...
pip install --upgrade Pillow openpyxl pyinstaller
if errorlevel 1 (
    echo ERROR: Failed to install packages. Check your internet connection.
    pause
    exit /b 1
)

echo.
echo [2/3] Building Windows executable (this takes 1-2 minutes)...
pyinstaller ^
  --name "Skoda Best Dealer in Town App" ^
  --windowed ^
  --onefile ^
  --icon skoda.ico ^
  --hidden-import PIL ^
  --hidden-import PIL.Image ^
  --hidden-import PIL.ImageTk ^
  --hidden-import openpyxl ^
  --hidden-import openpyxl.styles ^
  --hidden-import openpyxl.utils ^
  sales_score_calculator.py

if errorlevel 1 (
    echo ERROR: Build failed. See messages above.
    pause
    exit /b 1
)

echo.
echo [3/3] Done!
echo.
echo -------------------------------------------------------
echo  Your executable is ready:
echo  dist\Skoda Best Dealer in Town App.exe
echo.
echo  Copy that .exe file to any Windows computer.
echo  No Python installation needed to run it.
echo.
echo  Data and Excel file are saved to:
echo  C:\Users\YOUR_NAME\.skoda_bdit\
echo -------------------------------------------------------
echo.
pause
