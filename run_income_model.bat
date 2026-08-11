@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ==========================================
echo Taiwan Stock Income Model Generator
echo Python version: 3.12
echo ==========================================
echo.
echo Enter stock codes separated by spaces.
echo Example: 2330
echo Example: 2330 3532 2454
echo.

set /p STOCKS=Stock codes:

if "%STOCKS%"=="" (
    echo.
    echo No stock code entered.
    pause
    exit /b 1
)

echo.
echo Generating Excel model for: %STOCKS%
echo.

py -3.12 scripts\build_income_statement_model.py --stocks %STOCKS% --start_year 2024

if errorlevel 1 (
    echo.
    echo Failed to generate model using Python 3.12.
    echo Please check that Python 3.12 is installed:
    echo py -0p
    pause
    exit /b 1
)

echo.
echo Done. Opening output folder...
explorer output\models

echo.
pause
