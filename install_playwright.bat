@echo off
setlocal

REM Run from this script's directory
cd /d "%~dp0"

echo [1/2] Installing Playwright Python package...
pip install playwright
if errorlevel 1 (
    echo.
    echo [ERROR] pip install playwright failed.
    echo Please check Python/pip installation and your network connection.
    pause
    exit /b 1
)

echo.
echo [2/2] Installing Chromium for Playwright...
playwright install chromium
if errorlevel 1 (
    echo.
    echo [WARN] playwright command failed, trying python -m playwright...
    python -m playwright install chromium
    if errorlevel 1 (
        echo.
        echo [ERROR] Failed to install Chromium browser for Playwright.
        pause
        exit /b 1
    )
)

echo.
echo [OK] Playwright + Chromium installation completed.
pause
exit /b 0
