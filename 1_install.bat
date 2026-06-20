@echo off
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

pip install playwright
pip install pyside6
pip install pyinstaller
playwright install chromium
