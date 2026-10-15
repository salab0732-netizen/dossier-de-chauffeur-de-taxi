@echo off
chcp 65001 >nul
cd /d "%~dp0"
pip install -q fastapi "uvicorn[standard]" openpyxl pymupdf pytesseract 2>nul
python run.py --browser
if errorlevel 1 pause
