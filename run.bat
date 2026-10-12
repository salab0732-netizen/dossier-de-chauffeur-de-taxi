@echo off
cd /d "%~dp0"
echo Updating libraries (first run takes a minute)...
pip install -U google-genai anthropic pymupdf openpyxl -q
echo Starting...
python taxi_pro.py
if errorlevel 1 pause
