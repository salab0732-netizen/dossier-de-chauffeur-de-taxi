@echo off
chcp 65001 >nul
cd /d "%~dp0"
title نظام سيارات الأجرة - ولاية البيض

echo   نظام إدخال بيانات سيارات الأجرة — ولاية البيض (32)
echo   جارٍ التحقق من المكتبات...
pip install -q fastapi "uvicorn[standard]" openpyxl pymupdf pywebview google-genai anthropic 2>nul

echo   جارٍ التشغيل...
python run.py

if errorlevel 1 (
  echo.
  echo   تعذر التشغيل. تأكد من تثبيت Python 3.9 أو أحدث.
  pause
)
