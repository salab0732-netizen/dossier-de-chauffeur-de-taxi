@echo off
chcp 65001 > nul
title استخراج PDF محلي - بدون إنترنت

echo ════════════════════════════════════════════
echo   نظام استخراج PDF المحلي (Offline OCR)
echo   ولاية البيض - سيارات الأجرة
echo ════════════════════════════════════════════
echo.

REM التحقق من وجود tessdata
if not exist "%~dp0tessdata\ara.traineddata" (
    echo ⬇  تثبيت اللغة العربية أولاً...
    python "%~dp0offline_ocr.py" --setup
    if errorlevel 1 (
        echo ❌ فشل التثبيت. تحقق من الاتصال بالإنترنت.
        pause
        exit /b 1
    )
)

echo.
echo 🚀 بدء معالجة ملفات PDF...
python "%~dp0offline_ocr.py" %*

echo.
echo ════════════════════════════════════════════
pause
