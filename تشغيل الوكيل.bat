@echo off
cd /d "%~dp0"
pip install anthropic pymupdf openpyxl -q
python taxi_agent_gui.py
pause
