@echo off
chcp 65001 >nul
python word_exam_formatter.py
if errorlevel 1 pause
