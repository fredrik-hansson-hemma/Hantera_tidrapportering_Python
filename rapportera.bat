@echo off
chcp 65001 >nul
title Tidrapportering

.venv\Scripts\python.exe rapportera.py

echo.
pause
