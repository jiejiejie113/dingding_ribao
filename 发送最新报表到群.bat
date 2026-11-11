@echo off
setlocal
cd /d "%~dp0"
set "PYTHONUTF8=1"
chcp 65001 >nul
python -X utf8 "%~dp0daily_send.py"
pause
