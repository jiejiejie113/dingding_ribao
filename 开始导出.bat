@echo off
setlocal
cd /d "%~dp0"
set "PYTHONUTF8=1"
chcp 65001 >nul
python --version >nul 2>&1
if errorlevel 1 goto no_python
python -X utf8 "%~dp0launcher.py" export %*
set "RESULT=%ERRORLEVEL%"
pause
exit /b %RESULT%
:no_python
echo Python was not found. Install Python 3.10+ and enable Add Python to PATH.
pause
exit /b 1
