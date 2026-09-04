@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "CODE_DIR=%SCRIPT_DIR%.."
cd /d "%CODE_DIR%"
set "PYTHONPATH=%CODE_DIR%\src"
".venv\Scripts\python.exe" -m system1 unregister-schedule
pause

