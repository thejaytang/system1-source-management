@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "CODE_DIR=%SCRIPT_DIR%.."
cd /d "%CODE_DIR%"
if not exist ".venv\Scripts\python.exe" (
  echo The project environment is missing. Run Code\deployment\setup_windows.cmd first.
  pause
  exit /b 2
)
set "PYTHONPATH=%CODE_DIR%\src"
".venv\Scripts\python.exe" -m system1 menu --config config\config.json --schedule config\schedule.json
set "RESULT=%ERRORLEVEL%"
pause
exit /b %RESULT%

