@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "CODE_DIR=%SCRIPT_DIR%.."
cd /d "%CODE_DIR%"
set "PYTHONPATH=%CODE_DIR%\src"
".venv\Scripts\python.exe" -m system1 doctor --config config\config.json --schedule config\schedule.json
if errorlevel 1 (
  echo Environment validation failed. The schedule was not registered.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m system1 register-schedule --config config\config.json --schedule config\schedule.json
pause

