@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "CODE_DIR=%SCRIPT_DIR%.."
cd /d "%CODE_DIR%"
where py >nul 2>nul
if %errorlevel%==0 (
  set "PYTHON_CMD=py -3"
) else (
  where python >nul 2>nul
  if errorlevel 1 (
    echo Python 3.11 or newer is required. Install Python and run this setup again.
    pause
    exit /b 2
  )
  set "PYTHON_CMD=python"
)
%PYTHON_CMD% -c "import sys; raise SystemExit(0 if sys.version_info >= (3,11) else 1)"
if errorlevel 1 (
  echo Python 3.11 or newer is required.
  pause
  exit /b 2
)
%PYTHON_CMD% -m venv --prompt SmarterComplianceSystem1 .venv
if errorlevel 1 exit /b 1
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 exit /b 1
".venv\Scripts\python.exe" -m pip install -r deployment\requirements.txt
if errorlevel 1 exit /b 1
if not exist config\config.json copy config\config.example.json config\config.json >nul
if not exist config\schedule.json copy config\schedule.example.json config\schedule.json >nul
set "PYTHONPATH=%CODE_DIR%\src"
".venv\Scripts\python.exe" -m system1 doctor --config config\config.json --schedule config\schedule.json
echo.
echo Environment created. Review Code\config\config.json and schedule.json, then run the doctor again.
pause
