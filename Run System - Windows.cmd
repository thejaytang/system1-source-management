@echo off
set "ROOT_DIR=%~dp0"
call "%ROOT_DIR%Code\scripts\run_windows.cmd"
exit /b %ERRORLEVEL%
