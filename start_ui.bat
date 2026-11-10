@echo off
setlocal DisableDelayedExpansion
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
cd /d "%~dp0"
if errorlevel 1 goto failed
if not exist ".venv\Scripts\python.exe" goto install
".venv\Scripts\python.exe" -c "import sys; assert sys.version_info >= (3,10)" >nul 2>&1
if errorlevel 1 goto install
goto run
:install
echo The Python environment is missing or invalid. Installing dependencies...
call "%~dp0install_deps.bat"
if errorlevel 1 goto failed
:run
".venv\Scripts\python.exe" web_app.py
set "SYNC_EXIT=%ERRORLEVEL%"
pause
exit /b %SYNC_EXIT%
:failed
echo Startup did not complete. See the error above.
pause
exit /b 1
