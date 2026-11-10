@echo off
setlocal DisableDelayedExpansion
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
title Install dependencies
cd /d "%~dp0"
if errorlevel 1 goto failed
echo Installing project dependencies. Please keep this window open.
if not exist "requirements.txt" goto missing_requirements
call "%~dp0prepare_env.bat"
if errorlevel 1 goto failed
".venv\Scripts\python.exe" -m pip install -r requirements.txt --timeout 20 --retries 1
if errorlevel 1 goto retry_index
goto verify
:retry_index
echo Retrying with the official PyPI index...
".venv\Scripts\python.exe" -m pip install -r requirements.txt --index-url https://pypi.org/simple --timeout 20 --retries 1
if errorlevel 1 goto failed
:verify
".venv\Scripts\python.exe" -c "import lxml.etree; import openpyxl; import playwright.sync_api; import win32com.client; import win32clipboard; print('Dependency check passed.')"
if errorlevel 1 goto failed
echo Installation succeeded. Restart start_ui.bat to use the updated environment.
".venv\Scripts\python.exe" install_success.py
exit /b 0
:missing_requirements
echo requirements.txt was not found. Extract the complete project before installing.
:failed
echo Installation did not complete. See the error above.
pause
exit /b 1
