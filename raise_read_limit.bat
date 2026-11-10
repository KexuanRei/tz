@echo off
setlocal DisableDelayedExpansion
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
cd /d "%~dp0"
if errorlevel 1 goto failed
echo Raising the saved cell limit to at least 100,000,000.
echo Close this project's server before continuing; restart it after this finishes.
call "%~dp0prepare_env.bat"
if errorlevel 1 goto failed
".venv\Scripts\python.exe" apply_high_limit.py
if errorlevel 1 goto failed
pause
exit /b 0
:failed
echo The operation did not complete. See the error above.
pause
exit /b 1
