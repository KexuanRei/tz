@echo off
setlocal DisableDelayedExpansion
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
cd /d "%~dp0"
if errorlevel 1 exit /b 1
if "%~1"=="" goto usage
if not exist ".venv\Scripts\python.exe" goto noenv
".venv\Scripts\python.exe" make_table_c.py %1 %2 %3
set "SYNC_EXIT=%ERRORLEVEL%"
pause
exit /b %SYNC_EXIT%
:usage
echo Usage: ???C.bat ?A??.xlsx ?C.xlsx [????]
echo The first workbook is the local copy of table A; the second is the format-only table C.
pause
exit /b 1
:noenv
echo Python environment missing. Run ????.bat first.
pause
exit /b 1
