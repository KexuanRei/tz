@echo off
setlocal DisableDelayedExpansion
cd /d "%~dp0"
if errorlevel 1 exit /b 1
if not exist ".venv\Scripts\python.exe" goto find_python
".venv\Scripts\python.exe" -c "import sys,venv; assert sys.version_info >= (3,10)" >nul 2>&1
if not errorlevel 1 exit /b 0

:find_python
py -3 -c "import sys,venv; assert sys.version_info >= (3,10)" >nul 2>&1
if not errorlevel 1 goto use_py
python -c "import sys,venv; assert sys.version_info >= (3,10)" >nul 2>&1
if not errorlevel 1 goto use_python
echo Python 3.10 or newer was not found. Install Python with the launcher or add it to PATH.
exit /b 1

:use_py
set "SYNC_PYTHON=py -3"
goto rebuild
:use_python
set "SYNC_PYTHON=python"

:rebuild
if not exist ".venv" goto create
:backup_name
set "SYNC_BACKUP=.venv_old_%RANDOM%_%RANDOM%"
if exist "%SYNC_BACKUP%" goto backup_name
echo Preserving the old environment as %SYNC_BACKUP%
ren ".venv" "%SYNC_BACKUP%"
if errorlevel 1 goto backup_failed
:create
echo Creating a virtual environment with local Python...
%SYNC_PYTHON% -m venv .venv
if errorlevel 1 exit /b 1
".venv\Scripts\python.exe" -c "import sys; assert sys.prefix != sys.base_prefix" >nul 2>&1
if errorlevel 1 exit /b 1
exit /b 0

:backup_failed
echo Cannot rename the old environment. Close this project's running server and retry.
exit /b 1
