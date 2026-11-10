@echo off
setlocal DisableDelayedExpansion
chcp 65001 >nul
call "%~dp0install_deps.bat"
exit /b %ERRORLEVEL%
