@echo off
setlocal DisableDelayedExpansion
chcp 65001 >nul
call "%~dp0raise_read_limit.bat"
exit /b %ERRORLEVEL%
