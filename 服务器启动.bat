@echo off
setlocal DisableDelayedExpansion
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
cd /d "%~dp0"
if errorlevel 1 goto failed
if not exist ".venv\Scripts\python.exe" goto noenv

rem 无人值守：启动网页服务后自动开始定时任务。
rem 需要从家里/外网打开工作台时，把下面一行开头的 rem 去掉，并改成你的服务器地址与密码：
rem set "SYNC_EXTRA=--host 0.0.0.0 --allow-host 你的域名或公网IP --auth admin:改成你的密码"
rem 只用 RDP 或 SSH 隧道访问时可保持注释，默认只监听本机 127.0.0.1。

".venv\Scripts\python.exe" web_app.py --no-browser --auto-start %SYNC_EXTRA%
set "SYNC_EXIT=%ERRORLEVEL%"
pause
exit /b %SYNC_EXIT%

:noenv
echo 未找到 .venv。请先按教程创建虚拟环境并安装依赖。
pause
exit /b 1
:failed
echo 启动失败，请查看上面的错误信息。
pause
exit /b 1
