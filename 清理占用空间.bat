@echo off
chcp 65001 >nul
setlocal EnableExtensions
rem ============================================================================
rem 清理本项目占用的多余磁盘空间（安全、可重复运行）。
rem 只删除可再生的内容：
rem   1) .venv_old_*       prepare_env.bat 重建环境时留下的旧虚拟环境备份
rem   2) edge_copy_profile 专用Edge的浏览器缓存（Cache / Code Cache / Service Worker / GPU缓存等）
rem   3) __pycache__       Python 字节码缓存（下次导入自动重建）
rem   4) sync_diff_log.jsonl 差异诊断日志（超过 8 MB 时先压缩归档到 LOG\ 再清空）
rem 不会删除：edge_copy_profile 目录本身与登录状态（根目录 Local State、Network\Cookies、
rem   Local Storage、IndexedDB、Preferences 等）、.venv、DATA\、同步进度文件、LOG\、任何源码。
rem 运行前请关闭 start_ui.bat 服务窗口和专用 Edge，否则脚本会拒绝执行。
rem ============================================================================
cd /d "%~dp0"
if errorlevel 1 (
  echo 无法进入项目目录，请把本脚本放在项目根目录运行。
  pause
  exit /b 1
)
title 清理占用空间

echo ============================================================
echo  表C同步工具 - 磁盘空间清理
echo ============================================================
echo.
echo 正在检查同步程序是否在运行……
powershell -NoProfile -ExecutionPolicy Bypass -Command "$b = Get-CimInstance Win32_Process | Where-Object { ($_.Name -eq 'msedge.exe' -and $_.CommandLine -like '*edge_copy_profile*') -or ($_.Name -eq 'python.exe' -and $_.CommandLine -like '*wps_mysql_sync*') }; if ($b) { exit 1 } else { exit 0 }"
if errorlevel 1 (
  echo.
  echo [已中止] 同步程序或专用 Edge 仍在运行。
  echo 请先关闭 start_ui.bat 服务窗口和专用 Edge 窗口，再重新运行本脚本。
  echo.
  pause
  exit /b 1
)

powershell -NoProfile -Command "$s=(Get-ChildItem -LiteralPath '.' -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum; [math]::Round($s/1MB,1)" > "%TEMP%\wps_size_before.txt" 2>nul
set "BEFORE=0"
set /p BEFORE=<"%TEMP%\wps_size_before.txt"
del "%TEMP%\wps_size_before.txt" >nul 2>&1
echo 清理前占用：%BEFORE% MB
echo.

echo [1/4] 旧虚拟环境备份……
for /d %%d in (.venv_old_*) do (
  echo       删除 %%d
  rd /s /q "%%d" 2>nul
  if exist "%%d" echo       [跳过] 删除失败，可能仍在运行
)

echo [2/4] 专用 Edge 浏览器缓存……
set "PROFILE=edge_copy_profile"
for %%r in (
  "Default\Cache"
  "Default\Code Cache"
  "Default\Service Worker"
  "Default\GPUCache"
  "Default\DawnGraphiteCache"
  "Default\DawnWebGPUCache"
  "GrShaderCache"
  "ShaderCache"
  "GPUPersistentCache"
  "BrowserMetrics"
  "Crashpad"
  "component_crx_cache"
  "extensions_crx_cache"
  "SmartScreen"
) do (
  if exist "%PROFILE%\%%~r" (
    rd /s /q "%PROFILE%\%%~r" 2>nul
    if exist "%PROFILE%\%%~r" (echo       [跳过] %%~r 被占用) else (echo       已清理 %%~r)
  )
)
for %%f in (
  "Default\settings_diagnostic.log.old"
  "Default\settings_diagnostic.log"
  "Default\favorites_diagnostic.log"
  "Default\load_statistics.db"
) do if exist "%PROFILE%\%%~f" del /f /q "%PROFILE%\%%~f" 2>nul

echo [3/4] Python 字节码缓存……
for /d /r "%~dp0." %%d in (__pycache__) do @if exist "%%d" rd /s /q "%%d" 2>nul
echo       已清理

echo [4/4] 差异诊断日志……
powershell -NoProfile -ExecutionPolicy Bypass -Command "$f = 'sync_diff_log.jsonl'; if (Test-Path -LiteralPath $f) { $n = (Get-Item -LiteralPath $f).Length; if ($n -gt 8MB) { New-Item -ItemType Directory -Force -Path 'LOG' | Out-Null; $d = Join-Path 'LOG' ('sync_diff_log_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.jsonl.gz'); $i = [IO.File]::OpenRead((Resolve-Path -LiteralPath $f).Path); $o = [IO.File]::Create($d); $g = New-Object IO.Compression.GZipStream($o, [IO.Compression.CompressionMode]::Compress); $i.CopyTo($g); $g.Dispose(); $o.Dispose(); $i.Dispose(); Set-Content -LiteralPath $f -Value '' -NoNewline; Write-Host ('      已压缩归档 {0:N1} MB 到 {1}' -f ($n/1MB), $d) } else { Write-Host '      日志未超过 8 MB，无需归档' } } else { Write-Host '      日志不存在，无需处理' }"

echo.
powershell -NoProfile -Command "$s=(Get-ChildItem -LiteralPath '.' -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum; [math]::Round($s/1MB,1)" > "%TEMP%\wps_size_after.txt" 2>nul
set "AFTER=0"
set /p AFTER=<"%TEMP%\wps_size_after.txt"
del "%TEMP%\wps_size_after.txt" >nul 2>&1
echo ============================================================
echo  清理前：%BEFORE% MB
echo  清理后：%AFTER% MB
powershell -NoProfile -Command "Write-Host ('  已释放：' + [math]::Round(%BEFORE% - %AFTER%, 1) + ' MB')"
echo ============================================================
echo.
echo 说明：登录状态、表C、同步进度和源码均未改动。
echo 清空缓存后首次打开在线表格会略慢，属于正常现象。
echo.
pause
exit /b 0
