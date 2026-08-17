@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" "Codings\在线采集入口.py"
    set "COLLECT_EXIT_CODE=!ERRORLEVEL!"
    goto finish
)

if exist "D:\VSCode\python.exe" (
    "D:\VSCode\python.exe" "Codings\在线采集入口.py"
    set "COLLECT_EXIT_CODE=!ERRORLEVEL!"
    goto finish
)

where py >nul 2>nul
if not errorlevel 1 (
    py -3.13 "Codings\在线采集入口.py"
    set "COLLECT_EXIT_CODE=!ERRORLEVEL!"
    goto finish
)

where python >nul 2>nul
if not errorlevel 1 (
    python "Codings\在线采集入口.py"
    set "COLLECT_EXIT_CODE=!ERRORLEVEL!"
    goto finish
)

echo 未找到Python。请先安装Python 3.13并安装requirements.txt中的依赖。
set "COLLECT_EXIT_CODE=1"

:finish
echo.
pause
exit /b %COLLECT_EXIT_CODE%
