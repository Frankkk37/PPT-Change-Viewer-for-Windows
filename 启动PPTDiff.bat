@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo 正在启动 PPT Diff Tool v0.9 Windows MVP...
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    py -3 ppt_diff_tool.py --ui
    goto end
)

where python >nul 2>nul
if %errorlevel%==0 (
    python ppt_diff_tool.py --ui
    goto end
)

echo 未检测到 Python。
echo.
echo 请先安装 Python 3：
echo https://www.python.org/downloads/windows/
echo.
echo 安装时请勾选 "Add python.exe to PATH"。
echo.
pause

:end
