@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [1/3] Installing PyInstaller...
py -3 -m pip install -r requirements.txt
if errorlevel 1 (
    echo Failed to install requirements.
    pause
    exit /b 1
)

echo [2/3] Building EXE...
py -3 -m PyInstaller PPTDiffTool_v0_9.spec --clean --noconfirm
if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

echo [3/3] Done.
echo EXE path:
echo %cd%\dist\PPTDiffTool_v0_9.exe
pause
