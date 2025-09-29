@echo off
setlocal ENABLEEXTENSIONS

where python >nul 2>nul
if errorlevel 1 (
    echo [error] Python 3.10+ is required but not found in PATH.
    exit /b 1
)

echo [info] Installing build dependencies...
python -m pip install --upgrade pip >nul
python -m pip install pyinstaller >nul
python -m pip install -e . >nul

if errorlevel 1 (
    echo [error] Failed to install dependencies.
    exit /b 1
)

echo [info] Building glycoenum executable...
pyinstaller --noconfirm --onefile glycoenum/cli.py -n glycoenum

if exist dist\glycoenum.exe (
    echo [done] Build complete. Artifact: dist\glycoenum.exe
) else (
    echo [warn] PyInstaller finished but dist\glycoenum.exe was not found.
    exit /b 1
)

endlocal
