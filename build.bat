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

set TARGET_NAME=OligosaccharidePrediction

echo [info] Building %TARGET_NAME% executable...
pyinstaller --noconfirm --onefile --windowed glycoenum/gui.py -n %TARGET_NAME%

if exist dist\%TARGET_NAME%.exe (
    echo [done] Build complete. Artifact: dist\%TARGET_NAME%.exe
) else (
    echo [warn] PyInstaller finished but dist\%TARGET_NAME%.exe was not found.
    exit /b 1
)

endlocal
