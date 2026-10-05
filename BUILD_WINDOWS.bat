@echo off
setlocal
cd /d "%~dp0"
py -m pip install --upgrade pyinstaller openpyxl
pyinstaller --noconfirm --clean --windowed --onefile --name "Noor-ul-Hifz" --icon "icon.ico" main.py
if errorlevel 1 (
  echo.
  echo BUILD FAILED.
  pause
  exit /b 1
)
echo.
echo BUILD COMPLETE: dist\Noor-ul-Hifz.exe
pause
