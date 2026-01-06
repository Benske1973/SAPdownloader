@echo off
setlocal enabledelayedexpansion

REM Build Windows EXE with PyInstaller
REM Output: .\dist\EquansSAPDownloader.exe

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

REM Optional: only needed if you use Playwright-managed browsers.
REM This app uses channel="msedge" (system Edge), so this is usually not required.
REM python -m playwright install

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name "EquansSAPDownloader" ^
  --collect-all playwright ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  sapdownloadergui.py

echo.
echo Done. EXE is in: dist\EquansSAPDownloader.exe
endlocal
