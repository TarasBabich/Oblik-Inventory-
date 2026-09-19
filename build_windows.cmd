@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ============================================================
echo Oblik Inventory - Windows EXE build
echo ============================================================

where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py"
) else (
  set "PY=python"
)

if not exist ".venv\Scripts\python.exe" (
  %PY% -m venv .venv
  if errorlevel 1 goto :error
)

".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :error
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :error

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

".venv\Scripts\python.exe" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name Oblik ^
  --collect-all PySide6 ^
  --collect-all pandas ^
  src\main.py
if errorlevel 1 goto :error

echo.
echo ГОТОВО: dist\Oblik.exe
pause
exit /b 0

:error
echo.
echo ПОМИЛКА ЗБІРКИ.
pause
exit /b 1
