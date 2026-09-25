@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ============================================================
echo Oblik Inventory - Flet Windows EXE build
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

".venv\Scripts\flet.exe" pack ^
  --yes ^
  --name Oblik ^
  --distpath dist ^
  --product-name "Oblik Inventory" ^
  --product-version 0.2.4 ^
  --file-version 0.2.4.0 ^
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
