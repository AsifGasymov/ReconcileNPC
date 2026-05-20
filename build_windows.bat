@echo off
setlocal enabledelayedexpansion

set APP_NAME=NPCMode
set ENTRY=main.py
set ICON=resources\icons\app.ico

echo =^> Installing / upgrading dependencies...
pip install --upgrade pip wheel
pip install -r requirements.txt pyinstaller

echo =^> Building Windows .exe...
pyinstaller ^
  --clean ^
  --noconfirm ^
  --windowed ^
  --name "%APP_NAME%" ^
  --icon "%ICON%" ^
  --add-data "resources;resources" ^
  %ENTRY%

echo =^> Done: dist\%APP_NAME%\%APP_NAME%.exe
pause
