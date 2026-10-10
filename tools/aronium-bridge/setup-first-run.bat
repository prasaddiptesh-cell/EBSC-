@echo off
setlocal
cd /d "%~dp0"
if not exist "config.json" (
  copy "config.example.json" "config.json" >nul
  echo Created config.json. Edit the shop, source ID and database path.
  notepad "config.json"
)
python --version
if errorlevel 1 (
  echo Python 3.11 or newer was not found. Install Python and try again.
  pause
  exit /b 1
)
python bridge.py validate
if errorlevel 1 (
  echo Validation failed. Check the message and config.json.
  pause
  exit /b 1
)
python bridge.py init-queue
pause
