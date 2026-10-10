@echo off
setlocal
cd /d "%~dp0"
echo Easy Buys Aronium Sync Bridge - first run
echo.
where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python 3.11+ and select "Add python.exe to PATH".
  pause
  exit /b 2
)
if not exist config.json (
  copy config.example.json config.json >nul
  echo Created config.json. Edit the test source IDs and database path before continuing.
)
if not exist mapping.json (
  copy mapping.example.json mapping.json >nul
  echo Created mapping.json. All extraction is disabled by default.
)
python bridge.py init-queue
echo.
echo First run complete. Run LAUNCH-BRIDGE.bat to open the interface.
pause
