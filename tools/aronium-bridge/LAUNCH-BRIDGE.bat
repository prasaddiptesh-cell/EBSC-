@echo off
setlocal
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python 3.11+ and select "Add python.exe to PATH".
  pause
  exit /b 2
)
python bridge_ui.py
if errorlevel 1 (
  echo The bridge exited with an error. Try running "python bridge_ui.py" from Command Prompt.
  pause
)
