@echo off
title Orbit Physics Demo
cd /d "%~dp0"

echo.
echo  Orbit Physics - Local Demo
echo  ==========================
echo.
echo  Preparing the application...
python -m pip install -r requirements.txt --quiet
if errorlevel 1 (
  echo.
  echo  Dependency installation failed.
  echo  Check your internet connection for the first setup.
  pause
  exit /b 1
)

echo  Opening http://127.0.0.1:8000
echo  Keep this window open during the demo.
echo  Press Ctrl+C to stop the app.
echo.

start "" powershell -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 2; Start-Process 'http://127.0.0.1:8000'"
python -m uvicorn app:app --host 127.0.0.1 --port 8000

pause
