@echo off
title Reset Orbit Physics Demo
cd /d "%~dp0"

echo This will delete all local demo attempts, notes and progress.
set /p confirm="Type RESET to continue: "
if /I not "%confirm%"=="RESET" (
  echo Reset cancelled.
  pause
  exit /b 0
)

if exist physics_demo.db del physics_demo.db
echo Demo database removed.
echo Run run_demo.bat to recreate clean sample data.
pause
