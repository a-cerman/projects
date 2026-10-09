@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
 echo Complete the setup in README.md first.
 pause
 exit /b 1
)
".venv\Scripts\python.exe" app.py
if errorlevel 1 pause
