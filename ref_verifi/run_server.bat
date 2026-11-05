@echo off
setlocal
set "PYTHONUTF8=1"
set "PYTHONDONTWRITEBYTECODE=1"
set "VERIFY_PY=%~1"
if not defined VERIFY_PY set "VERIFY_PY=%~dp0.venv\Scripts\python.exe"
if not exist "%VERIFY_PY%" (
  echo Python not found. Pass the full path to your approved project venv python.exe.
  echo Example: run_server.bat "C:\project\.venv\Scripts\python.exe"
  exit /b 2
)
"%VERIFY_PY%" -X utf8 "%~dp0demo\server.py" --host 127.0.0.1 --port 8000 --verbose
exit /b %errorlevel%
