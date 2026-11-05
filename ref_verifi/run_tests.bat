@echo off
setlocal
set "PYTHONUTF8=1"
set "PYTHONDONTWRITEBYTECODE=1"
set "VERIFY_PY=%~1"
if not defined VERIFY_PY set "VERIFY_PY=%~dp0.venv\Scripts\python.exe"
if not exist "%VERIFY_PY%" (
  echo Python not found. Pass the full path to your approved project venv python.exe.
  echo Example: run_tests.bat "C:\project\.venv\Scripts\python.exe"
  exit /b 2
)
"%VERIFY_PY%" -X utf8 "%~dp0demo\run_tests.py" --report
exit /b %errorlevel%
