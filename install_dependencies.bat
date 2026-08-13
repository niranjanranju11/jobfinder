@echo off
setlocal
cd /d "%~dp0"
set "PYTHON_EXE=C:\Users\DELL\AppData\Local\Programs\Python\Python313\python.exe"
if exist "%PYTHON_EXE%" (
  "%PYTHON_EXE%" -m pip install -r requirements.txt
) else (
  python -m pip install -r requirements.txt
)
pause
