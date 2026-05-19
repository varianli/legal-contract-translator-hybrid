@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "BUNDLED_PY=C:\Users\varianli\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "PYTHON_EXE="
set "PYTHON_ARGS="

if exist "%BUNDLED_PY%" (
  set "PYTHON_EXE=%BUNDLED_PY%"
  goto have_python
)

py -3 --version >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_EXE=py"
  set "PYTHON_ARGS=-3"
  goto have_python
)

python --version >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_EXE=python"
  goto have_python
)

:no_python
echo Python was not found.
echo Please install Python 3.10 or newer, then double-click this file again.
echo Download: https://www.python.org/downloads/
pause
exit /b 1

:have_python
echo Using Python: %PYTHON_EXE% %PYTHON_ARGS%

if not exist ".venv\Scripts\python.exe" (
  echo Creating this tool's private Python environment...
  "%PYTHON_EXE%" %PYTHON_ARGS% -m venv .venv
  if errorlevel 1 (
    echo Failed to create the Python environment.
    pause
    exit /b 1
  )
)

echo Installing/checking dependencies. The first run may take a few minutes...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo Dependency installation failed. Please check your network and try again.
  pause
  exit /b 1
)

if /i "%~1"=="--check" (
  ".venv\Scripts\python.exe" -c "import openai, docx, tkinterdnd2; print('Dependency import check passed.')"
  exit /b %errorlevel%
)

echo Launching the legal contract translator...
".venv\Scripts\python.exe" legal_contract_translator.py
pause
