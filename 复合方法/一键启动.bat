@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "PARENT_VENV=%~dp0..\.venv\Scripts\python.exe"
set "BUNDLED_PY=C:\Users\varianli\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "PYTHON_EXE="
set "PYTHON_ARGS="

if exist "%PARENT_VENV%" (
  set "PYTHON_EXE=%PARENT_VENV%"
  goto have_python
)
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
echo Python was not found.
pause
exit /b 1

:have_python
if not exist ".venv\Scripts\python.exe" (
  echo Creating private Python environment...
  "%PYTHON_EXE%" %PYTHON_ARGS% -m venv .venv
  if errorlevel 1 (
    echo Failed to create environment.
    pause
    exit /b 1
  )
)

echo Installing/checking dependencies...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo Dependency installation failed.
  pause
  exit /b 1
)

if /i "%~1"=="--check" (
  ".venv\Scripts\python.exe" -c "import openai, docx, tkinterdnd2; print('Dependency import check passed.')"
  exit /b %errorlevel%
)

".venv\Scripts\python.exe" hybrid_markdown_run_translator.py
pause
