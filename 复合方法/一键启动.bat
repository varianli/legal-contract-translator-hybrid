@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "PARENT_VENV=%~dp0..\.venv\Scripts\python.exe"
set "PYTHON_EXE="
set "PYTHON_ARGS="

python --version >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_EXE=python"
  goto have_python
)
py -3 --version >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_EXE=py"
  set "PYTHON_ARGS=-3"
  goto have_python
)
py --version >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_EXE=py"
  goto have_python
)
if exist "%PARENT_VENV%" (
  set "PYTHON_EXE=%PARENT_VENV%"
  goto have_python
)

echo.
echo Python was not found on this computer.
echo Trying to install Python 3.12 automatically with winget...
winget --version >nul 2>nul
if errorlevel 1 goto python_manual_install

winget install --id Python.Python.3.12 --source winget --exact --accept-package-agreements --accept-source-agreements
if errorlevel 1 goto python_manual_install

python --version >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_EXE=python"
  goto have_python
)
py -3 --version >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_EXE=py"
  set "PYTHON_ARGS=-3"
  goto have_python
)
py --version >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_EXE=py"
  goto have_python
)

echo.
echo Python seems installed, but this window cannot find it yet.
echo Please close this window and double-click this bat file again.
pause
exit /b 1

:python_manual_install
echo.
echo Automatic Python installation was not available.
echo Please install Python 3.12 or newer, then run this bat again.
echo.
echo Recommended steps:
echo 1. Open https://www.python.org/downloads/
echo 2. Download Python for Windows.
echo 3. During installation, tick "Add python.exe to PATH".
echo 4. Finish installation, then double-click this bat again.
echo.
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
