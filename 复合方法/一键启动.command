#!/bin/bash
set -u

cd "$(dirname "$0")" || exit 1
export TK_SILENCE_DEPRECATION=1

PYTHON_EXE=""
if command -v python3 >/dev/null 2>&1; then
  PYTHON_EXE="$(command -v python3)"
elif command -v python >/dev/null 2>&1; then
  PYTHON_EXE="$(command -v python)"
fi

if [ -z "$PYTHON_EXE" ]; then
  echo
  echo "Python 3 was not found on this Mac."
  echo "Please install Python 3.12 or newer from:"
  echo "https://www.python.org/downloads/macos/"
  echo
  if command -v open >/dev/null 2>&1; then
    open "https://www.python.org/downloads/macos/" >/dev/null 2>&1 || true
  fi
  echo "After installation, double-click this command file again."
  echo
  read -r -p "Press Enter to close..."
  exit 1
fi

if ! "$PYTHON_EXE" - <<'PY'
import sys
raise SystemExit(0 if sys.version_info >= (3, 10) else 1)
PY
then
  echo
  echo "Python is too old. Please install Python 3.12 or newer:"
  echo "https://www.python.org/downloads/macos/"
  echo
  read -r -p "Press Enter to close..."
  exit 1
fi

if [ ! -x ".venv/bin/python" ]; then
  echo "Creating private Python environment..."
  "$PYTHON_EXE" -m venv .venv
  if [ $? -ne 0 ]; then
    echo
    echo "Failed to create the private Python environment."
    echo "Please install Python from python.org, then try again."
    echo
    read -r -p "Press Enter to close..."
    exit 1
  fi
fi

echo "Installing/checking dependencies..."
".venv/bin/python" -m ensurepip --upgrade >/dev/null 2>&1 || true
".venv/bin/python" -m pip install --upgrade pip
".venv/bin/python" -m pip install -r requirements.txt
if [ $? -ne 0 ]; then
  echo
  echo "Dependency installation failed."
  echo "Please check the network connection, then try again."
  echo
  read -r -p "Press Enter to close..."
  exit 1
fi

if [ "${1:-}" = "--check" ]; then
  ".venv/bin/python" -c "import openai, docx, tkinterdnd2; print('Dependency import check passed.')"
  exit $?
fi

".venv/bin/python" hybrid_markdown_run_translator.py
echo
read -r -p "Press Enter to close..."
