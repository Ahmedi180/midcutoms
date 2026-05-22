@echo off
REM One-click run helper for Windows (Command Prompt)
REM Creates a virtual environment (if missing), activates it,
REM installs dependencies from requirements.txt, and starts the dev server.

SETLOCAL ENABLEEXTENSIONS

python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
  echo Python not found on PATH. Please install Python 3.10+ and try again.
  pause
  exit /b 1
)

IF NOT EXIST "venv\Scripts\python.exe" (
  echo Creating virtual environment...
  python -m venv venv
  IF %ERRORLEVEL% NEQ 0 (
    echo Failed to create virtual environment.
    pause
    exit /b 1
  )
)

echo Activating virtual environment...
call venv\Scripts\activate.bat

echo Upgrading pip...
python -m pip install --upgrade pip >nul

IF EXIST "requirements.txt" (
  echo Installing dependencies from requirements.txt...
  pip install -r requirements.txt
)

echo Starting development server (Ctrl+C to stop)...
python app.py

ENDLOCAL
