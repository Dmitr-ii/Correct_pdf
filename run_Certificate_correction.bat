

@echo off
SETLOCAL

:: ========== CONFIGURATION ==========
set PYTHON_EXE=python

set VENV_ACTIVATE=C:\Users\Samtsovd\PycharmProjects\CorrectPDF\venv\Scripts\activate.bat
set SCRIPT_PATH=C:\Users\Samtsovd\PycharmProjects\CorrectPDF\certs_to_correct.py
set REQUIREMENTS=C:\Users\Samtsovd\PycharmProjects\CorrectPDF\requirements.txt

:: ========== DEBUG INFO ==========
echo.
echo [DEBUG] Python Path:
where %PYTHON_EXE%
echo.
echo [DEBUG] Python Version:
%PYTHON_EXE% --version
echo.
if exist %SCRIPT_PATH% (
    echo Script found at: %SCRIPT_PATH%
) else (
    echo ERROR: Script not found at %SCRIPT_PATH%
)
echo.


:: ========== EXECUTION ==========
:: 1. Activate venv if exists
if exist %VENV_ACTIVATE% (
    echo Activating virtual environment...
    call %VENV_ACTIVATE%
)

:: 2. Install requirements if file exists
if exist %REQUIREMENTS% (
    echo Installing dependencies...
    %PYTHON_EXE% -m pip install -r %REQUIREMENTS%
)

:: 3. Run script with error handling
echo.
echo Starting script execution...
%PYTHON_EXE% %SCRIPT_PATH%
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Script exited with code %errorlevel%
) else (
    echo.
    echo Script completed successfully
)


--