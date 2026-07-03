@echo off
setlocal enabledelayedexpansion

cd /d "%~dp0"

set VENV_DIR=

if exist "venv\Scripts\activate.bat" set VENV_DIR=venv
if exist ".venv\Scripts\activate.bat" set VENV_DIR=.venv
if exist "env\Scripts\activate.bat" set VENV_DIR=env

if "%VENV_DIR%"=="" (
    echo Virtual environment not found in this folder.
    echo Expected folder: venv, .venv or env
    pause
    exit /b 1
)

call "%VENV_DIR%\Scripts\activate.bat"

python main.py

pause
endlocal