@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo ERREUR : Python n'est pas installe ou n'est pas dans le PATH.
    echo Installe Python 3.10+ depuis https://www.python.org/downloads/ ^(coche "Add Python to PATH"^).
    pause
    exit /b 1
)

echo Installation des dependances...
python -m pip install -r requirements.txt --quiet

echo.
python codifier.py

echo.
pause
