@echo off
rem Build DiplomGenerator.exe on Windows (run from the project folder).
rem Requires Python 3.11+ installed with "py" launcher (python.org installer).
py -3 -m venv .venv
call .venv\Scripts\activate.bat
pip install -r requirements.txt
pyinstaller diplom.spec --noconfirm
echo.
echo Ready: dist\DiplomGenerator.exe
pause
