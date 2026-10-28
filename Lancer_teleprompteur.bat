@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set PY=py& set PYW=pyw) else (set PY=python& set PYW=pythonw)
echo Installation / verification des modules...
%PY% -m pip install --quiet --disable-pip-version-check -r requirements.txt
start "" %PYW% teleprompteur.py %*
