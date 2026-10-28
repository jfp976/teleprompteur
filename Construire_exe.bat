@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set PY=py) else (set PY=python)
echo === Installation des modules ===
%PY% -m pip install --quiet --disable-pip-version-check -r requirements.txt pyinstaller
if errorlevel 1 goto erreur
echo === Construction de Teleprompteur.exe (1 a 3 minutes) ===
%PY% -m PyInstaller --noconfirm --clean --onefile --windowed --name Teleprompteur --collect-data arabic_reshaper --hidden-import win32com.client teleprompteur.py
if errorlevel 1 goto erreur
copy /y dist\Teleprompteur.exe . >nul
echo.
echo === Termine : Teleprompteur.exe est dans ce dossier ===
pause
exit /b 0
:erreur
echo.
echo *** La construction a echoue, voir les messages ci-dessus ***
pause
exit /b 1
