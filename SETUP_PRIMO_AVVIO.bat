@echo off
chcp 65001 >nul
title Sakarya Dashboard - Installazione (primo avvio)
cd /d "%~dp0"
set "NEEDRESTART=0"

echo ============================================
echo   SAKARYA DASHBOARD - INSTALLAZIONE
echo ============================================
echo.
echo Questo prepara il PC con tutto il necessario:
echo   - Python
echo   - Git
echo   - libreria openpyxl
echo   - icona "Sakarya Dashboard" sul Desktop
echo.
echo Potrebbe comparire una richiesta di Windows: accetta.
echo Se installa qualcosa, ti chiedera' di rilanciare il SETUP.
echo.
pause
echo.

REM =================== PYTHON ===================
where python >nul 2>&1 && goto PY_OK
where py >nul 2>&1 && goto PY_OK
echo [Python] non trovato: lo installo...
where winget >nul 2>&1 || goto NO_WINGET_PY
winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements
set "NEEDRESTART=1"
goto GIT_CHECK
:NO_WINGET_PY
echo.
echo [!] Installazione automatica non disponibile su questo PC.
echo     Apro la pagina di Python: installalo spuntando
echo     "Add Python to PATH", poi rilancia questo SETUP.
start "" https://www.python.org/downloads/
echo.
pause
exit /b 1
:PY_OK
echo [Python] gia' presente. OK.

REM =================== GIT ===================
:GIT_CHECK
where git >nul 2>&1 && goto GIT_OK
echo [Git] non trovato: lo installo...
where winget >nul 2>&1 || goto NO_WINGET_GIT
winget install -e --id Git.Git --accept-package-agreements --accept-source-agreements
set "NEEDRESTART=1"
goto AFTER_INSTALL
:NO_WINGET_GIT
echo.
echo [!] Installazione automatica non disponibile su questo PC.
echo     Apro la pagina di Git: installalo con le opzioni di
echo     default, poi rilancia questo SETUP.
start "" https://git-scm.com/download/win
echo.
pause
exit /b 1
:GIT_OK
echo [Git] gia' presente. OK.

:AFTER_INSTALL
if "%NEEDRESTART%"=="1" goto RESTART

REM =================== OPENPYXL ===================
set "PY=python"
where python >nul 2>&1 || set "PY=py"
echo.
echo [openpyxl] installo/aggiorno la libreria...
%PY% -m pip install --upgrade openpyxl
if errorlevel 1 goto PIP_ERR

REM =================== SCORCIATOIA DESKTOP ===================
echo [Icona] creo la scorciatoia sul Desktop...
powershell -NoProfile -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\Sakarya Dashboard.lnk'); $s.TargetPath='%~dp0AVVIA_GUI.bat'; $s.WorkingDirectory='%~dp0'; $s.IconLocation='%SystemRoot%\System32\shell32.dll,220'; $s.Save()"

echo.
echo ============================================
echo   INSTALLAZIONE COMPLETATA!
echo ============================================
echo.
echo Resta UNA cosa che nessun installer puo' fare al posto tuo:
echo   - Chi gestisce il progetto deve aggiungerti come
echo     COLLABORATORE al repository GitHub (serve il tuo nome
echo     utente GitHub). Senza, il tasto "Pubblica" dara' errore.
echo   - Al primo "Pubblica", il browser ti chiedera' il login
echo     GitHub: accetta una volta e poi non lo richiede piu'.
echo.
echo Per usare la dashboard: doppio clic sull'icona
echo "Sakarya Dashboard" sul Desktop.
echo.
choice /C SN /M "Vuoi aprire ora il programma (S/N)"
if errorlevel 2 goto END
start "" pythonw "%~dp0app_dashboard_gui.py"
goto END

:RESTART
echo.
echo ============================================
echo   PRIMO PASSO FATTO - RILANCIA IL SETUP
echo ============================================
echo.
echo Python e/o Git sono stati appena installati, ma Windows
echo li "vede" solo in una nuova finestra.
echo.
echo  1) CHIUDI questa finestra
echo  2) Fai doppio clic di NUOVO su SETUP_PRIMO_AVVIO.bat
echo     per completare (openpyxl + icona Desktop).
echo.
pause
exit /b 0

:PIP_ERR
echo.
echo [ERRORE] Installazione di openpyxl fallita.
echo Controlla la connessione a internet e rilancia il SETUP.
echo.
pause
exit /b 1

:END
echo.
pause
