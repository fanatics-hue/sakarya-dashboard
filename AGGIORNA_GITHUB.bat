@echo off
chcp 65001 >nul
title Sakarya Dashboard — Aggiornamento GitHub

echo.
echo ============================================
echo   SAKARYA DASHBOARD — SYNC CON GITHUB
echo ============================================
echo.

cd /d "%~dp0"

:: Controlla se git è disponibile
where git >nul 2>&1
if errorlevel 1 (
    echo [ERRORE] Git non trovato. Installa Git da https://git-scm.com
    pause
    exit /b 1
)

:: Mostra i file modificati
echo [1/4] Controllo file modificati...
git status --short
echo.

:: Verifica che ci sia qualcosa da committare
git diff --quiet && git diff --cached --quiet && (
    echo Nessuna modifica rilevata. Niente da aggiornare.
    echo.
    pause
    exit /b 0
)

:: Aggiunge tutti i file
echo [2/4] Aggiunta file...
git add .

:: Crea commit con data automatica
set OGGI=%date:~6,4%-%date:~3,2%-%date:~0,2%
echo [3/4] Creazione commit...
git commit -m "Aggiornamento dati settimanale — %OGGI%"

:: Push su GitHub
echo [4/4] Push su GitHub...
git push origin main 2>&1
if errorlevel 1 (
    echo.
    echo [ATTENZIONE] Push fallito su 'main'. Provo con 'master'...
    git push origin master 2>&1
    if errorlevel 1 (
        echo.
        echo [ERRORE] Push fallito. Controlla la connessione e le credenziali GitHub.
        echo Potresti dover fare login con: git config --global credential.helper manager
        pause
        exit /b 1
    )
)

echo.
echo ============================================
echo   COMPLETATO! Dashboard aggiornato su GitHub
echo ============================================
echo.
pause
