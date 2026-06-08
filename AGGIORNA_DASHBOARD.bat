@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title Sakarya Dashboard — Aggiornamento (1 clic)
cd /d "%~dp0"

echo.
echo ============================================
echo   SAKARYA DASHBOARD — AGGIORNAMENTO 1 CLIC
echo ============================================
echo.

:: ---- 1. Python ---------------------------------------------------
set PY=python
where python >nul 2>&1 || set PY=py
where %PY% >nul 2>&1
if errorlevel 1 (
    echo [ERRORE] Python non trovato. Installalo da https://python.org
    pause & exit /b 1
)

:: ---- 2. Genera i CSV dall'Excel ----------------------------------
echo [1/4] Lettura Excel e generazione dati...
%PY% "%~dp0genera_dati.py"
if errorlevel 1 (
    echo.
    echo [ERRORE] Generazione dati fallita. Controlla l'Excel.
    pause & exit /b 1
)
echo.

:: ---- 3. Git: c'e' qualcosa da pubblicare? ------------------------
where git >nul 2>&1
if errorlevel 1 (
    echo [ERRORE] Git non trovato. Installa Git da https://git-scm.com
    pause & exit /b 1
)

echo [2/4] Controllo modifiche...
git add -A
git diff --cached --quiet && (
    echo Nessuna modifica: la dashboard e' gia' aggiornata.
    echo.
    pause & exit /b 0
)

:: ---- 4. Commit + push sul branch corrente ------------------------
set FILENAME=aggiornamento dati
for %%f in ("Sakarya Inspection Overall Status as of *.xlsx") do set FILENAME=%%~nf

echo [3/4] Creazione commit...
git commit -m "Aggiornamento dashboard: !FILENAME!"

for /f "delims=" %%b in ('git branch --show-current') do set BRANCH=%%b
echo [4/4] Push su GitHub (branch: !BRANCH!)...
git push origin !BRANCH!
if errorlevel 1 (
    echo.
    echo [ERRORE] Push fallito. Controlla connessione/credenziali GitHub.
    pause & exit /b 1
)

echo.
echo ============================================
echo   COMPLETATO! Dashboard aggiornata online.
echo   (online tra ~1 minuto — ricarica con F5)
echo ============================================
echo.
pause
