@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title Sakarya Dashboard - Aggiorna (scegli Excel)
cd /d "%~dp0"

echo.
echo ============================================
echo   SAKARYA DASHBOARD - AGGIORNAMENTO
echo ============================================
echo.
echo Si apre la finestra "Scegli file":
echo seleziona l'Excel "Overall Status" della settimana.
echo.

REM ---- 1. Finestra Scegli file (OpenFileDialog) ----
set "TMPSEL=%TEMP%\sakarya_sel.txt"
del "%TMPSEL%" 2>nul

powershell -NoProfile -STA -Command "Add-Type -AssemblyName System.Windows.Forms; $d=New-Object System.Windows.Forms.OpenFileDialog; $d.Filter='Excel (*.xlsx)|*.xlsx'; $d.Title='Scegli l''Excel Overall Status della settimana'; $dt=[Environment]::GetFolderPath('Desktop'); if(Test-Path $dt){ $d.InitialDirectory=$dt }; if($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){ [System.IO.File]::WriteAllText('%TMPSEL%', $d.FileName) }"

set "SEL="
if exist "%TMPSEL%" set /p SEL=<"%TMPSEL%"
del "%TMPSEL%" 2>nul

if not defined SEL (
    echo Nessun file selezionato. Operazione annullata.
    echo.
    pause
    exit /b 0
)

for %%I in ("!SEL!") do set "BASE=%%~nxI"
echo File scelto: !BASE!
echo.

REM ---- 2. Controllo nome del file ----
echo !BASE! | findstr /i /c:"Overall Status as of" >nul
if errorlevel 1 (
    echo [ATTENZIONE] Il file non sembra un "Overall Status as of ...".
    echo Nome trovato: !BASE!
    echo Per sicurezza non ho toccato niente. Riprova col file giusto.
    echo.
    pause
    exit /b 1
)

REM ---- 3. Python e Git presenti? ----
set "PY=python"
where python >nul 2>&1 || set "PY=py"
where !PY! >nul 2>&1
if errorlevel 1 (
    echo [ERRORE] Python non trovato. Installalo da https://python.org
    pause
    exit /b 1
)
where git >nul 2>&1
if errorlevel 1 (
    echo [ERRORE] Git non trovato. Installa Git da https://git-scm.com
    pause
    exit /b 1
)

REM ---- 4. Inserisci l'Excel nella cartella (gestisce anche lo stesso file) ----
echo [1/4] Inserimento Excel nella cartella...
copy /y "!SEL!" "%~dp0_overall_nuovo.tmp" >nul
if errorlevel 1 (
    echo [ERRORE] Impossibile copiare il file scelto.
    pause
    exit /b 1
)
move /y "%~dp0_overall_nuovo.tmp" "%~dp0!BASE!" >nul

REM ---- 5. Genera i CSV ----
echo [2/4] Generazione dati dall'Excel...
!PY! "%~dp0genera_dati.py"
if errorlevel 1 (
    echo.
    echo [ERRORE] Generazione dati fallita. Controlla l'Excel.
    pause
    exit /b 1
)
echo.

REM ---- 6. Commit ----
echo [3/4] Salvataggio modifiche...
git add -A
git diff --cached --quiet && (
    echo Nessuna modifica: la dashboard era gia' aggiornata.
    echo.
    pause
    exit /b 0
)
git commit -m "Aggiornamento dashboard: !BASE!"

REM ---- 7. Push ----
for /f "delims=" %%b in ('git branch --show-current') do set "BRANCH=%%b"
echo [4/4] Pubblicazione su GitHub ^(branch !BRANCH!^)...
git push origin !BRANCH!
if errorlevel 1 (
    echo.
    echo [ERRORE] Push fallito. Controlla connessione/credenziali GitHub.
    pause
    exit /b 1
)

REM ---- 8. Verifica che GitHub Pages abbia davvero pubblicato ----
echo.
echo [5/5] Verifica pubblicazione online ^(fino a 1 minuto^)...
for /f "delims=" %%s in ('git rev-parse HEAD') do set "COMMITSHA=%%s"
set "TMPVER=%TEMP%\sakarya_verify.txt"
del "%TMPVER%" 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0verifica_deploy_pages.ps1" -Sha "!COMMITSHA!" > "%TMPVER%" 2>nul

set "ESITO="
set "LOGURL="
for /f "tokens=1,* delims=:" %%a in ('findstr /b "STATO: LOG:" "%TMPVER%"') do (
    if "%%a"=="STATO" set "ESITO=%%b"
    if "%%a"=="LOG" set "LOGURL=%%b"
)
del "%TMPVER%" 2>nul

echo.
if "!ESITO!"=="SUCCESS" (
    echo ============================================
    echo   COMPLETATO! Dashboard pubblicata online.
    echo   Ricarica con CTRL+F5.
    echo ============================================
) else if "!ESITO!"=="FAILURE" (
    echo ============================================
    echo   [ATTENZIONE] Il push e' andato a buon fine ma
    echo   la pubblicazione su GitHub Pages e' FALLITA.
    echo   CTRL+F5 non servira' a nulla: il sito online
    echo   non e' stato aggiornato.
    echo.
    echo   Come risolvere:
    echo   1^) Apri questo link e clicca "Re-run failed jobs":
    echo      !LOGURL!
    echo   2^) Se non si risolve, rilancia questo stesso file
    echo      ^(anche con lo stesso Excel^): forzera' un nuovo
    echo      tentativo di pubblicazione.
    echo ============================================
) else (
    echo ============================================
    echo   Push completato, ma la pubblicazione e' ancora
    echo   in corso o non e' stato possibile verificarla
    echo   in tempo. Aspetta 1-2 minuti e ricarica con
    echo   CTRL+F5. Se dopo 5 minuti i dati sono ancora
    echo   vecchi, rilancia questo file.
    echo ============================================
)
echo.
pause
