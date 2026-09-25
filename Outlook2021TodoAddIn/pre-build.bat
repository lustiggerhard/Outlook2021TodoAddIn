@echo off
:: @file       pre-build.bat
:: @brief      Git-Sicherung vor jedem Build (vollstaendig inkl. bin/obj)
:: @author     Gerhard Lustig <gerhard@lustig.at>
:: @version    1.1.0
:: @date       2026-09-25
:: @history
::   1.1.0 (2026-09-25) - Fix: Datum/Uhrzeit vor dem if-Block ermitteln; %DATUM%/%ZEIT% wurden
::                        im Block beim Parsen expandiert und waren in der Commit-Message leer
::   1.0.0              - Initial release
::
:: Projekteigenschaften -> Buildereignisse -> Vor dem Buildvorgang:
:: "$(ProjectDir)pre-build.bat" "$(ProjectDir)" "$(ConfigurationName)"

setlocal

set PROJDIR=%~1
set CONFIG=%~2
if "%PROJDIR%"=="" set PROJDIR=%~dp0
if "%CONFIG%"=="" set CONFIG=Debug

cd /d "%PROJDIR%\.."

git rev-parse --git-dir >nul 2>&1
if errorlevel 1 (
    echo [pre-build] Kein Git-Repository gefunden.
    exit /b 0
)

:: Datum/Uhrzeit ausserhalb jedes Klammerblocks setzen, damit %DATUM%/%ZEIT% unten gefuellt sind
for /f "tokens=1,2" %%a in ('powershell -NoProfile -Command "Get-Date -Format 'dd.MM.yyyy HH:mm'"') do (
    set DATUM=%%a
    set ZEIT=%%b
)

:: Alles stagen außer temporäre VS-Dateien
git add -A
git reset HEAD -- "*.suo" "*.user" ".vs/" 2>nul

:: Prüfen ob es was zu committen gibt
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "pre-build %CONFIG% %DATUM% %ZEIT%"
    echo [pre-build] Commit: %CONFIG% %DATUM% %ZEIT%
) else (
    echo [pre-build] Keine Änderungen.
)

exit /b 0
