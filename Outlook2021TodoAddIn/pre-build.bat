@echo off
:: pre-build.bat - Git Sicherung vor jedem Build (vollständig inkl. bin/obj)
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

:: Alles stagen außer temporäre VS-Dateien
git add -A
git reset HEAD -- "*.suo" "*.user" ".vs/" 2>nul

:: Prüfen ob es was zu committen gibt
git diff --cached --quiet
if errorlevel 1 (
    for /f "tokens=1-3 delims=." %%a in ('powershell -NoProfile -Command "Get-Date -Format dd.MM.yyyy"') do set DATUM=%%a.%%b.%%c
    for /f %%a in ('powershell -NoProfile -Command "Get-Date -Format HH:mm"') do set ZEIT=%%a
    git commit -m "pre-build %CONFIG% %DATUM% %ZEIT%"
    echo [pre-build] Commit: %CONFIG% %DATUM% %ZEIT%
) else (
    echo [pre-build] Keine Änderungen.
)

exit /b 0