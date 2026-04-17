@echo off
:: pre-build.bat - Git Sicherung vor jedem Build
:: Einbinden: Projekteigenschaften -> Buildereignisse -> Vor dem Buildvorgang
:: $(ProjectDir)pre-build.bat "$(ProjectDir)" "$(ConfigurationName)"

setlocal

set PROJDIR=%~1
set CONFIG=%~2
if "%PROJDIR%"=="" set PROJDIR=%~dp0
if "%CONFIG%"=="" set CONFIG=Debug

:: Ins Projektverzeichnis wechseln
cd /d "%PROJDIR%\.."

:: Prüfen ob Git da ist
git rev-parse --git-dir >nul 2>&1
if errorlevel 1 (
    echo [pre-build] Kein Git-Repository gefunden, Sicherung übersprungen.
    exit /b 0
)

:: Nur Source-Dateien stagen (kein bin/obj)
git add *.cs *.csproj *.resx *.config *.settings *.Designer.cs *.xml *.png *.ico 2>nul
git add Outlook2021TodoAddIn\*.cs 2>nul
git add Outlook2021TodoAddIn\*.csproj 2>nul
git add Outlook2021TodoAddIn\Properties\*.* 2>nul
git add Outlook2021TodoAddIn\Forms\*.* 2>nul
git add Outlook2021TodoAddIn\Resources\*.* 2>nul

:: Prüfen ob es was zu committen gibt
git diff --cached --quiet
if errorlevel 1 (
    :: Zeitstempel als Commit-Message
    for /f "tokens=1-3 delims=. " %%a in ('date /t') do set DATUM=%%c-%%b-%%a
    for /f "tokens=1-2 delims=: " %%a in ('time /t') do set ZEIT=%%a:%%b
    git commit -m "pre-build %CONFIG% %DATUM% %ZEIT%"
    echo [pre-build] Commit erstellt: %CONFIG% %DATUM% %ZEIT%
) else (
    echo [pre-build] Keine Änderungen, kein Commit nötig.
)

exit /b 0