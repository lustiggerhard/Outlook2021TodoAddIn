@echo off
:: post-publish.bat - Git Push nach jedem Veröffentlichen
:: Projekteigenschaften -> Buildereignisse -> Nach dem Buildvorgang:
:: "$(ProjectDir)post-publish.bat" "$(ProjectDir)"

setlocal

set PROJDIR=%~1
if "%PROJDIR%"=="" set PROJDIR=%~dp0

cd /d "%PROJDIR%\.."

git rev-parse --git-dir >nul 2>&1
if errorlevel 1 (
    echo [post-publish] Kein Git-Repository gefunden.
    exit /b 0
)

git push
if errorlevel 1 (
    echo [post-publish] FEHLER beim Push!
    exit /b 1
)

echo [post-publish] Push erfolgreich.
exit /b 0