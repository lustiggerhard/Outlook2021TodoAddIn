@echo off
:: @file       build.cmd
:: @brief      Build + ClickOnce-Publish (Release) fuer Outlook2021TodoAddIn
:: @author     Gerhard Lustig <gerhard@lustig.at>
:: @version    1.3.0
:: @date       2026-09-25
:: @history
::   1.3.0 (2026-09-25) - Alte Versionsordner in "Application Files" loeschen (nur aktuelle bleibt)
::   1.2.0 (2026-09-25) - app.publish nach PublishUrl kopieren (MSBuild-CLI kopiert nicht, nur die VS-IDE)
::   1.1.1 (2026-09-25) - Fix: vswhere-Pfad im echo quoten, "(x86)" beendete den if-Block
::   1.1.0 (2026-09-25) - pre-build.bat (Git-Sicherung) vor dem Build aufrufen
::   1.0.0 (2026-09-25) - Initial release

setlocal EnableExtensions

:: Projektverzeichnis = Ordner dieser Datei
set "PROJDIR=%~dp0"
set "CSPROJ=%PROJDIR%Outlook2021TodoAddIn.csproj"

:: MSBuild ueber vswhere finden (VS 2017+ Standardpfad des Installers)
set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" (
    echo [build] FEHLER: vswhere.exe nicht gefunden: "%VSWHERE%"
    exit /b 1
)
set "MSBUILD="
for /f "usebackq delims=" %%i in (`"%VSWHERE%" -latest -products * -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe`) do set "MSBUILD=%%i"
if not defined MSBUILD (
    echo [build] FEHLER: MSBuild.exe nicht gefunden.
    exit /b 1
)
echo [build] MSBuild: %MSBUILD%

:: Git-Sicherung des aktuellen Stands vor dem Build (Commit "pre-build Release ...")
call "%PROJDIR%pre-build.bat" "%PROJDIR%." "Release"
if errorlevel 1 (
    echo [build] FEHLER: pre-build.bat fehlgeschlagen.
    exit /b 1
)

:: ApplicationRevision hochzaehlen (AutoIncrementApplicationRevision greift nur in der VS-IDE, nicht bei MSBuild-CLI)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$p='%CSPROJ%'; $x=Get-Content -Raw -Encoding UTF8 $p; $m=[regex]::Match($x,'<ApplicationVersion>(\d+)\.(\d+)\.(\d+)\.(\d+)</ApplicationVersion>'); if(-not $m.Success){ Write-Host '[build] FEHLER: ApplicationVersion nicht gefunden'; exit 1 }; $v='{0}.{1}.{2}.{3}' -f $m.Groups[1].Value,$m.Groups[2].Value,$m.Groups[3].Value,([int]$m.Groups[4].Value+1); $x=$x.Replace($m.Value,'<ApplicationVersion>'+$v+'</ApplicationVersion>'); [IO.File]::WriteAllText($p,$x,(New-Object Text.UTF8Encoding $true)); Write-Host ('[build] ApplicationVersion: ' + $v)"
if errorlevel 1 exit /b 1

:: Rebuild + Publish Release; MSBuild-CLI erzeugt das Paket nur in bin\Release\app.publish\
"%MSBUILD%" "%CSPROJ%" /t:Rebuild;Publish /p:Configuration=Release /p:Platform=AnyCPU /m /nologo /v:minimal
if errorlevel 1 (
    echo [build] FEHLER: Build/Publish fehlgeschlagen.
    exit /b 1
)

:: app.publish nach PublishUrl der .csproj kopieren; /E ohne /MIR = alte Versionsordner bleiben erhalten
robocopy "%PROJDIR%bin\Release\app.publish" "D:\temp\publish" /E /R:2 /W:2 /NP /NFL /NDL /NJH
:: robocopy: Exitcode 0-7 = OK, ab 8 = Fehler
if errorlevel 8 (
    echo [build] FEHLER: Kopieren nach D:\temp\publish fehlgeschlagen.
    exit /b 1
)

:: Alte Versionsordner loeschen; nur wenn der Ordner der aktuellen Version im jeweiligen Verzeichnis existiert
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$x=Get-Content -Raw -Encoding UTF8 '%CSPROJ%'; $m=[regex]::Match($x,'<ApplicationVersion>([\d\.]+)</ApplicationVersion>'); if(-not $m.Success){ Write-Host '[build] FEHLER: ApplicationVersion nicht gefunden'; exit 1 }; $cur='Outlook2021TodoAddIn_' + $m.Groups[1].Value.Replace('.','_'); foreach($d in @('%PROJDIR%bin\Release\app.publish\Application Files','D:\temp\publish\Application Files')){ if(-not (Test-Path -LiteralPath (Join-Path $d $cur))){ Write-Host ('[build] FEHLER: ' + $cur + ' fehlt in ' + $d + ' - nichts geloescht'); exit 1 }; Get-ChildItem -LiteralPath $d -Directory -Filter 'Outlook2021TodoAddIn_*' | Where-Object { $_.Name -ne $cur } | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Recurse -Force; Write-Host ('[build] Geloescht: ' + $_.FullName) } }"
if errorlevel 1 exit /b 1

echo [build] OK - Release gebaut und nach D:\temp\publish\ veroeffentlicht.
exit /b 0
