# Outlook2021TodoAddIn

VSTO-Add-In für Outlook 2021, das Termine und Aufgaben in einem kombinierten Custom Task Pane darstellt.

---

## 🇩🇪 Deutsch

### Überblick

Das Add-In erweitert Outlook 2021 um ein Custom Task Pane mit konsolidierter Tages- und Wochenübersicht – Termine und Aufgaben in einem Bedienelement, ohne ständigen Wechsel zwischen Kalender- und Aufgaben-Modul.

### Features

- Kombinierte Anzeige von Terminen (Appointments) und Aufgaben (Tasks)
- Custom Task Pane, dockbar im Outlook-Hauptfenster
- Automatische Aktualisierung bei Änderungen am Kalender / an Tasks
- Signierbar mit eigenem Code-Signing-Zertifikat (EKU OID `1.3.6.1.5.5.7.3.3`)
- Lokale Konfiguration, kein externer Service erforderlich

### Voraussetzungen

| Komponente | Version |
|---|---|
| Outlook | 2021 (Desktop, MSI oder C2R) |
| .NET Framework | 4.8 |
| VSTO Runtime | 10.0 oder höher |
| Visual Studio | 2022 mit Workload *Office/SharePoint development* |
| Code-Signing-Zertifikat | Eigenes Zertifikat mit EKU Code Signing (selbst-signiert, interne PKI oder kommerzielle CA) |

### Installation

1. Release-ZIP entpacken nach `C:\@Tools\Outlook2021TodoAddIn\`
2. `setup.exe` ausführen (ClickOnce / VSTO-Installer)
3. Outlook neu starten
4. Add-In erscheint unter *Datei → Optionen → Add-Ins → COM-Add-Ins*

### Build

```cmd
cd D:\VisualStudio\Outlook2021TodoAddIn
msbuild Outlook2021TodoAddIn.sln /p:Configuration=Release /p:Platform="Any CPU"
```

Signierung der erzeugten DLL:

```cmd
signtool sign /f cert.pfx /p <PFX-Passwort> /t http://timestamp.digicert.com /fd SHA256 bin\Release\Outlook2021TodoAddIn.dll
```

### Zertifikats-Trust (einmalig pro Client)

```powershell
# Eigene Root CA in Trusted Root einbinden (Pfad zur eigenen CA anpassen)
Import-Certificate -FilePath "C:\Pfad\zu\RootCA.crt" -CertStoreLocation Cert:\LocalMachine\Root

# Eigenes Code-Signing-Zertifikat als Trusted Publisher
Import-Certificate -FilePath "C:\Pfad\zu\codesign.cer" -CertStoreLocation Cert:\LocalMachine\TrustedPublisher

# CRL-Check für Outlook-Makros deaktivieren (Umgebung ohne CRL-Distribution)
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security" `
                 -Name "DontCheckPublisherCertRevocation" -Value 1 -PropertyType DWORD -Force
```

### Projektstruktur

```
Outlook2021TodoAddIn/
├── ThisAddIn.cs                  # VSTO-Einstiegspunkt
├── TaskPaneControl.xaml(.cs)     # WPF-Host für Task Pane
├── AppointmentsControl.cs        # Termin-Logik (Backup-Pflicht: .bak vor Änderung)
├── TasksControl.cs               # Aufgaben-Logik
├── Properties/
│   └── AssemblyInfo.cs           # Version, Copyright
├── build.bat                     # Build- und Sign-Wrapper
└── README.md
```

### Wichtige Konventionen

- **Backup-Pflicht:** Vor jeder Änderung an `AppointmentsControl.cs` → `AppointmentsControl.cs.bak` anlegen
- **Versionierung:** SemVer in `AssemblyInfo.cs`, Hochzählen bei jedem Release
- **Header:** Jede `.cs`-Datei mit Doku-Block (`@file`, `@brief`, `@author`, `@version`, `@date`, `@history` – neueste zuerst)

### Bekannte Probleme

| Problem | Ursache | Lösung |
|---|---|---|
| Add-In wird beim Start deaktiviert | Slow Add-In Detection | Registry: `HKCU\...\Resiliency\DoNotDisableAddinList` setzen |
| Makro-Warnung trotz Signatur | CRL nicht erreichbar | `DontCheckPublisherCertRevocation = 1` (siehe oben) |
| Task Pane leer nach Outlook-Update | VSTO-Cache | `%LOCALAPPDATA%\assembly\dl3` löschen, Outlook neu starten |

### Lizenz

Internes Tool, keine öffentliche Lizenz. Alle Rechte vorbehalten.

### Kontakt

**Gerhard Lustig** · gerhard@lustig.at · netz24.at

---

## 🇬🇧 English

### Overview

VSTO add-in for Outlook 2021 providing a Custom Task Pane with a combined view of appointments and tasks – no more switching between Calendar and Tasks modules.

### Features

- Combined display of Appointments and Tasks
- Dockable Custom Task Pane inside the main Outlook window
- Auto-refresh on calendar/task changes
- Signable with your own code-signing certificate (EKU OID `1.3.6.1.5.5.7.3.3`)
- Fully local, no external service required

### Requirements

| Component | Version |
|---|---|
| Outlook | 2021 (Desktop, MSI or C2R) |
| .NET Framework | 4.8 |
| VSTO Runtime | 10.0+ |
| Visual Studio | 2022 with *Office/SharePoint development* workload |
| Code-Signing Certificate | Your own certificate with Code Signing EKU (self-signed, internal PKI, or commercial CA) |

### Installation

1. Extract release ZIP to `C:\@Tools\Outlook2021TodoAddIn\`
2. Run `setup.exe` (ClickOnce / VSTO installer)
3. Restart Outlook
4. Verify under *File → Options → Add-Ins → COM Add-Ins*

### Build

```cmd
cd D:\VisualStudio\Outlook2021TodoAddIn
msbuild Outlook2021TodoAddIn.sln /p:Configuration=Release /p:Platform="Any CPU"
```

Sign the resulting DLL:

```cmd
signtool sign /f cert.pfx /p <PFX-password> /t http://timestamp.digicert.com /fd SHA256 bin\Release\Outlook2021TodoAddIn.dll
```

### Certificate Trust (one-time per client)

See PowerShell snippet in the German section above.

### Project Structure

```
Outlook2021TodoAddIn/
├── ThisAddIn.cs                  # VSTO entry point
├── TaskPaneControl.xaml(.cs)     # WPF host for task pane
├── AppointmentsControl.cs        # Appointment logic (backup .bak required before edits)
├── TasksControl.cs               # Task logic
├── Properties/AssemblyInfo.cs    # Version, copyright
├── build.bat                     # Build + sign wrapper
└── README.md
```

### Conventions

- **Backup rule:** Create `.bak` copy before modifying `AppointmentsControl.cs`
- **Versioning:** SemVer in `AssemblyInfo.cs`, increment on every release
- **Headers:** Every `.cs` file carries a doc block (`@file`, `@brief`, `@author`, `@version`, `@date`, `@history` – newest first)

### Known Issues

| Issue | Cause | Fix |
|---|---|---|
| Add-in disabled on startup | Slow add-in detection | Registry: `HKCU\...\Resiliency\DoNotDisableAddinList` |
| Macro warning despite signature | CRL unreachable | `DontCheckPublisherCertRevocation = 1` |
| Empty task pane after Outlook update | Stale VSTO cache | Delete `%LOCALAPPDATA%\assembly\dl3`, restart Outlook |

### License

Internal tool, no public license. All rights reserved.

### Contact

**Gerhard Lustig** · gerhard@lustig.at · netz24.at# Outlook2021TodoAddIn

VSTO-Add-In für Outlook 2021, das Termine und Aufgaben in einem kombinierten Custom Task Pane darstellt.

---

## 🇩🇪 Deutsch

### Überblick

Das Add-In erweitert Outlook 2021 um ein Custom Task Pane mit konsolidierter Tages- und Wochenübersicht – Termine und Aufgaben in einem Bedienelement, ohne ständigen Wechsel zwischen Kalender- und Aufgaben-Modul.

### Features

- Kombinierte Anzeige von Terminen (Appointments) und Aufgaben (Tasks)
- Custom Task Pane, dockbar im Outlook-Hauptfenster
- Automatische Aktualisierung bei Änderungen am Kalender / an Tasks
- VBA-signiert via interner PKI (Root CA: `MyRootCA`, EKU Code Signing OID `1.3.6.1.5.5.7.3.3`)
- Lokale Konfiguration, kein externer Service erforderlich

### Voraussetzungen

| Komponente | Version |
|---|---|
| Outlook | 2021 (Desktop, MSI oder C2R) |
| .NET Framework | 4.8 |
| VSTO Runtime | 10.0 oder höher |
| Visual Studio | 2022 mit Workload *Office/SharePoint development* |
| Code-Signing-Zertifikat | Aus interner PKI (`sign_cert.sh -codesign`) |

### Installation

1. Release-ZIP entpacken nach `C:\@Tools\Outlook2021TodoAddIn\`
2. `setup.exe` ausführen (ClickOnce / VSTO-Installer)
3. Outlook neu starten
4. Add-In erscheint unter *Datei → Optionen → Add-Ins → COM-Add-Ins*

### Build

```cmd
cd D:\VisualStudio\Outlook2021TodoAddIn
msbuild Outlook2021TodoAddIn.sln /p:Configuration=Release /p:Platform="Any CPU"
```

Signierung der erzeugten DLL:

```cmd
signtool sign /f cert.pfx /p <PFX-Passwort> /t http://timestamp.digicert.com /fd SHA256 bin\Release\Outlook2021TodoAddIn.dll
```

### Zertifikats-Trust (einmalig pro Client)

```powershell
# Root CA in Trusted Root einbinden
Import-Certificate -FilePath "\\zerberus\share\MyRootCA.crt" -CertStoreLocation Cert:\LocalMachine\Root

# Code-Signing-Zertifikat als Trusted Publisher
Import-Certificate -FilePath "\\zerberus\share\codesign.cer" -CertStoreLocation Cert:\LocalMachine\TrustedPublisher

# CRL-Check für Outlook-Makros deaktivieren (Domänenumgebung ohne CRL-Distribution)
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security" `
                 -Name "DontCheckPublisherCertRevocation" -Value 1 -PropertyType DWORD -Force
```

### Projektstruktur

```
Outlook2021TodoAddIn/
├── ThisAddIn.cs                  # VSTO-Einstiegspunkt
├── TaskPaneControl.xaml(.cs)     # WPF-Host für Task Pane
├── AppointmentsControl.cs        # Termin-Logik (Backup-Pflicht: .bak vor Änderung)
├── TasksControl.cs               # Aufgaben-Logik
├── Properties/
│   └── AssemblyInfo.cs           # Version, Copyright
├── build.bat                     # Build- und Sign-Wrapper
└── README.md
```

### Wichtige Konventionen

- **Backup-Pflicht:** Vor jeder Änderung an `AppointmentsControl.cs` → `AppointmentsControl.cs.bak` anlegen
- **Versionierung:** SemVer in `AssemblyInfo.cs`, Hochzählen bei jedem Release
- **Header:** Jede `.cs`-Datei mit Doku-Block (`@file`, `@brief`, `@author`, `@version`, `@date`, `@history` – neueste zuerst)

### Bekannte Probleme

| Problem | Ursache | Lösung |
|---|---|---|
| Add-In wird beim Start deaktiviert | Slow Add-In Detection | Registry: `HKCU\...\Resiliency\DoNotDisableAddinList` setzen |
| Makro-Warnung trotz Signatur | CRL nicht erreichbar | `DontCheckPublisherCertRevocation = 1` (siehe oben) |
| Task Pane leer nach Outlook-Update | VSTO-Cache | `%LOCALAPPDATA%\assembly\dl3` löschen, Outlook neu starten |

### Lizenz

Internes Tool, keine öffentliche Lizenz. Alle Rechte vorbehalten.

### Kontakt

**Gerhard Lustig** · gerhard@lustig.at · netz24.at

---

## 🇬🇧 English

### Overview

VSTO add-in for Outlook 2021 providing a Custom Task Pane with a combined view of appointments and tasks – no more switching between Calendar and Tasks modules.

### Features

- Combined display of Appointments and Tasks
- Dockable Custom Task Pane inside the main Outlook window
- Auto-refresh on calendar/task changes
- Signed via internal PKI (Root CA `MyRootCA`, Code Signing EKU `1.3.6.1.5.5.7.3.3`)
- Fully local, no external service required

### Requirements

| Component | Version |
|---|---|
| Outlook | 2021 (Desktop, MSI or C2R) |
| .NET Framework | 4.8 |
| VSTO Runtime | 10.0+ |
| Visual Studio | 2022 with *Office/SharePoint development* workload |
| Code-Signing Certificate | From internal PKI (`sign_cert.sh -codesign`) |

### Installation

1. Extract release ZIP to `C:\@Tools\Outlook2021TodoAddIn\`
2. Run `setup.exe` (ClickOnce / VSTO installer)
3. Restart Outlook
4. Verify under *File → Options → Add-Ins → COM Add-Ins*

### Build

```cmd
cd D:\VisualStudio\Outlook2021TodoAddIn
msbuild Outlook2021TodoAddIn.sln /p:Configuration=Release /p:Platform="Any CPU"
```

Sign the resulting DLL:

```cmd
signtool sign /f cert.pfx /p <PFX-password> /t http://timestamp.digicert.com /fd SHA256 bin\Release\Outlook2021TodoAddIn.dll
```

### Certificate Trust (one-time per client)

See PowerShell snippet in the German section above.

### Project Structure

```
Outlook2021TodoAddIn/
├── ThisAddIn.cs                  # VSTO entry point
├── TaskPaneControl.xaml(.cs)     # WPF host for task pane
├── AppointmentsControl.cs        # Appointment logic (backup .bak required before edits)
├── TasksControl.cs               # Task logic
├── Properties/AssemblyInfo.cs    # Version, copyright
├── build.bat                     # Build + sign wrapper
└── README.md
```

### Conventions

- **Backup rule:** Create `.bak` copy before modifying `AppointmentsControl.cs`
- **Versioning:** SemVer in `AssemblyInfo.cs`, increment on every release
- **Headers:** Every `.cs` file carries a doc block (`@file`, `@brief`, `@author`, `@version`, `@date`, `@history` – newest first)

### Known Issues

| Issue | Cause | Fix |
|---|---|---|
| Add-in disabled on startup | Slow add-in detection | Registry: `HKCU\...\Resiliency\DoNotDisableAddinList` |
| Macro warning despite signature | CRL unreachable | `DontCheckPublisherCertRevocation = 1` |
| Empty task pane after Outlook update | Stale VSTO cache | Delete `%LOCALAPPDATA%\assembly\dl3`, restart Outlook |

### License

Internal tool, no public license. All rights reserved.

### Contact

**Gerhard Lustig** · gerhard@lustig.at · netz24.at
