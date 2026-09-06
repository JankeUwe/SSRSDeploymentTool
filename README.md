# SSRSDeploymentTool

PowerShell WinForms-Tool für das Deployment von SQL Server Reporting Services (SSRS) Inhalten — entwickelt von [dtcSoftware](https://www.powershelldba.de) (Uwe Janke).

## Übersicht

`SSRSDeploymentTool` ist eine grafische PowerShell-Anwendung (WinForms) zum Deployen von SSRS-Reports, Datenquellen und Shared Datasets über die SSRS REST API v2.0. Bestehende Verbindungseinstellungen von Datenquellen bleiben beim Deployment erhalten. Seit v4.0.0 wird zusätzlich Power BI Report Server (PBIRS) sowie ein eigener Migration-Tab (Export/Import zwischen Servern) unterstützt.

**Version:** 4.0.0 | **Getestet auf:** SQL Server 2022, SSRS 16.x / PBIRS

## Features

- **WinForms GUI**: TreeView des SSRS-Serverordners links, Deployment-Konfiguration rechts
- **Reports (.rdl)**: werden immer überschrieben
- **Power BI Reports (.pbix)**: nur auf Power BI Report Server (PBIRS)
- **Datenquellen (.rds/.rsds)**: bestehende Verbindungen bleiben erhalten, neue werden angelegt
- **Shared Datasets (.rsd)**: werden immer überschrieben
- **Authentifizierung**: Windows-Auth (automatisch) oder manuelle Credentials
- **Serverordner-Verwaltung**: TreeView mit Rechtsklick → Neuer Ordner
- **Migration-Tab**: Export von einem Quell-SSRS in ein Zwischenverzeichnis, Import in ein Ziel-SSRS
- **Sprachauswahl**: Deutsch / Englisch (Strings/de.ps1, Strings/en.ps1)
- **Auto-Scan**: Ordner werden automatisch nach deploybaren Dateien durchsucht
- **RDL-Fixes**: Undeklarierten `df:`-Namespace-Präfix wird vor Upload automatisch ergänzt (keine Änderung an Originaldateien)
- **TLS 1.2**: Explizit aktiviert — funktioniert auch mit selbstsignierten / internen Zertifikaten

## Voraussetzungen

| Anforderung | Mindestversion |
|-------------|---------------|
| PowerShell | 5.1 |
| SQL Server Reporting Services | 2019 / 2022 / 2025 |
| SSRS REST API | v2.0 (SSRS 16.x) |

## Verwendung

```powershell
# Direkt starten
.\ReportDeplyment.ps1

# Oder per Doppelklick auf Start-SSRSDeployment.cmd (einheitlicher Starter)
```

## Deployment-Verhalten

| Dateityp | Verhalten |
|----------|-----------|
| `.rdl` — Reports | Immer überschrieben |
| `.rds` / `.rsds` — Datenquellen | Verbindung bleibt erhalten, neue werden angelegt |
| `.rsd` — Shared Datasets | Immer überschrieben |

## Projektstruktur

```
SSRSDeploymentTool/
├── ReportDeplyment.ps1                    # Hauptskript (WinForms GUI + REST API Logik)
├── Start-SSRSDeployment.cmd               # Einheitlicher Starter
├── Strings/                               # de.ps1 / en.ps1 (Sprachauswahl)
├── Docs/                                  # Präsentation und technische Doku
└── CHANGELOG.md
```

## Version

Siehe [CHANGELOG.md](CHANGELOG.md) für die vollständige Historie.

- **4.0.0** — Aktuelle Version
  - PBIRS-Support (.pbix)
  - Migration-Tab (Export/Import zwischen Servern)
  - Sprachsystem (DE/EN) & Auto-Scan
  - SSRS REST API v2.0
  - RDL df:-Namespace Fix
  - TLS 1.2 / selbstsignierte Zertifikate

## Mehr Informationen

- Projektseite: [powershelldba.de/ssrsdeploymenttool](https://www.powershelldba.de/ssrsdeploymenttool/)
- Website: [www.powershelldba.de](https://www.powershelldba.de)
- Entwickler: Uwe Janke, Senior IT-Spezialist / SQL Server DBA
