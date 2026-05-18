# SSRSDeploymentTool

PowerShell WinForms-Tool für das Deployment von SQL Server Reporting Services (SSRS) Inhalten — entwickelt von [dtcSoftware](https://www.powershelldba.de) (Uwe Janke).

## Übersicht

`SSRSDeploymentTool` ist eine grafische PowerShell-Anwendung (WinForms) zum Deployen von SSRS-Reports, Datenquellen und Shared Datasets über die SSRS REST API v2.0. Bestehende Verbindungseinstellungen von Datenquellen bleiben beim Deployment erhalten.

**Version:** 3.0.1 | **Getestet auf:** SQL Server 2019 / 2022 / 2025, SSRS 16.x

## Features

- **WinForms GUI**: TreeView des SSRS-Serverordners links, Deployment-Konfiguration rechts
- **Reports (.rdl)**: werden immer überschrieben
- **Datenquellen (.rds/.rsds)**: bestehende Verbindungen bleiben erhalten, neue werden angelegt
- **Shared Datasets (.rsd)**: werden immer überschrieben
- **Authentifizierung**: Windows-Auth (automatisch) oder manuelle Credentials
- **Serverordner-Verwaltung**: TreeView mit Rechtsklick → Neuer Ordner
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

# Oder per Doppelklick auf Install.cmd (cross-domain Share)
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
├── SSRS_Deployment_Tool_v3.html           # Technische Dokumentation
└── SSRS_Deployment_Tool_v3_Dokumentation.docx
```

## Version

- **3.0.1** — Aktuelle Version
  - SSRS REST API v2.0
  - RDL df:-Namespace Fix
  - TLS 1.2 / selbstsignierte Zertifikate

## Mehr Informationen

- Website: [www.powershelldba.de](https://www.powershelldba.de)
- Entwickler: Uwe Janke, Senior IT-Spezialist / SQL Server DBA
