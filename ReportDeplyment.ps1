#Requires -Version 5.1
<#
.SYNOPSIS
    SSRS Report Deployment Tool v3
    WinForms-basiertes Deployment-Tool fuer SQL Server Reporting Services (2019/2022/2025).

.DESCRIPTION
    - Reports (.rdl)           : werden immer ueberschrieben
    - Datenquellen (.rds/.rsds): bestehende Connections bleiben erhalten, neue werden angelegt
    - Shared Datasets (.rsd)   : werden immer ueberschrieben
    - Authentifizierung        : Windows-Auth oder manuelle Credentials
    - Serverordner             : TreeView links, Rechtsklick Neuer Ordner

.NOTES
    Version : 3.0.1
    API     : SSRS REST API v2.0 (SQL Server 2022 / SSRS 16.x)

    RDL-Fixes (automatisch, kein Eingriff in Originaldateien):
    - Undeklarierten df:-Namespace-Praefix wird vor Upload ergaenzt

    Layout-Regeln (nicht aendern):
    - Keine AutoSize auf GroupBox oder TLP mit Dock=Top
    - Traeger-Panel mit fester Hoehe traegt die Konfig-GroupBox
    - SplitterDistance wird ausschliesslich in Form_Shown nach DoEvents gesetzt
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------------------
# SSL-Zertifikatsvalidierung – akzeptiert selbstsignierte / interne Zertifikate.
# Setzt ServerCertificateValidationCallback (funktioniert in PS 5.1 auf allen
# .NET-Versionen, da ICertificatePolicy in neueren .NET-Laufzeiten entfernt wurde).
# ---------------------------------------------------------------------------
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {
    param($sender, $certificate, $chain, $sslPolicyErrors)
    return $true
}
# TLS 1.1 und 1.2 aktivieren (manche SSRS-Instanzen benoetigen explizit TLS 1.2)
[System.Net.ServicePointManager]::SecurityProtocol =
    [System.Net.SecurityProtocolType]::Tls12 -bor
    [System.Net.SecurityProtocolType]::Tls11 -bor
    [System.Net.SecurityProtocolType]::Tls

# =============================================================================
# REST-Hilfsfunktionen
# =============================================================================

function Get-SSRSApiBase {
    param([string]$ServerUrl)
    $url = $ServerUrl.TrimEnd('/')
    # Benutzer kann sowohl /ReportServer als auch /Reports eingeben.
    # API-Endpunkt ist immer unterhalb von /Reports/api/v2.0
    if ($url -match '/api/v2\.0$') { return $url }
    if ($url -match '/ReportServer$') { $url = $url -replace '/ReportServer$', '/Reports' }
    return "$url/api/v2.0"
}

function Invoke-SSRSRequest {
    param(
        [string]$Uri,
        [string]$Method      = 'GET',
        [object]$Body        = $null,
        [string]$ContentType = 'application/json',
        [System.Management.Automation.PSCredential]$Credential = $null
    )
    $p = @{
        Uri             = $Uri
        Method          = $Method
        Headers         = @{ 'Content-Type' = 'application/json' }
        UseBasicParsing = $true
        ErrorAction     = 'Stop'
    }
    if ($Body) {
        $jsonBody = $Body | ConvertTo-Json -Depth 10 -Compress
        $p['Body']        = $jsonBody
        $p['ContentType'] = $ContentType
    }
    if ($Credential) { $p['Credential'] = $Credential } else { $p['UseDefaultCredentials'] = $true }
    try {
        return Invoke-RestMethod @p
    } catch {
        $detail = ''
        try {
            $resp = $_.Exception.Response
            if ($resp) {
                $stream = $resp.GetResponseStream()
                $stream.Position = 0
                $reader = New-Object System.IO.StreamReader($stream)
                $detail = $reader.ReadToEnd()
                $reader.Dispose()
                $stream.Dispose()
            }
        } catch { $detail = "(Response-Body nicht lesbar: $($_.Exception.Message))" }

        # JSON-Body des Requests fuer Diagnose
        $sentBody = if ($Body) { $jsonBody } else { '(kein Body)' }

        $msg = "$($_.Exception.Message) | URI: $Uri | Body: $sentBody | Server-Antwort: $detail"
        throw $msg
    }
}

function Test-SSRSConnection {
    param([string]$ServerUrl, [System.Management.Automation.PSCredential]$Credential = $null)
    return Invoke-SSRSRequest -Uri "$(Get-SSRSApiBase $ServerUrl)/System" -Credential $Credential
}

function Get-SSRSFolders {
    param([string]$ApiBase, [System.Management.Automation.PSCredential]$Credential = $null)
    return (Invoke-SSRSRequest -Uri "$ApiBase/Folders?`$orderby=Path" -Credential $Credential).value
}

function Get-SSRSFolderExists {
    param([string]$ApiBase, [string]$FolderPath, [System.Management.Automation.PSCredential]$Credential = $null)
    try {
        $p = @{
            Uri             = "$ApiBase/Folders($([Uri]::EscapeDataString($FolderPath)))"
            Method          = 'GET'
            UseBasicParsing = $true
            ErrorAction     = 'Stop'
        }
        if ($Credential) { $p['Credential'] = $Credential } else { $p['UseDefaultCredentials'] = $true }
        Invoke-RestMethod @p | Out-Null
        return $true
    } catch { return $false }
}

function New-SSRSFolderSingle {
    param([string]$ApiBase, [string]$FolderPath, [System.Management.Automation.PSCredential]$Credential = $null)
    $parts = $FolderPath.TrimStart('/') -split '/'
    $name  = $parts[-1]
    $body  = @{ Name = $name; Path = $FolderPath }
    try {
        Invoke-SSRSRequest -Uri "$ApiBase/Folders" -Method 'POST' -Body $body -Credential $Credential | Out-Null
    } catch {
        # 409 = Ordner existiert bereits – kein Fehler
        if ($_ -notmatch '409') { throw }
    }
}

function New-SSRSFolderRecursive {
    param([string]$ApiBase, [string]$FolderPath, [System.Management.Automation.PSCredential]$Credential = $null)
    $cur = ''
    foreach ($part in ($FolderPath.TrimStart('/') -split '/')) {
        if ([string]::IsNullOrWhiteSpace($part)) { continue }
        $cur = "$cur/$part"
        if (-not (Get-SSRSFolderExists -ApiBase $ApiBase -FolderPath $cur -Credential $Credential)) {
            New-SSRSFolderSingle -ApiBase $ApiBase -FolderPath $cur -Credential $Credential
        }
    }
}

function Get-SSRSItemExists {
    param([string]$ApiBase, [string]$ItemPath, [System.Management.Automation.PSCredential]$Credential = $null)
    try {
        # Zuverlässiger als OData-Key: $filter auf Path-Feld
        $filter = "Path eq '$($ItemPath.Replace("'","''"))'"
        $p = @{
            Uri             = "$ApiBase/CatalogItems?`$filter=$([Uri]::EscapeDataString($filter))"
            Method          = 'GET'
            UseBasicParsing = $true
            ErrorAction     = 'Stop'
        }
        if ($Credential) { $p['Credential'] = $Credential } else { $p['UseDefaultCredentials'] = $true }
        $result = Invoke-RestMethod @p
        if ($result.value -and $result.value.Count -gt 0) { return $result.value[0] }
        return $null
    } catch { return $null }
}

# =============================================================================
# Deploy-Funktionen
# =============================================================================

function Repair-RDLContent {
    <#
    .SYNOPSIS
        Bereinigt bekannte RDL-Probleme im Speicher vor dem Deployment.
        Die Originaldatei auf dem Dateisystem wird NICHT veraendert.

    Bekannte Fixes:
        1. Undeklarierten df:-Namespace-Praefix im Report-Root-Element deklarieren
           (SSRS 2022 RTM: "'df' is an undeclared prefix. Line 5, position 4.")
    #>
    param([byte[]]$RawBytes)

    try {
        $text = [System.Text.Encoding]::UTF8.GetString($RawBytes).TrimStart([char]0xFEFF)

        # Fix: df:-Praefix im XML-Text vorhanden aber xmlns:df fehlt im Root-Element
        if ($text -match 'df:' -and $text -notmatch 'xmlns:df') {
            # Namespace-Deklaration direkt im oeffnenden <Report ...>-Tag ergaenzen
            $dfNs = 'xmlns:df="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition/defaultfontfamily"'
            $text = $text -replace '(<Report\b[^>]*)(>)', "`$1 $dfNs`$2"
        }

        # UTF-8 ohne BOM zurueckgeben
        return [System.Text.Encoding]::UTF8.GetBytes($text)
    } catch {
        # Bei Fehler unveraenderte Bytes zurueckgeben
        return $RawBytes
    }
}

function Deploy-Report {
    param([string]$ApiBase, [string]$FilePath, [string]$TargetFolder, [System.Management.Automation.PSCredential]$Credential = $null)
    $name  = [System.IO.Path]::GetFileNameWithoutExtension($FilePath).Trim()
    $ipath = "$TargetFolder/$name"
    # RDL einlesen und im Speicher bereinigen (Original bleibt unveraendert)
    $rawBytes = [System.IO.File]::ReadAllBytes($FilePath)
    $bytes    = Repair-RDLContent -RawBytes $rawBytes
    if ($bytes.Length -ne $rawBytes.Length) {
        Write-Log "  [RDL-Fix] Namespace-Deklaration automatisch ergaenzt (df:)" -Lv 'Warning'
    }
    $content  = [Convert]::ToBase64String($bytes)

    $existing = Get-SSRSItemExists -ApiBase $ApiBase -ItemPath $ipath -Credential $Credential
    if ($existing) {
        Invoke-SSRSRequest -Uri "$ApiBase/CatalogItems($($existing.Id))" -Method 'DELETE' -Credential $Credential | Out-Null
    }
    Invoke-SSRSRequest -Uri "$ApiBase/Reports" -Method 'POST' -Credential $Credential `
        -Body @{ Name=$name; Path=$ipath; Content=$content } | Out-Null
    if ($existing) { return 'UPDATED' } else { return 'CREATED' }
}

function Deploy-DataSource {
    param([string]$ApiBase, [string]$FilePath, [string]$TargetFolder, [System.Management.Automation.PSCredential]$Credential = $null)
    $name  = [System.IO.Path]::GetFileNameWithoutExtension($FilePath).Trim()
    $ipath = "$TargetFolder/$name"
    if (Get-SSRSItemExists -ApiBase $ApiBase -ItemPath $ipath -Credential $Credential) {
        return 'SKIPPED (bereits vorhanden - Connection bleibt erhalten)'
    }
    [xml]$x = Get-Content -Path $FilePath -Encoding UTF8
    $cp     = $x.RptDataSource.ConnectionProperties

    # CredentialRetrieval fuer die REST API bestimmen.
    # Auswertungsreihenfolge:
    #   1. <IntegratedSecurity>true</IntegratedSecurity>  → "integrated"
    #   2. <CredentialRetrieval>...</CredentialRetrieval> → Wert lowercase normalisieren
    #      (API akzeptiert: integrated | prompt | store | none)
    #   3. Fallback                                       → "integrated"
    $credRaw = if ($cp.IntegratedSecurity -eq 'true') {
                   'integrated'
               } elseif (-not [string]::IsNullOrWhiteSpace($cp.CredentialRetrieval)) {
                   switch ($cp.CredentialRetrieval.Trim().ToLower()) {
                       'integrated' { 'integrated' }
                       'prompt'     { 'prompt'     }
                       'store'      { 'store'      }
                       'none'       { 'none'       }
                       default      { 'integrated' }
                   }
               } else {
                   'integrated'
               }

    Invoke-SSRSRequest -Uri "$ApiBase/DataSources" -Method 'POST' -Credential $Credential -Body @{
        Name                = $name
        Path                = $ipath
        DataSourceType      = if ($cp.Extension)      { $cp.Extension }      else { 'SQL' }
        ConnectionString    = if ($cp.ConnectString)  { $cp.ConnectString }  else { '' }
        CredentialRetrieval = $credRaw
    } | Out-Null
    return 'CREATED'
}

function Deploy-SharedDataset {
    param([string]$ApiBase, [string]$FilePath, [string]$TargetFolder, [System.Management.Automation.PSCredential]$Credential = $null)
    $name    = [System.IO.Path]::GetFileNameWithoutExtension($FilePath).Trim()
    $ipath   = "$TargetFolder/$name"
    $content = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($FilePath))

    $existing = Get-SSRSItemExists -ApiBase $ApiBase -ItemPath $ipath -Credential $Credential
    if ($existing) {
        # Bestehenden Dataset loeschen – vermeidet alle PATCH/409-Probleme
        Invoke-SSRSRequest -Uri "$ApiBase/CatalogItems($($existing.Id))" -Method 'DELETE' -Credential $Credential | Out-Null
    }
    Invoke-SSRSRequest -Uri "$ApiBase/DataSets" -Method 'POST' -Credential $Credential `
        -Body @{ Name=$name; Path=$ipath; Content=$content } | Out-Null
    if ($existing) { return 'UPDATED' } else { return 'CREATED' }
}

# =============================================================================
# Lokalen SSRS-Server erkennen
# =============================================================================

function Find-LocalSSRS {
    <#
    .SYNOPSIS
        Prueft ob auf dem lokalen Rechner ein SSRS laeuft.
        Probiert gaengige URL-Varianten und gibt die erste funktionierende zurueck.
        Gibt $null zurueck wenn kein Server gefunden.
    #>

    $hostname   = $env:COMPUTERNAME
    # HTTPS-Varianten werden vor HTTP geprueft; bei selbstsignierten Zertifikaten
    # greift der global gesetzte ServerCertificateValidationCallback (s.o.).
    $candidates = @(
        "https://$hostname/Reports"
        "https://$hostname/ReportServer"
        "http://$hostname/Reports"
        "http://$hostname/ReportServer"
        "https://localhost/Reports"
        "https://localhost/ReportServer"
        "http://localhost/Reports"
        "http://localhost/ReportServer"
    )

    foreach ($url in $candidates) {
        try {
            $api = Get-SSRSApiBase -ServerUrl $url
            $p   = @{
                Uri                   = "$api/System"
                Method                = 'GET'
                UseBasicParsing       = $true
                UseDefaultCredentials = $true
                ErrorAction           = 'Stop'
                TimeoutSec            = 4
            }
            $info = Invoke-RestMethod @p
            # Erreichbar – URL zurueckgeben (normalisiert auf /Reports Basis)
            return (Get-SSRSApiBase -ServerUrl $url) -replace '/api/v2\.0$', ''
        } catch {
            # Nicht erreichbar oder Fehler – naechste Variante
        }
    }
    return $null
}

# =============================================================================
# GUI
# =============================================================================

function Show-DeploymentTool {

    # -------------------------------------------------------------------------
    # Design
    # -------------------------------------------------------------------------
    $cBg     = [System.Drawing.Color]::FromArgb(240, 242, 245)
    $cPanel  = [System.Drawing.Color]::White
    $cHeader = [System.Drawing.Color]::FromArgb(0,   70, 127)
    $cAccent = [System.Drawing.Color]::FromArgb(0,  114, 198)
    $cOk     = [System.Drawing.Color]::FromArgb(0,  140,  70)
    $cWarn   = [System.Drawing.Color]::FromArgb(190, 100,   0)
    $cErr    = [System.Drawing.Color]::FromArgb(180,  30,  30)
    $cSkip   = [System.Drawing.Color]::FromArgb(110, 110, 160)
    $cLogBg  = [System.Drawing.Color]::FromArgb(18,   20,  28)
    $cLogFg  = [System.Drawing.Color]::FromArgb(200, 212, 224)

    $fDef   = New-Object System.Drawing.Font('Segoe UI',  9)
    $fBold  = New-Object System.Drawing.Font('Segoe UI',  9, [System.Drawing.FontStyle]::Bold)
    $fSmall = New-Object System.Drawing.Font('Segoe UI',  8)
    $fMono  = New-Object System.Drawing.Font('Consolas',  8)
    $fTitle = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
    $fHead  = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)

    # -------------------------------------------------------------------------
    # Layout-Konstanten (einmal definieren, ueberall verwenden)
    # -------------------------------------------------------------------------
    $headerH    = 35    # blauer Header
    $cfgRowH    = 34    # Hoehe einer Konfigzeile
    $cfgRows    = 5     # Anzahl Zeilen
    $cfgGrpTop  = 22    # GroupBox-Titelhoehe (geschaetzt)
    $cfgPadTB   = 10    # Padding top + bottom je 10 = 20 innen, plus 4 aussen = 24
    $cfgH       = $cfgRows * $cfgRowH + $cfgGrpTop + $cfgPadTB * 2 + 4
    # = 5*34 + 22 + 20 + 4 = 216
    $rootPadTop = 10    # pnlRoot Padding oben
    $sepH       = 8     # Trennstreifen
    # Gesamthoehe Konfigbereich im pnlRoot = cfgH + sepH
    $cfgAreaH   = $cfgH + $sepH

    # -------------------------------------------------------------------------
    # Hauptfenster
    # -------------------------------------------------------------------------
    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'SSRS Deployment Tool  v3.0'
    $form.Size            = New-Object System.Drawing.Size(1080, 840)
    $form.MinimumSize     = New-Object System.Drawing.Size(900, 700)
    $form.StartPosition   = 'CenterScreen'
    $form.BackColor       = $cBg
    $form.Font            = $fDef
    $form.FormBorderStyle = 'Sizable'

    # -------------------------------------------------------------------------
    # Header-Panel  (Dock=Top, feste Hoehe $headerH)
    # -------------------------------------------------------------------------
    $pnlHeader           = New-Object System.Windows.Forms.Panel
    $pnlHeader.Dock      = 'Top'
    $pnlHeader.Height    = $headerH
    $pnlHeader.BackColor = $cHeader
    $pnlHeader.Padding   = New-Object System.Windows.Forms.Padding(14, 0, 14, 0)

    $lblTitle            = New-Object System.Windows.Forms.Label
    $lblTitle.Text       = 'dtcSoftware SSRS Deployment Tool'
    $lblTitle.ForeColor  = [System.Drawing.Color]::White
    $lblTitle.Font       = $fTitle
    $lblTitle.Dock       = 'Fill'
    $lblTitle.TextAlign  = 'MiddleLeft'

    $lblVer              = New-Object System.Windows.Forms.Label
    $lblVer.Text         = '2025-26 v3.0  |  REST API v2.0'
    $lblVer.ForeColor    = [System.Drawing.Color]::FromArgb(170, 205, 240)
    $lblVer.Font         = $fSmall
    $lblVer.Dock         = 'Right'
    $lblVer.Width        = 160
    $lblVer.TextAlign    = 'MiddleRight'

    $pnlHeader.Controls.AddRange(@($lblTitle, $lblVer))
    $form.Controls.Add($pnlHeader)

    # -------------------------------------------------------------------------
    # Root-Panel  (Dock=Fill, traegt den Rest)
    # Padding: links/rechts 10, oben $rootPadTop, unten 6
    # -------------------------------------------------------------------------
    $pnlRoot           = New-Object System.Windows.Forms.Panel
    $pnlRoot.Dock      = 'Fill'
    $pnlRoot.BackColor = $cBg
    $pnlRoot.Padding   = New-Object System.Windows.Forms.Padding(10, $rootPadTop, 10, 6)
    $form.Controls.Add($pnlRoot)

    # -------------------------------------------------------------------------
    # Traeger-Panel fuer Konfiguration  (Dock=Top, Hoehe = $cfgH)
    # Durch dieses extra Panel bekommt die GroupBox eine stabile Hoehe
    # ohne dass AutoSize oder Dock=Fill das Layout zerstoert.
    # -------------------------------------------------------------------------
    $pnlCfgWrap           = New-Object System.Windows.Forms.Panel
    $pnlCfgWrap.Dock      = 'Top'
    $pnlCfgWrap.Height    = $cfgH
    $pnlCfgWrap.BackColor = $cBg

    $grpCfg              = New-Object System.Windows.Forms.GroupBox
    $grpCfg.Text         = ' Konfiguration'
    $grpCfg.Font         = $fHead
    $grpCfg.ForeColor    = $cHeader
    $grpCfg.BackColor    = $cPanel
    $grpCfg.Dock         = 'Fill'   # fuellt den Traeger-Panel vollstaendig
    $grpCfg.Padding      = New-Object System.Windows.Forms.Padding(10, $cfgPadTB, 10, $cfgPadTB)

    # Internes TableLayoutPanel – Dock=Fill, feste Zeilenhoehen
    $tlpCfg              = New-Object System.Windows.Forms.TableLayoutPanel
    $tlpCfg.Dock         = 'Fill'
    $tlpCfg.ColumnCount  = 3
    $tlpCfg.RowCount     = $cfgRows
    $tlpCfg.BackColor    = $cPanel

    [void]$tlpCfg.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 175)))
    [void]$tlpCfg.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$tlpCfg.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 116)))
    for ($i = 0; $i -lt $cfgRows; $i++) {
        [void]$tlpCfg.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $cfgRowH)))
    }

    # Helfer
    function New-CfgLabel  { param([string]$t) $l=New-Object System.Windows.Forms.Label; $l.Text=$t; $l.Font=$fBold; $l.TextAlign='MiddleRight'; $l.Dock='Fill'; $l.AutoSize=$false; $l.AutoEllipsis=$true; $l }
    function New-CfgText   { param([string]$p='') $t=New-Object System.Windows.Forms.TextBox; $t.Dock='Fill'; $t.Font=$fDef; if($p){$t.Text=$p}; $t }
    function New-CfgButton { param([string]$t) $b=New-Object System.Windows.Forms.Button; $b.Text=$t; $b.Dock='Fill'; $b.Font=$fDef; $b.BackColor=$cAccent; $b.ForeColor=[System.Drawing.Color]::White; $b.FlatStyle='Flat'; $b.FlatAppearance.BorderSize=0; $b.Cursor='Hand'; $b }

    # Zeile 0 – Server URL
    $lblSrv  = New-CfgLabel  'Report Server URL:'
    $txtSrv  = New-CfgText   ''
    $txtSrv.ForeColor = [System.Drawing.Color]::FromArgb(130,130,130)
    $txtSrv.Text      = 'Wird gesucht ...'
    $btnConn = New-CfgButton 'Verbinden'

    # Zeile 1 – Quellverzeichnis
    $lblSrc  = New-CfgLabel  'Quellverzeichnis:'
    $txtSrc  = New-CfgText   ''
    $btnBrw  = New-CfgButton 'Durchsuchen'

    # Zeile 2 – Zielordner (readonly, per Tree gesetzt)
    $lblTgt  = New-CfgLabel  'Zielordner (Server):'
    $txtTgt  = New-CfgText   ''
    $txtTgt.BackColor = [System.Drawing.Color]::FromArgb(238, 244, 252)

    $pnlHint           = New-Object System.Windows.Forms.Panel
    $pnlHint.Dock      = 'Fill'
    $lblHint           = New-Object System.Windows.Forms.Label
    $lblHint.Text      = 'Im Baum links auswaehlen'
    $lblHint.ForeColor = $cSkip
    $lblHint.Font      = $fSmall
    $lblHint.Dock      = 'Fill'
    $lblHint.TextAlign = 'MiddleLeft'
    $pnlHint.Controls.Add($lblHint)

    # Zeile 3 – Credentials Checkbox
    $lblCrd  = New-CfgLabel 'Authentifizierung:'
    $chkCrd  = New-Object System.Windows.Forms.CheckBox
    $chkCrd.Text = 'Manuelle Credentials verwenden'; $chkCrd.Font = $fDef; $chkCrd.Dock = 'Fill'

    # Zeile 4 – User / Passwort
    $pnlCrd        = New-Object System.Windows.Forms.Panel
    $pnlCrd.Dock   = 'Fill'
    $pnlCrd.Enabled = $false

    $tlpCr = New-Object System.Windows.Forms.TableLayoutPanel
    $tlpCr.Dock = 'Fill'; $tlpCr.ColumnCount = 4; $tlpCr.RowCount = 1; $tlpCr.BackColor = $cPanel
    [void]$tlpCr.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,  62)))
    [void]$tlpCr.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,  50)))
    [void]$tlpCr.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,  80)))
    [void]$tlpCr.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,  50)))

    $lU = New-Object System.Windows.Forms.Label; $lU.Text='User:'; $lU.Font=$fBold; $lU.Dock='Fill'; $lU.TextAlign='MiddleRight'
    $tU = New-Object System.Windows.Forms.TextBox; $tU.Dock='Fill'; $tU.Font=$fDef; $tU.Text="$env:USERDOMAIN\$env:USERNAME"
    $lP = New-Object System.Windows.Forms.Label; $lP.Text='Kennwort:'; $lP.Font=$fBold; $lP.Dock='Fill'; $lP.TextAlign='MiddleRight'
    $tP = New-Object System.Windows.Forms.TextBox; $tP.Dock='Fill'; $tP.Font=$fDef; $tP.PasswordChar=[char]0x2022

    $tlpCr.Controls.Add($lU,0,0); $tlpCr.Controls.Add($tU,1,0)
    $tlpCr.Controls.Add($lP,2,0); $tlpCr.Controls.Add($tP,3,0)
    $pnlCrd.Controls.Add($tlpCr)
    $chkCrd.Add_CheckedChanged({ $pnlCrd.Enabled = $chkCrd.Checked })

    # TLP fuellen
    $tlpCfg.Controls.Add($lblSrv, 0,0); $tlpCfg.Controls.Add($txtSrv, 1,0); $tlpCfg.Controls.Add($btnConn,2,0)
    $tlpCfg.Controls.Add($lblSrc, 0,1); $tlpCfg.Controls.Add($txtSrc, 1,1); $tlpCfg.Controls.Add($btnBrw, 2,1)
    $tlpCfg.Controls.Add($lblTgt, 0,2); $tlpCfg.Controls.Add($txtTgt, 1,2); $tlpCfg.Controls.Add($pnlHint,2,2)
    $tlpCfg.Controls.Add($lblCrd, 0,3); $tlpCfg.Controls.Add($chkCrd, 1,3); $tlpCfg.Controls.Add((New-Object System.Windows.Forms.Panel),2,3)
    $tlpCfg.Controls.Add((New-Object System.Windows.Forms.Panel),0,4); $tlpCfg.Controls.Add($pnlCrd,1,4)

    $grpCfg.Controls.Add($tlpCfg)
    $pnlCfgWrap.Controls.Add($grpCfg)

    # =========================================================================
    # Hauptbereich – horizontaler SplitContainer (Tree | rechts)
    # WICHTIG: Dock=Fill Control ZUERST zu pnlRoot hinzufuegen,
    #          danach erst die Dock=Top Controls – sonst wird Fill verdraengt.
    # SplitterDistance wird AUSSCHLIESSLICH in Form_Shown gesetzt.
    # =========================================================================
    $splitH               = New-Object System.Windows.Forms.SplitContainer
    $splitH.Dock          = 'Fill'
    $splitH.Orientation   = 'Vertical'
    $splitH.BorderStyle   = 'None'
    $splitH.BackColor     = $cBg
    $splitH.Panel1MinSize = 160
    $splitH.Panel2MinSize = 16
    $pnlRoot.Controls.Add($splitH)   # <-- Fill zuerst

    # Trennstreifen und Konfig-Wrapper danach (Dock=Top)
    $sep           = New-Object System.Windows.Forms.Panel
    $sep.Dock      = 'Top'
    $sep.Height    = $sepH
    $sep.BackColor = $cBg
    $pnlRoot.Controls.Add($sep)
    $pnlRoot.Controls.Add($pnlCfgWrap)

    # =========================================================================
    # LINKS – TreeView
    # =========================================================================
    $grpTree           = New-Object System.Windows.Forms.GroupBox
    $grpTree.Text      = ' Serverordner'
    $grpTree.Font      = $fHead
    $grpTree.ForeColor = $cHeader
    $grpTree.BackColor = $cPanel
    $grpTree.Dock      = 'Fill'
    $grpTree.Padding   = New-Object System.Windows.Forms.Padding(6, 4, 6, 4)

    $tv               = New-Object System.Windows.Forms.TreeView
    $tv.Dock          = 'Fill'
    $tv.Font          = $fDef
    $tv.BackColor     = $cPanel
    $tv.BorderStyle   = 'FixedSingle'
    $tv.HideSelection = $false
    $tv.ShowRootLines = $true
    $tv.ShowPlusMinus = $true
    $tv.FullRowSelect = $true

    # Ordner-Icons per GDI
    $imgList            = New-Object System.Windows.Forms.ImageList
    $imgList.ImageSize  = New-Object System.Drawing.Size(16,16)
    $imgList.ColorDepth = 'Depth32Bit'

    $bmpClosed = New-Object System.Drawing.Bitmap(16,16)
    $gC = [System.Drawing.Graphics]::FromImage($bmpClosed)
    $brC = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(220,170,20))
    $gC.FillRectangle($brC,0,5,15,10); $gC.FillRectangle($brC,0,3,7,4); $gC.Dispose(); $brC.Dispose()

    $bmpOpen = New-Object System.Drawing.Bitmap(16,16)
    $gO = [System.Drawing.Graphics]::FromImage($bmpOpen)
    $brO = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255,200,50))
    $gO.FillRectangle($brO,0,5,15,10); $gO.FillRectangle($brO,0,3,7,4); $gO.Dispose(); $brO.Dispose()

    $imgList.Images.Add($bmpClosed)  # Index 0 = geschlossen
    $imgList.Images.Add($bmpOpen)    # Index 1 = offen
    $tv.ImageList          = $imgList
    $tv.ImageIndex         = 0
    $tv.SelectedImageIndex = 1

    $lblTreeStat           = New-Object System.Windows.Forms.Label
    $lblTreeStat.Dock      = 'Bottom'
    $lblTreeStat.Height    = 22
    $lblTreeStat.Font      = $fSmall
    $lblTreeStat.ForeColor = $cSkip
    $lblTreeStat.Text      = 'Zuerst verbinden ...'
    $lblTreeStat.TextAlign = 'MiddleLeft'
    $lblTreeStat.BackColor = $cPanel

    $ctxTree             = New-Object System.Windows.Forms.ContextMenuStrip
    $mnuNew              = New-Object System.Windows.Forms.ToolStripMenuItem('Neuen Ordner anlegen')
    $mnuNew.Font         = $fDef
    $mnuRefresh          = New-Object System.Windows.Forms.ToolStripMenuItem('Baumstruktur aktualisieren')
    $mnuRefresh.Font     = $fDef
    [void]$ctxTree.Items.Add($mnuNew)
    [void]$ctxTree.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    [void]$ctxTree.Items.Add($mnuRefresh)
    $tv.ContextMenuStrip = $ctxTree

    $grpTree.Controls.AddRange(@($tv, $lblTreeStat))
    $splitH.Panel1.Controls.Add($grpTree)

    # =========================================================================
    # RECHTS – Deploy-Leiste zuerst (Dock=Bottom), dann splitV (Dock=Fill)
    # Reihenfolge ist zwingend: Bottom/Top vor Fill hinzufuegen
    # =========================================================================
    $pnlDep           = New-Object System.Windows.Forms.Panel
    $pnlDep.Dock      = 'Bottom'
    $pnlDep.Height    = 46
    $pnlDep.BackColor = $cBg
    $pnlDep.Padding   = New-Object System.Windows.Forms.Padding(0,5,0,0)

    $btnDeploy           = New-Object System.Windows.Forms.Button
    $btnDeploy.Text      = '▶  Deployment starten'
    $btnDeploy.Font      = New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
    $btnDeploy.Dock      = 'Right'
    $btnDeploy.Width     = 215
    $btnDeploy.BackColor = $cHeader
    $btnDeploy.ForeColor = [System.Drawing.Color]::White
    $btnDeploy.FlatStyle = 'Flat'
    $btnDeploy.FlatAppearance.BorderSize = 0
    $btnDeploy.Cursor    = 'Hand'
    $btnDeploy.Enabled   = $false

    $pbProg         = New-Object System.Windows.Forms.ProgressBar
    $pbProg.Dock    = 'Fill'
    $pbProg.Minimum = 0
    $pbProg.Value   = 0
    $pbProg.Style   = 'Continuous'

    $pnlDep.Controls.AddRange(@($btnDeploy, $pbProg))
    $splitH.Panel2.Controls.Add($pnlDep)   # Bottom zuerst

    # splitV (Fill) nach pnlDep (Bottom) hinzufuegen
    $splitV               = New-Object System.Windows.Forms.SplitContainer
    $splitV.Dock          = 'Fill'
    $splitV.Orientation   = 'Horizontal'
    $splitV.BorderStyle   = 'None'
    $splitV.BackColor     = $cBg
    $splitV.Panel1MinSize = 80
    $splitV.Panel2MinSize = 80
    $splitH.Panel2.Controls.Add($splitV)   # Fill nach Bottom
    $grpFiles           = New-Object System.Windows.Forms.GroupBox
    $grpFiles.Text      = ' Gefundene Dateien'
    $grpFiles.Font      = $fHead
    $grpFiles.ForeColor = $cHeader
    $grpFiles.BackColor = $cPanel
    $grpFiles.Dock      = 'Fill'
    $grpFiles.Padding   = New-Object System.Windows.Forms.Padding(6)

    $clv               = New-Object System.Windows.Forms.ListView
    $clv.Dock          = 'Fill'; $clv.View='Details'; $clv.CheckBoxes=$true
    $clv.FullRowSelect = $true;  $clv.GridLines=$true; $clv.Font=$fMono; $clv.BackColor=$cPanel
    [void]$clv.Columns.Add('Datei',     220)
    [void]$clv.Columns.Add('Typ',        64)
    [void]$clv.Columns.Add('Groesse',    72)
    [void]$clv.Columns.Add('Geaendert', 148)
    [void]$clv.Columns.Add('Status',    215)

    $pnlFBar        = New-Object System.Windows.Forms.Panel
    $pnlFBar.Dock   = 'Bottom'; $pnlFBar.Height=32; $pnlFBar.BackColor=$cPanel

    $btnScan           = New-Object System.Windows.Forms.Button
    $btnScan.Text      = 'Scannen'; $btnScan.Width=90; $btnScan.Dock='Left'; $btnScan.Font=$fDef
    $btnScan.BackColor = $cAccent; $btnScan.ForeColor=[System.Drawing.Color]::White
    $btnScan.FlatStyle = 'Flat'; $btnScan.FlatAppearance.BorderSize=0; $btnScan.Cursor='Hand'

    $btnAll  = New-Object System.Windows.Forms.Button; $btnAll.Text='Alle';  $btnAll.Width=52;  $btnAll.Dock='Left'; $btnAll.Font=$fSmall;  $btnAll.FlatStyle='Flat'
    $btnNone = New-Object System.Windows.Forms.Button; $btnNone.Text='Keine'; $btnNone.Width=52; $btnNone.Dock='Left'; $btnNone.Font=$fSmall; $btnNone.FlatStyle='Flat'

    $lblFCnt           = New-Object System.Windows.Forms.Label
    $lblFCnt.Text      = 'Noch kein Scan'; $lblFCnt.Dock='Fill'
    $lblFCnt.TextAlign = 'MiddleRight'; $lblFCnt.Font=$fSmall; $lblFCnt.ForeColor=$cSkip

    $pnlFBar.Controls.AddRange(@($btnScan,$btnAll,$btnNone,$lblFCnt))
    $grpFiles.Controls.AddRange(@($clv,$pnlFBar))
    $splitV.Panel1.Controls.Add($grpFiles)

    # Log
    $grpLog           = New-Object System.Windows.Forms.GroupBox
    $grpLog.Text      = ' Deployment-Log'
    $grpLog.Font      = $fHead
    $grpLog.ForeColor = $cHeader
    $grpLog.BackColor = $cPanel
    $grpLog.Dock      = 'Fill'
    $grpLog.Padding   = New-Object System.Windows.Forms.Padding(6)

    $rtb              = New-Object System.Windows.Forms.RichTextBox
    $rtb.Dock         = 'Fill'; $rtb.ReadOnly=$true; $rtb.Font=$fMono
    $rtb.BackColor    = $cLogBg; $rtb.ForeColor=$cLogFg; $rtb.ScrollBars='Vertical'; $rtb.WordWrap=$false

    $pnlLBar        = New-Object System.Windows.Forms.Panel
    $pnlLBar.Dock   = 'Bottom'; $pnlLBar.Height=32; $pnlLBar.BackColor=$cPanel

    $btnClrLog = New-Object System.Windows.Forms.Button; $btnClrLog.Text='Log leeren';   $btnClrLog.Width=90;  $btnClrLog.Dock='Left'; $btnClrLog.Font=$fSmall; $btnClrLog.FlatStyle='Flat'
    $btnSavLog = New-Object System.Windows.Forms.Button; $btnSavLog.Text='Log speichern'; $btnSavLog.Width=100; $btnSavLog.Dock='Left'; $btnSavLog.Font=$fSmall; $btnSavLog.FlatStyle='Flat'

    $lblSum           = New-Object System.Windows.Forms.Label
    $lblSum.Text      = ''; $lblSum.Dock='Fill'; $lblSum.TextAlign='MiddleRight'; $lblSum.Font=$fBold

    $pnlLBar.Controls.AddRange(@($btnClrLog,$btnSavLog,$lblSum))
    $grpLog.Controls.AddRange(@($rtb,$pnlLBar))
    $splitV.Panel2.Controls.Add($grpLog)

    # =========================================================================
    # Script-Zustand
    # =========================================================================
    $script:Connected = $false
    $script:ApiBase   = ''
    $script:Cred      = $null

    # =========================================================================
    # Hilfsfunktionen
    # =========================================================================
    function Write-Log {
        param([string]$Msg, [ValidateSet('Info','Success','Warning','Error','Skip','Header')][string]$Lv='Info')
        $col = switch($Lv){'Success'{$cOk}'Warning'{$cWarn}'Error'{$cErr}'Skip'{$cSkip}'Header'{$cAccent}default{$cLogFg}}
        $pfx = switch($Lv){'Success'{'[OK]     '}'Warning'{'[WARN]   '}'Error'{'[FEHLER] '}'Skip'{'[SKIP]   '}'Header'{'─────────'}default{'[INFO]   '}}
        $rtb.SelectionColor = $col
        $rtb.AppendText("[$(Get-Date -Format 'HH:mm:ss')] $pfx $Msg`n")
        $rtb.ScrollToCaret()
    }

    function Get-Cred {
        if ($chkCrd.Checked) {
            $u=$tU.Text.Trim(); $p=$tP.Text
            if([string]::IsNullOrWhiteSpace($u) -or [string]::IsNullOrWhiteSpace($p)){
                [System.Windows.Forms.MessageBox]::Show('Bitte Benutzer und Kennwort eingeben.','Credentials fehlen','OK','Warning')|Out-Null
                return $null
            }
            return New-Object System.Management.Automation.PSCredential($u,(ConvertTo-SecureString $p -AsPlainText -Force))
        }
        return $null
    }

    function Update-DeployBtn {
        $btnDeploy.Enabled = ($script:Connected -and $clv.Items.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($txtTgt.Text))
    }

    function Load-Tree {
        $tv.Nodes.Clear(); $lblTreeStat.Text='Lade ...'
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $folders  = @(Get-SSRSFolders -ApiBase $script:ApiBase -Credential $script:Cred)
            $root     = New-Object System.Windows.Forms.TreeNode('/ (Wurzel)')
            $root.Tag ='/'; $root.ImageIndex=0; $root.SelectedImageIndex=1
            [void]$tv.Nodes.Add($root)
            $map = @{'/'=$root}
            foreach ($f in ($folders | Sort-Object { ($_.Path -split '/').Count })) {
                if ($f.Path -eq '/') { continue }
                $parts  = $f.Path.TrimStart('/') -split '/'
                $pp     = if($parts.Count -eq 1){'/'} else {'/'+($parts[0..($parts.Count-2)] -join '/')}
                $node   = New-Object System.Windows.Forms.TreeNode($parts[-1])
                $node.Tag=$f.Path; $node.ImageIndex=0; $node.SelectedImageIndex=1
                $parent = if($map.ContainsKey($pp)){$map[$pp]} else {$root}
                [void]$parent.Nodes.Add($node)
                $map[$f.Path]=$node
            }
            $root.Expand()
            $lblTreeStat.Text="$($folders.Count) Ordner"
            Write-Log "Ordnerstruktur geladen – $($folders.Count) Ordner" -Lv 'Success'
        } catch {
            $lblTreeStat.Text='Fehler'
            Write-Log "Fehler Ordnerbaum: $($_.Exception.Message)" -Lv 'Error'
        }
    }

    # =========================================================================
    # Events
    # =========================================================================

    # Form_Shown – SplitterDistance erst hier, wenn reale Pixel bekannt sind
    $form.Add_Shown({
        [System.Windows.Forms.Application]::DoEvents()
        $splitH.SplitterDistance = 270
        $splitV.SplitterDistance = [int]($splitV.Height * 0.45)

        Write-Log 'SSRS Deployment Tool v3.0 bereit.' -Lv 'Header'
        Write-Log "Benutzer: $env:USERDOMAIN\$env:USERNAME" -Lv 'Info'

        # Lokalen SSRS suchen
        Write-Log 'Suche lokalen Report Server ...' -Lv 'Info'
        [System.Windows.Forms.Application]::DoEvents()
        $found = Find-LocalSSRS
        if ($found) {
            $txtSrv.Text      = $found
            $txtSrv.ForeColor = [System.Drawing.Color]::Black
            Write-Log "Lokaler Report Server gefunden: $found" -Lv 'Success'
            Write-Log 'URL wurde vorbelegt – bitte Verbinden klicken.' -Lv 'Info'
        } else {
            $txtSrv.Text      = ''
            $txtSrv.ForeColor = [System.Drawing.Color]::Black
            Write-Log 'Kein lokaler Report Server gefunden – URL manuell eingeben.' -Lv 'Warning'
        }

        Write-Log '1. Server-URL  2. Verbinden  3. Ordner  4. Scannen  5. Deployen' -Lv 'Info'
    })

    # Verbinden
    $btnConn.Add_Click({
        $url=$txtSrv.Text.Trim()
        if([string]::IsNullOrWhiteSpace($url)){[System.Windows.Forms.MessageBox]::Show('Bitte Report Server URL eingeben.','Fehlt','OK','Warning')|Out-Null; return}
        $btnConn.Enabled=$false; $btnConn.Text='Verbinde ...'
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $script:Cred=Get-Cred
            if($chkCrd.Checked -and $null -eq $script:Cred){return}

            # ----------------------------------------------------------------
            # HTTP <-> HTTPS Fallback:
            # Schlaegt die eingegebene URL fehl, wird automatisch das jeweils
            # andere Schema probiert (https->http / http->https).
            # Bei HTTPS greift ServerCertificateValidationCallback fuer selbstsignierte Zertifikate.
            # ----------------------------------------------------------------
            $usedUrl   = $url
            $lastError = $null
            $info      = $null

            # Fallback-URL vorberechnen (PS 5.1 erlaubt kein if-Ausdruck direkt in @())
            $altUrl = if ($url -match '^https://') { $url -replace '^https://', 'http://' }
                      else                         { $url -replace '^http://',  'https://' }

            foreach ($candidate in @($url, $altUrl)) {
                try {
                    $info    = Test-SSRSConnection -ServerUrl $candidate -Credential $script:Cred
                    $usedUrl = $candidate
                    $lastError = $null
                    break
                } catch {
                    $lastError = $_
                    $altSchema = if ($candidate -match '^https://') { 'http' } else { 'https' }
                    Write-Log "[$($candidate.Split('//')[0].TrimEnd(':'))] Verbindung fehlgeschlagen – versuche $altSchema ..." -Lv 'Warning'
                }
            }

            if ($null -ne $lastError) { throw $lastError }

            # URL-Feld auf tatsaechlich verwendete URL aktualisieren
            if ($usedUrl -ne $url) {
                $txtSrv.Text = $usedUrl
                Write-Log "URL automatisch auf $usedUrl korrigiert." -Lv 'Warning'
            }

            $pn = if ($info.PSObject.Properties['ProductName'])    { $info.ProductName }    else { 'SSRS' }
            $pv = if ($info.PSObject.Properties['ProductVersion']) { $info.ProductVersion } else { '' }
            $script:ApiBase   = Get-SSRSApiBase $usedUrl
            $script:Connected = $true
            Write-Log "Verbunden: $pn $pv  [$usedUrl]" -Lv 'Success'
            Load-Tree; Update-DeployBtn
        } catch {
            Write-Log "Verbindungsfehler: $($_.Exception.Message)" -Lv 'Error'
            [System.Windows.Forms.MessageBox]::Show("Verbindung fehlgeschlagen:`n$($_.Exception.Message)",'Fehler','OK','Error')|Out-Null
            $script:Connected=$false
        } finally { $btnConn.Enabled=$true; $btnConn.Text='Verbinden' }
    })

    # Tree-Auswahl
    $tv.Add_AfterSelect({
        if($tv.SelectedNode -and $tv.SelectedNode.Tag){
            $txtTgt.Text=$tv.SelectedNode.Tag.ToString(); $lblHint.Text=''
            Write-Log "Zielordner: $($txtTgt.Text)" -Lv 'Info'
            Update-DeployBtn
        }
    })

    # Neuer Ordner
    $mnuNew.Add_Click({
        if(-not $script:Connected){[System.Windows.Forms.MessageBox]::Show('Zuerst verbinden.','','OK','Warning')|Out-Null; return}
        $pn=$tv.SelectedNode
        if($null -eq $pn){[System.Windows.Forms.MessageBox]::Show('Uebergeordneten Ordner auswaehlen.','','OK','Information')|Out-Null; return}
        $pp=$pn.Tag.ToString()

        $dlg=New-Object System.Windows.Forms.Form
        $dlg.Text='Neuen Ordner anlegen'; $dlg.Size=New-Object System.Drawing.Size(400,158)
        $dlg.StartPosition='CenterParent'; $dlg.FormBorderStyle='FixedDialog'; $dlg.MaximizeBox=$false; $dlg.MinimizeBox=$false; $dlg.BackColor=$cPanel

        $lD=New-Object System.Windows.Forms.Label; $lD.Text="Neuer Ordner unter:  $pp"; $lD.Font=$fSmall; $lD.SetBounds(12,14,370,18)
        $tD=New-Object System.Windows.Forms.TextBox; $tD.Font=$fDef; $tD.SetBounds(12,38,370,26)
        $bOk=New-Object System.Windows.Forms.Button; $bOk.Text='Anlegen'; $bOk.Font=$fDef; $bOk.SetBounds(202,78,86,30)
        $bOk.BackColor=$cAccent; $bOk.ForeColor=[System.Drawing.Color]::White; $bOk.FlatStyle='Flat'; $bOk.FlatAppearance.BorderSize=0; $bOk.DialogResult='OK'; $dlg.AcceptButton=$bOk
        $bCan=New-Object System.Windows.Forms.Button; $bCan.Text='Abbrechen'; $bCan.Font=$fSmall; $bCan.SetBounds(296,78,86,30); $bCan.FlatStyle='Flat'; $bCan.DialogResult='Cancel'; $dlg.CancelButton=$bCan
        $dlg.Controls.AddRange(@($lD,$tD,$bOk,$bCan)); $tD.Select()

        if($dlg.ShowDialog($form) -ne 'OK'){return}
        $nn=$tD.Text.Trim()
        if([string]::IsNullOrWhiteSpace($nn)){return}
        if($nn -match '[/\\<>:"|?*]'){[System.Windows.Forms.MessageBox]::Show('Ungueltige Zeichen im Namen.','Fehler','OK','Warning')|Out-Null; return}

        $np=if($pp -eq '/'){"/$nn"} else {"$pp/$nn"}
        try {
            New-SSRSFolderSingle -ApiBase $script:ApiBase -FolderPath $np -Credential $script:Cred
            $node=New-Object System.Windows.Forms.TreeNode($nn); $node.Tag=$np; $node.ImageIndex=0; $node.SelectedImageIndex=1
            [void]$pn.Nodes.Add($node); $pn.Expand(); $tv.SelectedNode=$node
            $txtTgt.Text=$np; $lblHint.Text=''
            $lblTreeStat.Text="Angelegt: $np"
            Write-Log "Ordner angelegt: $np" -Lv 'Success'
            Update-DeployBtn
        } catch {
            Write-Log "Fehler Ordner '$np': $($_.Exception.Message)" -Lv 'Error'
            [System.Windows.Forms.MessageBox]::Show("Fehler:`n$($_.Exception.Message)",'Fehler','OK','Error')|Out-Null
        }
    })

    # Baum aktualisieren
    $mnuRefresh.Add_Click({ if($script:Connected){Load-Tree} })

    # Verzeichnis-Browser
    $btnBrw.Add_Click({
        $fbd=New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description='Quellverzeichnis auswaehlen'; $fbd.ShowNewFolderButton=$false
        if($txtSrc.Text -and (Test-Path $txtSrc.Text)){$fbd.SelectedPath=$txtSrc.Text}
        if($fbd.ShowDialog() -eq 'OK'){$txtSrc.Text=$fbd.SelectedPath; Write-Log "Quellverzeichnis: $($fbd.SelectedPath)" -Lv 'Info'}
    })

    # Scannen
    $btnScan.Add_Click({
        $src=$txtSrc.Text.Trim()
        if(-not(Test-Path $src -PathType Container)){[System.Windows.Forms.MessageBox]::Show("Verzeichnis nicht gefunden:`n$src",'Fehler','OK','Warning')|Out-Null; return}
        $clv.Items.Clear()
        $files=foreach($ext in @('*.rdl','*.rds','*.rsds','*.rsd')){Get-ChildItem -Path $src -Filter $ext -File -ErrorAction SilentlyContinue}
        if(-not $files){Write-Log "Keine Dateien in: $src" -Lv 'Warning'; $lblFCnt.Text='0 Dateien'; return}
        foreach($f in ($files|Sort-Object Name)){
            $typ=switch($f.Extension.ToLower()){'.rdl'{'Report'}'.rds'{'DSrc'}'.rsds'{'DSrc'}'.rsd'{'Dataset'}default{'?'}}
            $lvi=New-Object System.Windows.Forms.ListViewItem($f.Name)
            [void]$lvi.SubItems.Add($typ)
            [void]$lvi.SubItems.Add("$('{0:N1}'-f($f.Length/1KB)) KB")
            [void]$lvi.SubItems.Add($f.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))
            [void]$lvi.SubItems.Add('')
            $lvi.Checked=$true; $lvi.Tag=$f.FullName; [void]$clv.Items.Add($lvi)
        }
        $n=$clv.Items.Count; $lblFCnt.Text="$n Datei(en)"
        Write-Log "Scan: $n Datei(en) – $src" -Lv 'Info'
        Update-DeployBtn
    })

    $btnAll.Add_Click({  foreach($i in $clv.Items){$i.Checked=$true}  })
    $btnNone.Add_Click({ foreach($i in $clv.Items){$i.Checked=$false} })

    $btnClrLog.Add_Click({ $rtb.Clear(); $lblSum.Text='' })
    $btnSavLog.Add_Click({
        $sfd=New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter='Textdatei (*.txt)|*.txt|Alle (*.*)|*.*'
        $sfd.FileName="SSRS_Deploy_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        if($sfd.ShowDialog() -eq 'OK'){$rtb.Text|Out-File -FilePath $sfd.FileName -Encoding UTF8; Write-Log "Log: $($sfd.FileName)" -Lv 'Info'}
    })

    # Deploy
    $btnDeploy.Add_Click({
        $tgt=$txtTgt.Text.Trim()
        if([string]::IsNullOrWhiteSpace($tgt)){[System.Windows.Forms.MessageBox]::Show('Zielordner auswaehlen.','','OK','Warning')|Out-Null; return}
        if(-not $tgt.StartsWith('/')){$tgt="/$tgt"}
        $sel=@($clv.Items|Where-Object{$_.Checked})
        if($sel.Count -eq 0){[System.Windows.Forms.MessageBox]::Show('Keine Dateien ausgewaehlt.','','OK','Information')|Out-Null; return}
        $cred=Get-Cred
        if($chkCrd.Checked -and $null -eq $cred){return}

        $btnDeploy.Enabled=$false; $btnScan.Enabled=$false; $btnConn.Enabled=$false
        $pbProg.Maximum=$sel.Count; $pbProg.Value=0
        $ok=0; $sk=0; $er=0

        Write-Log "═══ Deploy → $tgt  ($($sel.Count) Datei(en)) ═══" -Lv 'Header'
        try { New-SSRSFolderRecursive -ApiBase $script:ApiBase -FolderPath $tgt -Credential $cred }
        catch { Write-Log "Fehler Zielordner: $($_.Exception.Message)" -Lv 'Error'; $btnDeploy.Enabled=$true; $btnScan.Enabled=$true; $btnConn.Enabled=$true; return }

        foreach($item in $sel){
            $fp=$item.Tag; $fn=$item.SubItems[0].Text; $typ=$item.SubItems[1].Text
            try {
                $res=switch($typ){
                    'Report'  {Deploy-Report        -ApiBase $script:ApiBase -FilePath $fp -TargetFolder $tgt -Credential $cred}
                    'DSrc'    {Deploy-DataSource    -ApiBase $script:ApiBase -FilePath $fp -TargetFolder $tgt -Credential $cred}
                    'Dataset' {Deploy-SharedDataset -ApiBase $script:ApiBase -FilePath $fp -TargetFolder $tgt -Credential $cred}
                    default   {'UNBEKANNTER TYP'}
                }
                if($res -like 'SKIPPED*'){
                    Write-Log "$fn → $res" -Lv 'Skip'; $item.SubItems[4].Text=$res; $item.ForeColor=$cSkip; $sk++
                } else {
                    Write-Log "$fn → $res" -Lv 'Success'; $item.SubItems[4].Text=$res; $item.ForeColor=$cOk; $ok++
                }
            } catch {
                $em=$_.Exception.Message
                Write-Log "$fn → FEHLER: $em" -Lv 'Error'; $item.SubItems[4].Text="FEHLER: $em"; $item.ForeColor=$cErr; $er++
            }
            $pbProg.Value++
            [System.Windows.Forms.Application]::DoEvents()
        }

        $sum="Abgeschlossen:  $ok OK  |  $sk Uebersprungen  |  $er Fehler"
        Write-Log $sum -Lv 'Header'
        $lblSum.Text=$sum; $lblSum.ForeColor=if($er-gt 0){$cErr} elseif($sk-gt 0){$cWarn} else{$cOk}
        if($script:Connected){Load-Tree}
        $btnDeploy.Enabled=$true; $btnScan.Enabled=$true; $btnConn.Enabled=$true
        Update-DeployBtn
    })

    [void]$form.ShowDialog()
    $form.Dispose()
}

# =============================================================================
# Einstiegspunkt
# =============================================================================
Show-DeploymentTool