#Requires -Version 5.1
<#
.SYNOPSIS
    SSRS / PBIRS Report Deployment Tool v4.0.0
    WinForms-basiertes Deployment- und Migrations-Tool fuer SQL Server Reporting Services
    sowie Power BI Report Server (PBIRS).

.DESCRIPTION
    - Reports (.rdl)           : werden immer ueberschrieben
    - Power BI Reports (.pbix) : nur auf Power BI Report Server (PBIRS)
    - Datenquellen (.rds/.rsds): bestehende Connections bleiben erhalten, neue werden angelegt
    - Shared Datasets (.rsd)   : werden immer ueberschrieben
    - Authentifizierung        : Windows-Auth oder manuelle Credentials
    - Serverordner             : TreeView links, Rechtsklick Neuer Ordner
    - Migration Tab            : Export von Quell-SSRS (Zwischenverzeichnis), Import in Ziel-SSRS

.NOTES
    Version : 4.0.0
    API     : SSRS REST API v2.0 (SQL Server 2022 / SSRS 16.x / PBIRS)

    RDL-Fixes (automatisch, kein Eingriff in Originaldateien):
    - Undeklarierter df:-Namespace-Praefix wird vor Upload ergaenzt

    Layout-Regeln:
    - Dock=Fill Control ZUERST zu Parent hinzufuegen, danach Dock=Top/Bottom
    - SplitterDistance ausschliesslich in Form_Shown nach DoEvents setzen
    - Kein AutoSize auf GroupBox oder TLP mit Dock=Top
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------------------
# SSL - akzeptiert selbstsignierte / interne Zertifikate
# ---------------------------------------------------------------------------
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {
    param($sender, $certificate, $chain, $sslPolicyErrors)
    return $true
}
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
                $reader.Dispose(); $stream.Dispose()
            }
        } catch { $detail = "(Response-Body nicht lesbar: $($_.Exception.Message))" }
        $sentBody = if ($Body) { $jsonBody } else { '(kein Body)' }
        throw "$($_.Exception.Message) | URI: $Uri | Body: $sentBody | Server-Antwort: $detail"
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
    param([byte[]]$RawBytes)
    try {
        $text = [System.Text.Encoding]::UTF8.GetString($RawBytes).TrimStart([char]0xFEFF)
        if ($text -match 'df:' -and $text -notmatch 'xmlns:df') {
            $dfNs = 'xmlns:df="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition/defaultfontfamily"'
            $text = $text -replace '(<Report\b[^>]*)(>)', "`$1 $dfNs`$2"
        }
        return [System.Text.Encoding]::UTF8.GetBytes($text)
    } catch { return $RawBytes }
}

function Deploy-Report {
    param([string]$ApiBase, [string]$FilePath, [string]$TargetFolder,
          [System.Management.Automation.PSCredential]$Credential = $null)
    $name     = [System.IO.Path]::GetFileNameWithoutExtension($FilePath).Trim()
    $ipath    = "$TargetFolder/$name"
    $rawBytes = [System.IO.File]::ReadAllBytes($FilePath)
    $bytes    = Repair-RDLContent -RawBytes $rawBytes
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
    param([string]$ApiBase, [string]$FilePath, [string]$TargetFolder,
          [System.Management.Automation.PSCredential]$Credential = $null)
    $name  = [System.IO.Path]::GetFileNameWithoutExtension($FilePath).Trim()
    $ipath = "$TargetFolder/$name"
    if (Get-SSRSItemExists -ApiBase $ApiBase -ItemPath $ipath -Credential $Credential) {
        return 'SKIPPED (bereits vorhanden - Connection bleibt erhalten)'
    }
    [xml]$x = Get-Content -Path $FilePath -Encoding UTF8
    $cp     = $x.RptDataSource.ConnectionProperties
    $credRaw = if ($cp.IntegratedSecurity -eq 'true') { 'integrated' }
               elseif (-not [string]::IsNullOrWhiteSpace($cp.CredentialRetrieval)) {
                   switch ($cp.CredentialRetrieval.Trim().ToLower()) {
                       'integrated' { 'integrated' }
                       'prompt'     { 'prompt'     }
                       'store'      { 'store'      }
                       'none'       { 'none'       }
                       default      { 'integrated' }
                   }
               } else { 'integrated' }
    Invoke-SSRSRequest -Uri "$ApiBase/DataSources" -Method 'POST' -Credential $Credential -Body @{
        Name                = $name
        Path                = $ipath
        DataSourceType      = if ($cp.Extension)     { $cp.Extension }     else { 'SQL' }
        ConnectionString    = if ($cp.ConnectString) { $cp.ConnectString } else { '' }
        CredentialRetrieval = $credRaw
    } | Out-Null
    return 'CREATED'
}

function Deploy-SharedDataset {
    param([string]$ApiBase, [string]$FilePath, [string]$TargetFolder,
          [System.Management.Automation.PSCredential]$Credential = $null)
    $name     = [System.IO.Path]::GetFileNameWithoutExtension($FilePath).Trim()
    $ipath    = "$TargetFolder/$name"
    $content  = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($FilePath))
    $existing = Get-SSRSItemExists -ApiBase $ApiBase -ItemPath $ipath -Credential $Credential
    if ($existing) {
        Invoke-SSRSRequest -Uri "$ApiBase/CatalogItems($($existing.Id))" -Method 'DELETE' -Credential $Credential | Out-Null
    }
    Invoke-SSRSRequest -Uri "$ApiBase/DataSets" -Method 'POST' -Credential $Credential `
        -Body @{ Name=$name; Path=$ipath; Content=$content } | Out-Null
    if ($existing) { return 'UPDATED' } else { return 'CREATED' }
}

# =============================================================================
# v4 - Neue Funktionen
# =============================================================================

function Get-SSRSBinaryContent {
    param([string]$Uri, [string]$OutFile,
          [System.Management.Automation.PSCredential]$Credential = $null)
    $p = @{ Uri=$Uri; Method='GET'; OutFile=$OutFile; UseBasicParsing=$true; ErrorAction='Stop' }
    if ($Credential) { $p['Credential']=$Credential } else { $p['UseDefaultCredentials']=$true }
    Invoke-WebRequest @p | Out-Null
}

function Deploy-PowerBIReport {
    param([string]$ApiBase, [string]$FilePath, [string]$TargetFolder,
          [System.Management.Automation.PSCredential]$Credential = $null)
    Add-Type -AssemblyName System.Net.Http
    $name  = [System.IO.Path]::GetFileNameWithoutExtension($FilePath).Trim()
    $ipath = "$TargetFolder/$name"
    $existing = Get-SSRSItemExists -ApiBase $ApiBase -ItemPath $ipath -Credential $Credential
    if ($existing) {
        Invoke-SSRSRequest -Uri "$ApiBase/CatalogItems($($existing.Id))" -Method 'DELETE' -Credential $Credential | Out-Null
    }
    $handler = New-Object System.Net.Http.HttpClientHandler
    if ($Credential) {
        $nc = $Credential.GetNetworkCredential()
        $handler.Credentials = New-Object System.Net.NetworkCredential($nc.UserName, $nc.Password, $nc.Domain)
    } else { $handler.UseDefaultCredentials = $true }
    $client = New-Object System.Net.Http.HttpClient($handler)
    try {
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        $mp    = New-Object System.Net.Http.MultipartFormDataContent
        $fc    = New-Object System.Net.Http.ByteArrayContent(,$bytes)
        $fc.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse('application/octet-stream')
        $mp.Add($fc, 'file', [System.IO.Path]::GetFileName($FilePath))
        $resp = $client.PostAsync("$ApiBase/PowerBIReports", $mp).GetAwaiter().GetResult()
        if (-not $resp.IsSuccessStatusCode) {
            $body = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            throw "HTTP $([int]$resp.StatusCode) $($resp.ReasonPhrase): $body"
        }
    } finally { $client.Dispose() }
    if ($existing) { return 'UPDATED' } else { return 'CREATED' }
}

function Export-SSRSContent {
    param(
        [string]$ApiBase,
        [string]$SourceFolderPath,
        [string]$ExportPath,
        [hashtable]$DsMappings = @{},
        [System.Management.Automation.PSCredential]$Credential = $null
    )
    if (-not (Test-Path $ExportPath)) { New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null }
    $msgs = [System.Collections.Generic.List[string]]::new()
    $ok = 0; $err = 0

    # Ordnerstruktur
    $allFolders = @(Get-SSRSFolders -ApiBase $ApiBase -Credential $Credential)
    foreach ($f in ($allFolders | Where-Object { $_.Path -like "$SourceFolderPath/*" })) {
        $rel = $f.Path.Substring($SourceFolderPath.Length).TrimStart('/')
        $d   = Join-Path $ExportPath ($rel -replace '/', '\')
        if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    }

    # Reports -> .rdl
    try {
        $allRpts = @((Invoke-SSRSRequest -Uri "$ApiBase/Reports" -Credential $Credential).value)
        foreach ($r in ($allRpts | Where-Object { $_.Path -like "$SourceFolderPath/*" })) {
            try {
                $rel  = $r.Path.Substring($SourceFolderPath.Length).TrimStart('/')
                $file = Join-Path $ExportPath (($rel -replace '/', '\') + '.rdl')
                $dir  = Split-Path $file -Parent
                if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
                Get-SSRSBinaryContent -Uri "$ApiBase/Reports($($r.Id))/Content/`$value" -OutFile $file -Credential $Credential
                $msgs.Add("[OK] RDL: $($r.Path)"); $ok++
            } catch { $msgs.Add("[ERR] RDL $($r.Path): $($_.Exception.Message)"); $err++ }
        }
    } catch { $msgs.Add("[ERR] Reports-Liste: $($_.Exception.Message)"); $err++ }

    # DataSources -> .rds
    try {
        $allDs = @((Invoke-SSRSRequest -Uri "$ApiBase/DataSources" -Credential $Credential).value)
        foreach ($ds in ($allDs | Where-Object { $_.Path -like "$SourceFolderPath/*" })) {
            try {
                $rel    = $ds.Path.Substring($SourceFolderPath.Length).TrimStart('/')
                $file   = Join-Path $ExportPath (($rel -replace '/', '\') + '.rds')
                $dir    = Split-Path $file -Parent
                if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
                $detail = Invoke-SSRSRequest -Uri "$ApiBase/DataSources($($ds.Id))" -Credential $Credential
                $conn   = if ($detail.ConnectionString) { $detail.ConnectionString } else { '' }
                foreach ($key in $DsMappings.Keys) {
                    if ($conn -and $key) { $conn = $conn -replace [regex]::Escape($key), $DsMappings[$key] }
                }
                $intSec = if ($detail.CredentialRetrieval -ieq 'integrated') { 'true' } else { 'false' }
                $xml = "<?xml version=""1.0"" encoding=""utf-8""?>`r`n" +
                       "<RptDataSource xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " +
                       "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" Name=""$($ds.Name)"">`r`n" +
                       "  <ConnectionProperties>`r`n" +
                       "    <Extension>$($detail.DataSourceType)</Extension>`r`n" +
                       "    <ConnectString>$conn</ConnectString>`r`n" +
                       "    <IntegratedSecurity>$intSec</IntegratedSecurity>`r`n" +
                       "  </ConnectionProperties>`r`n</RptDataSource>"
                [System.IO.File]::WriteAllText($file, $xml, [System.Text.Encoding]::UTF8)
                $msgs.Add("[OK] RDS: $($ds.Path)"); $ok++
            } catch { $msgs.Add("[ERR] RDS $($ds.Path): $($_.Exception.Message)"); $err++ }
        }
    } catch { $msgs.Add("[ERR] DataSources-Liste: $($_.Exception.Message)"); $err++ }

    # Shared Datasets -> .rsd
    try {
        $allRsd = @((Invoke-SSRSRequest -Uri "$ApiBase/DataSets" -Credential $Credential).value)
        foreach ($rsd in ($allRsd | Where-Object { $_.Path -like "$SourceFolderPath/*" })) {
            try {
                $rel  = $rsd.Path.Substring($SourceFolderPath.Length).TrimStart('/')
                $file = Join-Path $ExportPath (($rel -replace '/', '\') + '.rsd')
                $dir  = Split-Path $file -Parent
                if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
                Get-SSRSBinaryContent -Uri "$ApiBase/DataSets($($rsd.Id))/Content/`$value" -OutFile $file -Credential $Credential
                $msgs.Add("[OK] RSD: $($rsd.Path)"); $ok++
            } catch { $msgs.Add("[ERR] RSD $($rsd.Path): $($_.Exception.Message)"); $err++ }
        }
    } catch { $msgs.Add("[ERR] DataSets-Liste: $($_.Exception.Message)"); $err++ }

    return [PSCustomObject]@{ Exported=$ok; Errors=$err; Messages=$msgs; ExportPath=$ExportPath }
}

# =============================================================================
# Lokalen SSRS-Server erkennen
# =============================================================================

function Find-LocalSSRS {
    $hostname   = $env:COMPUTERNAME
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
            Invoke-RestMethod @p | Out-Null
            return (Get-SSRSApiBase -ServerUrl $url) -replace '/api/v2\.0$', ''
        } catch { }
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
    # Layout-Konstanten
    # -------------------------------------------------------------------------
    $cfgRowH   = 34
    $cfgRows   = 5
    $cfgPadTB  = 10
    $cfgGrpTop = 22
    $cfgH      = $cfgRows * $cfgRowH + $cfgGrpTop + $cfgPadTB * 2 + 4
    $sepH      = 8

    # -------------------------------------------------------------------------
    # Hauptfenster
    # -------------------------------------------------------------------------
    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'dtcSoftware Report Deployment Tool  v4.0'
    $form.Size            = New-Object System.Drawing.Size(1080, 840)
    $form.MinimumSize     = New-Object System.Drawing.Size(900, 700)
    $form.StartPosition   = 'CenterScreen'
    $form.BackColor       = $cBg
    $form.Font            = $fDef
    $form.FormBorderStyle = 'Sizable'

    # -------------------------------------------------------------------------
    # Header (Dock=Top)
    # -------------------------------------------------------------------------
    $pnlHeader           = New-Object System.Windows.Forms.Panel
    $pnlHeader.Dock      = 'Top'
    $pnlHeader.Height    = 35
    $pnlHeader.BackColor = $cHeader
    $pnlHeader.Padding   = New-Object System.Windows.Forms.Padding(14, 0, 14, 0)

    $lblTitle           = New-Object System.Windows.Forms.Label
    $lblTitle.Text      = 'dtcSoftware Report Deployment Tool'
    $lblTitle.ForeColor = [System.Drawing.Color]::White
    $lblTitle.Font      = $fTitle
    $lblTitle.Dock      = 'Fill'
    $lblTitle.TextAlign = 'MiddleLeft'

    $lblVer             = New-Object System.Windows.Forms.Label
    $lblVer.Text        = '2026 v4.0  |  REST API v2.0'
    $lblVer.ForeColor   = [System.Drawing.Color]::FromArgb(170, 205, 240)
    $lblVer.Font        = $fSmall
    $lblVer.Dock        = 'Right'
    $lblVer.Width       = 170
    $lblVer.TextAlign   = 'MiddleRight'

    $pnlHeader.Controls.AddRange(@($lblTitle, $lblVer))
    $form.Controls.Add($pnlHeader)

    # -------------------------------------------------------------------------
    # TabControl (Dock=Fill) - nach Header hinzufuegen
    # -------------------------------------------------------------------------
    $tabCtrl            = New-Object System.Windows.Forms.TabControl
    $tabCtrl.Dock       = 'Fill'
    $tabCtrl.Font       = $fHead
    $tabCtrl.Appearance = 'Normal'
    $tabCtrl.BackColor  = $cBg
    $form.Controls.Add($tabCtrl)

    # =========================================================================
    # TAB 1 - Deploy
    # =========================================================================
    $tabDeploy           = New-Object System.Windows.Forms.TabPage('  Deploy  ')
    $tabDeploy.BackColor = $cBg
    [void]$tabCtrl.TabPages.Add($tabDeploy)

    $pnlD           = New-Object System.Windows.Forms.Panel
    $pnlD.Dock      = 'Fill'
    $pnlD.BackColor = $cBg
    $pnlD.Padding   = New-Object System.Windows.Forms.Padding(10, 10, 10, 6)
    $tabDeploy.Controls.Add($pnlD)

    # --- Konfiguration GroupBox ---
    $pnlCfgWrap           = New-Object System.Windows.Forms.Panel
    $pnlCfgWrap.Dock      = 'Top'
    $pnlCfgWrap.Height    = $cfgH
    $pnlCfgWrap.BackColor = $cBg

    $grpCfg           = New-Object System.Windows.Forms.GroupBox
    $grpCfg.Text      = ' Konfiguration'
    $grpCfg.Font      = $fHead
    $grpCfg.ForeColor = $cHeader
    $grpCfg.BackColor = $cPanel
    $grpCfg.Dock      = 'Fill'
    $grpCfg.Padding   = New-Object System.Windows.Forms.Padding(10, $cfgPadTB, 10, $cfgPadTB)

    $tlpCfg             = New-Object System.Windows.Forms.TableLayoutPanel
    $tlpCfg.Dock        = 'Fill'
    $tlpCfg.ColumnCount = 3
    $tlpCfg.RowCount    = $cfgRows
    $tlpCfg.BackColor   = $cPanel
    [void]$tlpCfg.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 175)))
    [void]$tlpCfg.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$tlpCfg.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 116)))
    for ($i = 0; $i -lt $cfgRows; $i++) {
        [void]$tlpCfg.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $cfgRowH)))
    }

    function New-CfgLabel  { param([string]$t) $l=New-Object System.Windows.Forms.Label; $l.Text=$t; $l.Font=$fBold; $l.TextAlign='MiddleRight'; $l.Dock='Fill'; $l.AutoSize=$false; $l.AutoEllipsis=$true; $l }
    function New-CfgText   { param([string]$p='') $t=New-Object System.Windows.Forms.TextBox; $t.Dock='Fill'; $t.Font=$fDef; if($p){$t.Text=$p}; $t }
    function New-CfgButton { param([string]$t) $b=New-Object System.Windows.Forms.Button; $b.Text=$t; $b.Dock='Fill'; $b.Font=$fDef; $b.BackColor=$cAccent; $b.ForeColor=[System.Drawing.Color]::White; $b.FlatStyle='Flat'; $b.FlatAppearance.BorderSize=0; $b.Cursor='Hand'; $b }

    $lblSrv  = New-CfgLabel  'Report Server URL:'
    $txtSrv  = New-CfgText   ''
    $txtSrv.ForeColor = [System.Drawing.Color]::FromArgb(130,130,130)
    $txtSrv.Text      = 'Wird gesucht ...'
    $btnConn = New-CfgButton 'Verbinden'

    $lblSrc  = New-CfgLabel  'Quellverzeichnis:'
    $txtSrc  = New-CfgText   ''
    $btnBrw  = New-CfgButton 'Durchsuchen'

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

    $lblCrd  = New-CfgLabel 'Authentifizierung:'
    $chkCrd  = New-Object System.Windows.Forms.CheckBox
    $chkCrd.Text='Manuelle Credentials verwenden'; $chkCrd.Font=$fDef; $chkCrd.Dock='Fill'

    $pnlCrd         = New-Object System.Windows.Forms.Panel
    $pnlCrd.Dock    = 'Fill'
    $pnlCrd.Enabled = $false

    $tlpCr = New-Object System.Windows.Forms.TableLayoutPanel
    $tlpCr.Dock='Fill'; $tlpCr.ColumnCount=4; $tlpCr.RowCount=1; $tlpCr.BackColor=$cPanel
    [void]$tlpCr.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,  62)))
    [void]$tlpCr.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,  50)))
    [void]$tlpCr.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,  80)))
    [void]$tlpCr.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,  50)))

    $lU=New-Object System.Windows.Forms.Label; $lU.Text='User:';     $lU.Font=$fBold; $lU.Dock='Fill'; $lU.TextAlign='MiddleRight'
    $tU=New-Object System.Windows.Forms.TextBox; $tU.Dock='Fill';    $tU.Font=$fDef;  $tU.Text="$env:USERDOMAIN\$env:USERNAME"
    $lP=New-Object System.Windows.Forms.Label; $lP.Text='Kennwort:'; $lP.Font=$fBold; $lP.Dock='Fill'; $lP.TextAlign='MiddleRight'
    $tP=New-Object System.Windows.Forms.TextBox; $tP.Dock='Fill';    $tP.Font=$fDef;  $tP.PasswordChar=[char]0x2022

    $tlpCr.Controls.Add($lU,0,0); $tlpCr.Controls.Add($tU,1,0)
    $tlpCr.Controls.Add($lP,2,0); $tlpCr.Controls.Add($tP,3,0)
    $pnlCrd.Controls.Add($tlpCr)
    $chkCrd.Add_CheckedChanged({ $pnlCrd.Enabled = $chkCrd.Checked })

    $tlpCfg.Controls.Add($lblSrv, 0,0); $tlpCfg.Controls.Add($txtSrv, 1,0); $tlpCfg.Controls.Add($btnConn, 2,0)
    $tlpCfg.Controls.Add($lblSrc, 0,1); $tlpCfg.Controls.Add($txtSrc, 1,1); $tlpCfg.Controls.Add($btnBrw,  2,1)
    $tlpCfg.Controls.Add($lblTgt, 0,2); $tlpCfg.Controls.Add($txtTgt, 1,2); $tlpCfg.Controls.Add($pnlHint, 2,2)
    $tlpCfg.Controls.Add($lblCrd, 0,3); $tlpCfg.Controls.Add($chkCrd, 1,3); $tlpCfg.Controls.Add((New-Object System.Windows.Forms.Panel),2,3)
    $tlpCfg.Controls.Add((New-Object System.Windows.Forms.Panel),0,4); $tlpCfg.Controls.Add($pnlCrd,1,4)

    $grpCfg.Controls.Add($tlpCfg)
    $pnlCfgWrap.Controls.Add($grpCfg)

    # --- splitH (Fill ZUERST), dann sep + cfgWrap (Top) ---
    $splitH               = New-Object System.Windows.Forms.SplitContainer
    $splitH.Dock          = 'Fill'
    $splitH.Orientation   = 'Vertical'
    $splitH.BorderStyle   = 'None'
    $splitH.BackColor     = $cBg
    $splitH.Panel1MinSize = 160
    $splitH.Panel2MinSize = 160
    $pnlD.Controls.Add($splitH)

    $sep        = New-Object System.Windows.Forms.Panel
    $sep.Dock   = 'Top'
    $sep.Height = $sepH
    $sep.BackColor = $cBg
    $pnlD.Controls.Add($sep)
    $pnlD.Controls.Add($pnlCfgWrap)

    # --- Panel1: TreeView ---
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

    # Ordner-Icons
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
    $imgList.Images.Add($bmpClosed)
    $imgList.Images.Add($bmpOpen)

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

    $ctxTree         = New-Object System.Windows.Forms.ContextMenuStrip
    $mnuNew          = New-Object System.Windows.Forms.ToolStripMenuItem('Neuen Ordner anlegen')
    $mnuNew.Font     = $fDef
    $mnuRefresh      = New-Object System.Windows.Forms.ToolStripMenuItem('Baumstruktur aktualisieren')
    $mnuRefresh.Font = $fDef
    [void]$ctxTree.Items.Add($mnuNew)
    [void]$ctxTree.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    [void]$ctxTree.Items.Add($mnuRefresh)
    $tv.ContextMenuStrip = $ctxTree

    $grpTree.Controls.AddRange(@($tv, $lblTreeStat))
    $splitH.Panel1.Controls.Add($grpTree)

    # --- Panel2: pnlDep (Bottom ZUERST), dann splitV (Fill) ---
    $pnlDep         = New-Object System.Windows.Forms.Panel
    $pnlDep.Dock    = 'Bottom'
    $pnlDep.Height  = 46
    $pnlDep.BackColor = $cBg
    $pnlDep.Padding = New-Object System.Windows.Forms.Padding(0,5,0,0)

    $btnDeploy           = New-Object System.Windows.Forms.Button
    $btnDeploy.Text      = '  Deployment starten'
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
    $splitH.Panel2.Controls.Add($pnlDep)

    $splitV               = New-Object System.Windows.Forms.SplitContainer
    $splitV.Dock          = 'Fill'
    $splitV.Orientation   = 'Horizontal'
    $splitV.BorderStyle   = 'None'
    $splitV.BackColor     = $cBg
    $splitV.Panel1MinSize = 80
    $splitV.Panel2MinSize = 80
    $splitH.Panel2.Controls.Add($splitV)

    # splitV.Panel1: Dateiliste
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
    [void]$clv.Columns.Add('Typ',        72)
    [void]$clv.Columns.Add('Groesse',    72)
    [void]$clv.Columns.Add('Geaendert', 148)
    [void]$clv.Columns.Add('Status',    215)

    $pnlFBar        = New-Object System.Windows.Forms.Panel
    $pnlFBar.Dock   = 'Bottom'; $pnlFBar.Height=32; $pnlFBar.BackColor=$cPanel

    $btnScan           = New-Object System.Windows.Forms.Button
    $btnScan.Text      = 'Scannen'; $btnScan.Width=90; $btnScan.Dock='Left'; $btnScan.Font=$fDef
    $btnScan.BackColor = $cAccent; $btnScan.ForeColor=[System.Drawing.Color]::White
    $btnScan.FlatStyle = 'Flat'; $btnScan.FlatAppearance.BorderSize=0; $btnScan.Cursor='Hand'

    $btnAll  = New-Object System.Windows.Forms.Button; $btnAll.Text='Alle';   $btnAll.Width=52;  $btnAll.Dock='Left'; $btnAll.Font=$fSmall; $btnAll.FlatStyle='Flat'
    $btnNone = New-Object System.Windows.Forms.Button; $btnNone.Text='Keine'; $btnNone.Width=52; $btnNone.Dock='Left'; $btnNone.Font=$fSmall; $btnNone.FlatStyle='Flat'

    $lblFCnt           = New-Object System.Windows.Forms.Label
    $lblFCnt.Text      = 'Noch kein Scan'; $lblFCnt.Dock='Fill'; $lblFCnt.TextAlign='MiddleRight'; $lblFCnt.Font=$fSmall; $lblFCnt.ForeColor=$cSkip

    $pnlFBar.Controls.AddRange(@($btnScan,$btnAll,$btnNone,$lblFCnt))
    $grpFiles.Controls.AddRange(@($clv,$pnlFBar))
    $splitV.Panel1.Controls.Add($grpFiles)

    # splitV.Panel2: Log
    $grpLog           = New-Object System.Windows.Forms.GroupBox
    $grpLog.Text      = ' Deployment-Log'
    $grpLog.Font      = $fHead
    $grpLog.ForeColor = $cHeader
    $grpLog.BackColor = $cPanel
    $grpLog.Dock      = 'Fill'
    $grpLog.Padding   = New-Object System.Windows.Forms.Padding(6)

    $rtb           = New-Object System.Windows.Forms.RichTextBox
    $rtb.Dock      = 'Fill'; $rtb.ReadOnly=$true; $rtb.Font=$fMono
    $rtb.BackColor = $cLogBg; $rtb.ForeColor=$cLogFg; $rtb.ScrollBars='Vertical'; $rtb.WordWrap=$false

    $pnlLBar = New-Object System.Windows.Forms.Panel; $pnlLBar.Dock='Bottom'; $pnlLBar.Height=32; $pnlLBar.BackColor=$cPanel
    $btnClrLog = New-Object System.Windows.Forms.Button; $btnClrLog.Text='Log leeren';    $btnClrLog.Width=90;  $btnClrLog.Dock='Left'; $btnClrLog.Font=$fSmall; $btnClrLog.FlatStyle='Flat'
    $btnSavLog = New-Object System.Windows.Forms.Button; $btnSavLog.Text='Log speichern'; $btnSavLog.Width=100; $btnSavLog.Dock='Left'; $btnSavLog.Font=$fSmall; $btnSavLog.FlatStyle='Flat'
    $lblSum    = New-Object System.Windows.Forms.Label; $lblSum.Text=''; $lblSum.Dock='Fill'; $lblSum.TextAlign='MiddleRight'; $lblSum.Font=$fBold
    $pnlLBar.Controls.AddRange(@($btnClrLog,$btnSavLog,$lblSum))
    $grpLog.Controls.AddRange(@($rtb,$pnlLBar))
    $splitV.Panel2.Controls.Add($grpLog)

    # =========================================================================
    # TAB 2 - Migration
    # =========================================================================
    $tabMigrate           = New-Object System.Windows.Forms.TabPage('  Migration  ')
    $tabMigrate.BackColor = $cBg
    [void]$tabCtrl.TabPages.Add($tabMigrate)

    $pnlM           = New-Object System.Windows.Forms.Panel
    $pnlM.Dock      = 'Fill'
    $pnlM.BackColor = $cBg
    $pnlM.Padding   = New-Object System.Windows.Forms.Padding(10, 10, 10, 6)
    $tabMigrate.Controls.Add($pnlM)

    # --- splitMig (Fill) ZUERST, dann pnlMigBot (Bottom), pnlMigOpts (Top), pnlMigSrv (Top) ---
    $splitMig               = New-Object System.Windows.Forms.SplitContainer
    $splitMig.Dock          = 'Fill'
    $splitMig.Orientation   = 'Vertical'
    $splitMig.BorderStyle   = 'None'
    $splitMig.BackColor     = $cBg
    $splitMig.Panel1MinSize = 160
    $splitMig.Panel2MinSize = 160
    $pnlM.Controls.Add($splitMig)

    # --- pnlMigBot (Dock=Bottom, h=230) ---
    $pnlMigBot        = New-Object System.Windows.Forms.Panel
    $pnlMigBot.Dock   = 'Bottom'
    $pnlMigBot.Height = 230
    $pnlMigBot.BackColor = $cBg

    # grpMigLog (Fill) ZUERST
    $grpMigLog           = New-Object System.Windows.Forms.GroupBox
    $grpMigLog.Text      = ' Migrations-Log'
    $grpMigLog.Font      = $fHead
    $grpMigLog.ForeColor = $cHeader
    $grpMigLog.BackColor = $cPanel
    $grpMigLog.Dock      = 'Fill'
    $grpMigLog.Padding   = New-Object System.Windows.Forms.Padding(6)

    $rtbMig           = New-Object System.Windows.Forms.RichTextBox
    $rtbMig.Dock      = 'Fill'; $rtbMig.ReadOnly=$true; $rtbMig.Font=$fMono
    $rtbMig.BackColor = $cLogBg; $rtbMig.ForeColor=$cLogFg; $rtbMig.ScrollBars='Vertical'; $rtbMig.WordWrap=$false

    $pnlMigLBar = New-Object System.Windows.Forms.Panel; $pnlMigLBar.Dock='Bottom'; $pnlMigLBar.Height=32; $pnlMigLBar.BackColor=$cPanel
    $btnMigClrLog = New-Object System.Windows.Forms.Button; $btnMigClrLog.Text='Log leeren';    $btnMigClrLog.Width=90;  $btnMigClrLog.Dock='Left'; $btnMigClrLog.Font=$fSmall; $btnMigClrLog.FlatStyle='Flat'
    $btnMigSavLog = New-Object System.Windows.Forms.Button; $btnMigSavLog.Text='Log speichern'; $btnMigSavLog.Width=100; $btnMigSavLog.Dock='Left'; $btnMigSavLog.Font=$fSmall; $btnMigSavLog.FlatStyle='Flat'
    $lblMigStat   = New-Object System.Windows.Forms.Label; $lblMigStat.Text=''; $lblMigStat.Dock='Fill'; $lblMigStat.TextAlign='MiddleRight'; $lblMigStat.Font=$fSmall; $lblMigStat.ForeColor=$cSkip
    $pnlMigLBar.Controls.AddRange(@($btnMigClrLog,$btnMigSavLog,$lblMigStat))
    $grpMigLog.Controls.AddRange(@($rtbMig,$pnlMigLBar))
    $pnlMigBot.Controls.Add($grpMigLog)

    # pnlDsWrap (Top, h=0, initial hidden)
    $pnlDsWrap        = New-Object System.Windows.Forms.Panel
    $pnlDsWrap.Dock   = 'Top'
    $pnlDsWrap.Height = 0
    $pnlDsWrap.Visible = $false
    $pnlDsWrap.BackColor = $cBg

    $dgvDs                            = New-Object System.Windows.Forms.DataGridView
    $dgvDs.Dock                       = 'Fill'
    $dgvDs.Font                       = $fMono
    $dgvDs.RowHeadersVisible          = $false
    $dgvDs.AllowUserToAddRows         = $true
    $dgvDs.BackgroundColor            = $cPanel
    $dgvDs.BorderStyle                = 'FixedSingle'
    $dgvDs.ColumnHeadersHeightSizeMode = 'DisableResizing'
    $dgvDs.AutoSizeColumnsMode        = 'Fill'
    [void]$dgvDs.Columns.Add('ColSrc', 'Quell-ConnectionString')
    [void]$dgvDs.Columns.Add('ColTgt', 'Ziel-ConnectionString')
    $pnlDsWrap.Controls.Add($dgvDs)
    $pnlMigBot.Controls.Add($pnlDsWrap)

    # pnlMigAct (Top, h=46)
    $pnlMigAct         = New-Object System.Windows.Forms.Panel
    $pnlMigAct.Dock    = 'Top'
    $pnlMigAct.Height  = 46
    $pnlMigAct.BackColor = $cBg
    $pnlMigAct.Padding = New-Object System.Windows.Forms.Padding(0,5,0,0)

    $btnMig           = New-Object System.Windows.Forms.Button
    $btnMig.Text      = '  Migrieren'
    $btnMig.Font      = New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
    $btnMig.Dock      = 'Right'
    $btnMig.Width     = 215
    $btnMig.BackColor = $cHeader
    $btnMig.ForeColor = [System.Drawing.Color]::White
    $btnMig.FlatStyle = 'Flat'
    $btnMig.FlatAppearance.BorderSize = 0
    $btnMig.Cursor    = 'Hand'
    $btnMig.Enabled   = $false

    $pbMig         = New-Object System.Windows.Forms.ProgressBar
    $pbMig.Dock    = 'Fill'
    $pbMig.Minimum = 0
    $pbMig.Value   = 0
    $pbMig.Style   = 'Continuous'

    $pnlMigAct.Controls.AddRange(@($btnMig, $pbMig))
    $pnlMigBot.Controls.Add($pnlMigAct)
    $pnlM.Controls.Add($pnlMigBot)

    # --- pnlMigOpts (Dock=Top, h=72) ---
    $pnlMigOpts        = New-Object System.Windows.Forms.Panel
    $pnlMigOpts.Dock   = 'Top'
    $pnlMigOpts.Height = 72
    $pnlMigOpts.BackColor = $cBg

    $grpOpts           = New-Object System.Windows.Forms.GroupBox
    $grpOpts.Text      = ' Export-Einstellungen'
    $grpOpts.Font      = $fHead
    $grpOpts.ForeColor = $cHeader
    $grpOpts.BackColor = $cPanel
    $grpOpts.Dock      = 'Fill'
    $grpOpts.Padding   = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)

    $tlpOpts             = New-Object System.Windows.Forms.TableLayoutPanel
    $tlpOpts.Dock        = 'Fill'
    $tlpOpts.ColumnCount = 3
    $tlpOpts.RowCount    = 2
    $tlpOpts.BackColor   = $cPanel
    [void]$tlpOpts.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 160)))
    [void]$tlpOpts.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$tlpOpts.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 100)))
    [void]$tlpOpts.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26)))
    [void]$tlpOpts.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26)))

    $lblExpPath = New-Object System.Windows.Forms.Label; $lblExpPath.Text='Zwischenverzeichnis:'; $lblExpPath.Font=$fBold; $lblExpPath.Dock='Fill'; $lblExpPath.TextAlign='MiddleRight'
    $txtExpPath = New-Object System.Windows.Forms.TextBox; $txtExpPath.Dock='Fill'; $txtExpPath.Font=$fDef
    $btnExpBrw  = New-CfgButton 'Durchsuchen'
    $lblDsMap   = New-Object System.Windows.Forms.Label; $lblDsMap.Text='DS-Mapping:'; $lblDsMap.Font=$fBold; $lblDsMap.Dock='Fill'; $lblDsMap.TextAlign='MiddleRight'
    $chkDsMap   = New-Object System.Windows.Forms.CheckBox; $chkDsMap.Text='Connection Strings anpassen (Tabelle unten)'; $chkDsMap.Font=$fDef; $chkDsMap.Dock='Fill'

    $tlpOpts.Controls.Add($lblExpPath, 0,0); $tlpOpts.Controls.Add($txtExpPath, 1,0); $tlpOpts.Controls.Add($btnExpBrw, 2,0)
    $tlpOpts.Controls.Add($lblDsMap,   0,1); $tlpOpts.Controls.Add($chkDsMap,   1,1)

    $grpOpts.Controls.Add($tlpOpts)
    $pnlMigOpts.Controls.Add($grpOpts)
    $pnlM.Controls.Add($pnlMigOpts)

    # --- pnlMigSrv (Dock=Top, h=120) ---
    $pnlMigSrv        = New-Object System.Windows.Forms.Panel
    $pnlMigSrv.Dock   = 'Top'
    $pnlMigSrv.Height = 120
    $pnlMigSrv.BackColor = $cBg

    $tlpSrvRow             = New-Object System.Windows.Forms.TableLayoutPanel
    $tlpSrvRow.Dock        = 'Fill'
    $tlpSrvRow.ColumnCount = 2
    $tlpSrvRow.RowCount    = 1
    $tlpSrvRow.BackColor   = $cBg
    [void]$tlpSrvRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    [void]$tlpSrvRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))

    function New-SrvGroupBox {
        param([string]$Title, [ref]$OutSrv, [ref]$OutBtn, [ref]$OutUser, [ref]$OutPwd, [ref]$OutChk)
        $grp = New-Object System.Windows.Forms.GroupBox
        $grp.Text = " $Title"; $grp.Font=$fHead; $grp.ForeColor=$cHeader; $grp.BackColor=$cPanel; $grp.Dock='Fill'
        $grp.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)

        $tlp = New-Object System.Windows.Forms.TableLayoutPanel
        $tlp.Dock='Fill'; $tlp.ColumnCount=3; $tlp.RowCount=3; $tlp.BackColor=$cPanel
        [void]$tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
        [void]$tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
        [void]$tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 100)))
        for ($ri=0;$ri-lt 3;$ri++){[void]$tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26)))}

        $lSrv=New-Object System.Windows.Forms.Label; $lSrv.Text='Report Server URL:'; $lSrv.Font=$fSmall; $lSrv.Dock='Fill'; $lSrv.TextAlign='MiddleRight'
        $tSrv=New-Object System.Windows.Forms.TextBox; $tSrv.Dock='Fill'; $tSrv.Font=$fDef
        $bCon=New-CfgButton 'Verbinden'

        $lUsr=New-Object System.Windows.Forms.Label; $lUsr.Text='Benutzer:'; $lUsr.Font=$fSmall; $lUsr.Dock='Fill'; $lUsr.TextAlign='MiddleRight'
        $tUsr=New-Object System.Windows.Forms.TextBox; $tUsr.Dock='Fill'; $tUsr.Font=$fSmall; $tUsr.Text="$env:USERDOMAIN\$env:USERNAME"; $tUsr.Enabled=$false
        $lPwd=New-Object System.Windows.Forms.Label; $lPwd.Text='Kennwort:'; $lPwd.Font=$fSmall; $lPwd.Dock='Fill'; $lPwd.TextAlign='MiddleRight'
        $tPwd=New-Object System.Windows.Forms.TextBox; $tPwd.Dock='Fill'; $tPwd.Font=$fSmall; $tPwd.PasswordChar=[char]0x2022; $tPwd.Enabled=$false

        $pnlRow1=New-Object System.Windows.Forms.Panel; $pnlRow1.Dock='Fill'
        $pnlRow2=New-Object System.Windows.Forms.Panel; $pnlRow2.Dock='Fill'

        $chkC=New-Object System.Windows.Forms.CheckBox; $chkC.Text='Manuelle Credentials'; $chkC.Font=$fSmall; $chkC.Dock='Bottom'; $chkC.Height=22
        $chkC.Add_CheckedChanged({ $tUsr.Enabled=$chkC.Checked; $tPwd.Enabled=$chkC.Checked })

        $tlp.Controls.Add($lSrv,0,0); $tlp.Controls.Add($tSrv,1,0); $tlp.Controls.Add($bCon,2,0)
        $tlp.Controls.Add($lUsr,0,1); $tlp.Controls.Add($tUsr,1,1); $tlp.Controls.Add($pnlRow1,2,1)
        $tlp.Controls.Add($lPwd,0,2); $tlp.Controls.Add($tPwd,1,2); $tlp.Controls.Add($pnlRow2,2,2)

        $grp.Controls.Add($tlp)
        $grp.Controls.Add($chkC)

        $OutSrv.Value=$tSrv; $OutBtn.Value=$bCon; $OutUser.Value=$tUsr; $OutPwd.Value=$tPwd; $OutChk.Value=$chkC
        return $grp
    }

    $refSrcSrv=$null;$refSrcBtn=$null;$refSrcUser=$null;$refSrcPwd=$null;$refSrcChk=$null
    $refTgtSrv=$null;$refTgtBtn=$null;$refTgtUser=$null;$refTgtPwd=$null;$refTgtChk=$null

    $grpSrc = New-SrvGroupBox 'Quell-Server' ([ref]$refSrcSrv) ([ref]$refSrcBtn) ([ref]$refSrcUser) ([ref]$refSrcPwd) ([ref]$refSrcChk)
    $grpTgt = New-SrvGroupBox 'Ziel-Server'  ([ref]$refTgtSrv) ([ref]$refTgtBtn) ([ref]$refTgtUser) ([ref]$refTgtPwd) ([ref]$refTgtChk)

    $txtSrcSrv=$refSrcSrv; $btnSrcConn=$refSrcBtn; $txtSrcUser=$refSrcUser; $txtSrcPwd=$refSrcPwd; $chkSrcCrd=$refSrcChk
    $txtTgtSrv=$refTgtSrv; $btnTgtConn=$refTgtBtn; $txtTgtUser=$refTgtUser; $txtTgtPwd=$refTgtPwd; $chkTgtCrd=$refTgtChk

    $tlpSrvRow.Controls.Add($grpSrc,0,0)
    $tlpSrvRow.Controls.Add($grpTgt,1,0)
    $pnlMigSrv.Controls.Add($tlpSrvRow)
    $pnlM.Controls.Add($pnlMigSrv)

    # --- Mig-TreeViews ---
    $grpSrcTree           = New-Object System.Windows.Forms.GroupBox
    $grpSrcTree.Text      = ' Quell-Ordner'
    $grpSrcTree.Font      = $fHead
    $grpSrcTree.ForeColor = $cHeader
    $grpSrcTree.BackColor = $cPanel
    $grpSrcTree.Dock      = 'Fill'
    $grpSrcTree.Padding   = New-Object System.Windows.Forms.Padding(6,4,6,4)

    $tvSrc               = New-Object System.Windows.Forms.TreeView
    $tvSrc.Dock          = 'Fill'; $tvSrc.Font=$fDef; $tvSrc.BackColor=$cPanel; $tvSrc.BorderStyle='FixedSingle'
    $tvSrc.HideSelection = $false; $tvSrc.ShowRootLines=$true; $tvSrc.FullRowSelect=$true
    $tvSrc.ImageList     = $imgList; $tvSrc.ImageIndex=0; $tvSrc.SelectedImageIndex=1

    $lblSrcStat = New-Object System.Windows.Forms.Label; $lblSrcStat.Dock='Bottom'; $lblSrcStat.Height=20
    $lblSrcStat.Font=$fSmall; $lblSrcStat.ForeColor=$cSkip; $lblSrcStat.Text='Zuerst verbinden'; $lblSrcStat.BackColor=$cPanel; $lblSrcStat.TextAlign='MiddleLeft'
    $grpSrcTree.Controls.AddRange(@($tvSrc,$lblSrcStat))
    $splitMig.Panel1.Controls.Add($grpSrcTree)

    $grpTgtTree           = New-Object System.Windows.Forms.GroupBox
    $grpTgtTree.Text      = ' Ziel-Ordner'
    $grpTgtTree.Font      = $fHead
    $grpTgtTree.ForeColor = $cHeader
    $grpTgtTree.BackColor = $cPanel
    $grpTgtTree.Dock      = 'Fill'
    $grpTgtTree.Padding   = New-Object System.Windows.Forms.Padding(6,4,6,4)

    $tvTgt               = New-Object System.Windows.Forms.TreeView
    $tvTgt.Dock          = 'Fill'; $tvTgt.Font=$fDef; $tvTgt.BackColor=$cPanel; $tvTgt.BorderStyle='FixedSingle'
    $tvTgt.HideSelection = $false; $tvTgt.ShowRootLines=$true; $tvTgt.FullRowSelect=$true
    $tvTgt.ImageList     = $imgList; $tvTgt.ImageIndex=0; $tvTgt.SelectedImageIndex=1

    $lblTgtStat = New-Object System.Windows.Forms.Label; $lblTgtStat.Dock='Bottom'; $lblTgtStat.Height=20
    $lblTgtStat.Font=$fSmall; $lblTgtStat.ForeColor=$cSkip; $lblTgtStat.Text='Zuerst verbinden'; $lblTgtStat.BackColor=$cPanel; $lblTgtStat.TextAlign='MiddleLeft'
    $grpTgtTree.Controls.AddRange(@($tvTgt,$lblTgtStat))
    $splitMig.Panel2.Controls.Add($grpTgtTree)

    # =========================================================================
    # Script-Zustand
    # =========================================================================
    $script:Connected    = $false
    $script:ApiBase      = ''
    $script:Cred         = $null
    $script:IsPBIRS      = $false
    $script:SrcApiBase   = ''
    $script:TgtApiBase   = ''
    $script:SrcCred      = $null
    $script:TgtCred      = $null
    $script:SrcConnected = $false
    $script:TgtConnected = $false
    $script:SrcFolder    = ''
    $script:TgtFolder    = ''

    # =========================================================================
    # Hilfsfunktionen
    # =========================================================================
    function Write-Log {
        param([string]$Msg, [ValidateSet('Info','Success','Warning','Error','Skip','Header')][string]$Lv='Info')
        $col = switch($Lv){'Success'{$cOk}'Warning'{$cWarn}'Error'{$cErr}'Skip'{$cSkip}'Header'{$cAccent}default{$cLogFg}}
        $pfx = switch($Lv){'Success'{'[OK]     '}'Warning'{'[WARN]   '}'Error'{'[FEHLER] '}'Skip'{'[SKIP]   '}'Header'{'========='}default{'[INFO]   '}}
        $rtb.SelectionColor = $col
        $rtb.AppendText("[$(Get-Date -Format 'HH:mm:ss')] $pfx $Msg`n")
        $rtb.ScrollToCaret()
    }

    function Write-MigLog {
        param([string]$Msg, [ValidateSet('Info','Success','Warning','Error','Skip','Header')][string]$Lv='Info')
        $col = switch($Lv){'Success'{$cOk}'Warning'{$cWarn}'Error'{$cErr}'Skip'{$cSkip}'Header'{$cAccent}default{$cLogFg}}
        $pfx = switch($Lv){'Success'{'[OK]     '}'Warning'{'[WARN]   '}'Error'{'[FEHLER] '}'Skip'{'[SKIP]   '}'Header'{'========='}default{'[INFO]   '}}
        $rtbMig.SelectionColor = $col
        $rtbMig.AppendText("[$(Get-Date -Format 'HH:mm:ss')] $pfx $Msg`n")
        $rtbMig.ScrollToCaret()
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

    function Get-MigCred {
        param([System.Windows.Forms.CheckBox]$chk, [System.Windows.Forms.TextBox]$txtU, [System.Windows.Forms.TextBox]$txtPw)
        if ($chk.Checked) {
            $u=$txtU.Text.Trim(); $p=$txtPw.Text
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

    function Update-MigBtn {
        $btnMig.Enabled = ($script:SrcConnected -and $script:TgtConnected -and
                           -not [string]::IsNullOrWhiteSpace($script:SrcFolder) -and
                           -not [string]::IsNullOrWhiteSpace($script:TgtFolder) -and
                           -not [string]::IsNullOrWhiteSpace($txtExpPath.Text))
    }

    function Load-Tree {
        $tv.Nodes.Clear(); $lblTreeStat.Text='Lade ...'
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $folders = @(Get-SSRSFolders -ApiBase $script:ApiBase -Credential $script:Cred)
            $root    = New-Object System.Windows.Forms.TreeNode('/ (Wurzel)')
            $root.Tag='/'; $root.ImageIndex=0; $root.SelectedImageIndex=1
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
            Write-Log "Ordnerstruktur geladen - $($folders.Count) Ordner" -Lv 'Success'
        } catch {
            $lblTreeStat.Text='Fehler'
            Write-Log "Fehler Ordnerbaum: $($_.Exception.Message)" -Lv 'Error'
        }
    }

    function Load-ServerTree {
        param(
            [string]$ApiBase,
            [System.Management.Automation.PSCredential]$Cred,
            [System.Windows.Forms.TreeView]$TreeView,
            [System.Windows.Forms.Label]$StatusLabel
        )
        $TreeView.Nodes.Clear(); $StatusLabel.Text='Lade ...'
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $folders = @(Get-SSRSFolders -ApiBase $ApiBase -Credential $Cred)
            $root    = New-Object System.Windows.Forms.TreeNode('/ (Wurzel)')
            $root.Tag='/'; $root.ImageIndex=0; $root.SelectedImageIndex=1
            [void]$TreeView.Nodes.Add($root)
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
            $StatusLabel.Text="$($folders.Count) Ordner"
        } catch {
            $StatusLabel.Text='Fehler'
            Write-MigLog "Fehler Ordnerbaum ($($TreeView.Name)): $($_.Exception.Message)" -Lv 'Error'
        }
    }

    # =========================================================================
    # Events - Deploy Tab
    # =========================================================================

    $form.Add_Shown({
        [System.Windows.Forms.Application]::DoEvents()
        $splitH.SplitterDistance   = 270
        $splitV.SplitterDistance   = [int]($splitV.Height * 0.45)
        $splitMig.SplitterDistance = [int]($splitMig.Width * 0.5)

        Write-Log 'dtcSoftware Report Deployment Tool v4.0 bereit.' -Lv 'Header'
        Write-Log "Benutzer: $env:USERDOMAIN\$env:USERNAME" -Lv 'Info'
        Write-Log 'Suche lokalen Report Server ...' -Lv 'Info'
        [System.Windows.Forms.Application]::DoEvents()
        $found = Find-LocalSSRS
        if ($found) {
            $txtSrv.Text      = $found
            $txtSrv.ForeColor = [System.Drawing.Color]::Black
            Write-Log "Lokaler Report Server gefunden: $found" -Lv 'Success'
            Write-Log 'URL vorbelegt - bitte Verbinden klicken.' -Lv 'Info'
        } else {
            $txtSrv.Text      = ''
            $txtSrv.ForeColor = [System.Drawing.Color]::Black
            Write-Log 'Kein lokaler Report Server gefunden - URL manuell eingeben.' -Lv 'Warning'
        }
        Write-Log '1. Server-URL  2. Verbinden  3. Ordner  4. Scannen  5. Deployen' -Lv 'Info'
    })

    $btnConn.Add_Click({
        $url=$txtSrv.Text.Trim()
        if([string]::IsNullOrWhiteSpace($url)){[System.Windows.Forms.MessageBox]::Show('Bitte Report Server URL eingeben.','Fehlt','OK','Warning')|Out-Null; return}
        $btnConn.Enabled=$false; $btnConn.Text='Verbinde ...'
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $script:Cred = Get-Cred
            if($chkCrd.Checked -and $null -eq $script:Cred){return}
            $altUrl = if($url -match '^https://'){$url -replace '^https://','http://'} else {$url -replace '^http://','https://'}
            $usedUrl=$url; $lastError=$null; $info=$null
            foreach ($c in @($url,$altUrl)) {
                try { $info=Test-SSRSConnection -ServerUrl $c -Credential $script:Cred; $usedUrl=$c; $lastError=$null; break }
                catch { $lastError=$_; $alt=if($c -match '^https://'){'http'} else {'https'}; Write-Log "[$($c.Split('//')[0].TrimEnd(':'))] fehlgeschlagen - versuche $alt ..." -Lv 'Warning' }
            }
            if($null -ne $lastError){throw $lastError}
            if($usedUrl -ne $url){$txtSrv.Text=$usedUrl; Write-Log "URL korrigiert auf $usedUrl" -Lv 'Warning'}

            $pn = if($info.PSObject.Properties['ProductName']){$info.ProductName} else {'SSRS'}
            $pv = if($info.PSObject.Properties['ProductVersion']){$info.ProductVersion} else {''}
            $script:ApiBase   = Get-SSRSApiBase $usedUrl
            $script:Connected = $true

            if($pn -match 'Power\s*BI'){
                $script:IsPBIRS=$true
                Write-Log "Verbunden: $pn $pv [$usedUrl]" -Lv 'Success'
                Write-Log '.pbix Upload aktiviert (Power BI Report Server erkannt).' -Lv 'Success'
            } else {
                $script:IsPBIRS=$false
                Write-Log "Verbunden: $pn $pv [$usedUrl]" -Lv 'Success'
                Write-Log '.pbix nur auf PBIRS verfuegbar (SSRS erkannt).' -Lv 'Info'
            }
            Load-Tree; Update-DeployBtn
        } catch {
            Write-Log "Verbindungsfehler: $($_.Exception.Message)" -Lv 'Error'
            [System.Windows.Forms.MessageBox]::Show("Verbindung fehlgeschlagen:`n$($_.Exception.Message)",'Fehler','OK','Error')|Out-Null
            $script:Connected=$false
        } finally { $btnConn.Enabled=$true; $btnConn.Text='Verbinden' }
    })

    $tv.Add_AfterSelect({
        if($tv.SelectedNode -and $tv.SelectedNode.Tag){
            $txtTgt.Text=$tv.SelectedNode.Tag.ToString(); $lblHint.Text=''
            Write-Log "Zielordner: $($txtTgt.Text)" -Lv 'Info'
            Update-DeployBtn
        }
    })

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
            $txtTgt.Text=$np; $lblHint.Text=''; $lblTreeStat.Text="Angelegt: $np"
            Write-Log "Ordner angelegt: $np" -Lv 'Success'; Update-DeployBtn
        } catch {
            Write-Log "Fehler Ordner '$np': $($_.Exception.Message)" -Lv 'Error'
            [System.Windows.Forms.MessageBox]::Show("Fehler:`n$($_.Exception.Message)",'Fehler','OK','Error')|Out-Null
        }
    })

    $mnuRefresh.Add_Click({ if($script:Connected){Load-Tree} })

    $btnBrw.Add_Click({
        $fbd=New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description='Quellverzeichnis auswaehlen'; $fbd.ShowNewFolderButton=$false
        if($txtSrc.Text -and (Test-Path $txtSrc.Text)){$fbd.SelectedPath=$txtSrc.Text}
        if($fbd.ShowDialog() -eq 'OK'){$txtSrc.Text=$fbd.SelectedPath; Write-Log "Quellverzeichnis: $($fbd.SelectedPath)" -Lv 'Info'}
    })

    $btnScan.Add_Click({
        $src=$txtSrc.Text.Trim()
        if(-not(Test-Path $src -PathType Container)){[System.Windows.Forms.MessageBox]::Show("Verzeichnis nicht gefunden:`n$src",'Fehler','OK','Warning')|Out-Null; return}
        $clv.Items.Clear()
        $files=foreach($ext in @('*.rdl','*.rds','*.rsds','*.rsd','*.pbix')){Get-ChildItem -Path $src -Filter $ext -File -ErrorAction SilentlyContinue}
        if(-not $files){Write-Log "Keine Dateien in: $src" -Lv 'Warning'; $lblFCnt.Text='0 Dateien'; return}
        foreach($f in ($files|Sort-Object Name)){
            $typ=switch($f.Extension.ToLower()){'.rdl'{'Report'}'.rds'{'DSrc'}'.rsds'{'DSrc'}'.rsd'{'Dataset'}'.pbix'{'PowerBI'}default{'?'}}
            $lvi=New-Object System.Windows.Forms.ListViewItem($f.Name)
            [void]$lvi.SubItems.Add($typ)
            [void]$lvi.SubItems.Add("$('{0:N1}'-f($f.Length/1KB)) KB")
            [void]$lvi.SubItems.Add($f.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))
            [void]$lvi.SubItems.Add('')
            $lvi.Checked=$true; $lvi.Tag=$f.FullName
            if($typ -eq 'PowerBI' -and -not $script:IsPBIRS){ $lvi.ForeColor=$cSkip }
            [void]$clv.Items.Add($lvi)
        }
        $n=$clv.Items.Count; $lblFCnt.Text="$n Datei(en)"
        Write-Log "Scan: $n Datei(en) - $src" -Lv 'Info'
        Update-DeployBtn
    })

    $btnAll.Add_Click({  foreach($i in $clv.Items){$i.Checked=$true}  })
    $btnNone.Add_Click({ foreach($i in $clv.Items){$i.Checked=$false} })

    $btnClrLog.Add_Click({ $rtb.Clear(); $lblSum.Text='' })
    $btnSavLog.Add_Click({
        $sfd=New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter='Textdatei (*.txt)|*.txt|Alle (*.*)|*.*'
        $sfd.FileName="SSRS_Deploy_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        if($sfd.ShowDialog() -eq 'OK'){$rtb.Text|Out-File -FilePath $sfd.FileName -Encoding UTF8; Write-Log "Log gespeichert: $($sfd.FileName)" -Lv 'Info'}
    })

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

        Write-Log "=== Deploy nach $tgt ($($sel.Count) Datei(en)) ===" -Lv 'Header'
        try { New-SSRSFolderRecursive -ApiBase $script:ApiBase -FolderPath $tgt -Credential $cred }
        catch { Write-Log "Fehler Zielordner: $($_.Exception.Message)" -Lv 'Error'; $btnDeploy.Enabled=$true; $btnScan.Enabled=$true; $btnConn.Enabled=$true; return }

        foreach($item in $sel){
            $fp=$item.Tag; $fn=$item.SubItems[0].Text; $typ=$item.SubItems[1].Text
            try {
                $res=switch($typ){
                    'Report'  {Deploy-Report        -ApiBase $script:ApiBase -FilePath $fp -TargetFolder $tgt -Credential $cred}
                    'DSrc'    {Deploy-DataSource    -ApiBase $script:ApiBase -FilePath $fp -TargetFolder $tgt -Credential $cred}
                    'Dataset' {Deploy-SharedDataset -ApiBase $script:ApiBase -FilePath $fp -TargetFolder $tgt -Credential $cred}
                    'PowerBI' {
                        if($script:IsPBIRS){
                            Deploy-PowerBIReport -ApiBase $script:ApiBase -FilePath $fp -TargetFolder $tgt -Credential $cred
                        } else {
                            'SKIP - .pbix nur auf Power BI Report Server (PBIRS) verfuegbar'
                        }
                    }
                    default   {'UNBEKANNTER TYP'}
                }
                if($res -like 'SKIP*' -or $res -like 'SKIPPED*'){
                    Write-Log "$fn - $res" -Lv 'Skip'; $item.SubItems[4].Text=$res; $item.ForeColor=$cSkip; $sk++
                } else {
                    Write-Log "$fn - $res" -Lv 'Success'; $item.SubItems[4].Text=$res; $item.ForeColor=$cOk; $ok++
                }
            } catch {
                $em=$_.Exception.Message
                Write-Log "$fn - FEHLER: $em" -Lv 'Error'; $item.SubItems[4].Text="FEHLER: $em"; $item.ForeColor=$cErr; $er++
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

    # =========================================================================
    # Events - Migration Tab
    # =========================================================================

    $btnSrcConn.Add_Click({
        $url=$txtSrcSrv.Text.Trim()
        if([string]::IsNullOrWhiteSpace($url)){[System.Windows.Forms.MessageBox]::Show('Quell-Server URL eingeben.','Fehlt','OK','Warning')|Out-Null; return}
        $btnSrcConn.Enabled=$false; $btnSrcConn.Text='Verbinde ...'
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $cred=Get-MigCred $chkSrcCrd $txtSrcUser $txtSrcPwd
            if($chkSrcCrd.Checked -and $null -eq $cred){return}
            $altUrl=if($url -match '^https://'){$url -replace '^https://','http://'} else {$url -replace '^http://','https://'}
            $usedUrl=$url; $lastErr=$null; $info=$null
            foreach($c in @($url,$altUrl)){
                try{$info=Test-SSRSConnection -ServerUrl $c -Credential $cred; $usedUrl=$c; $lastErr=$null; break}
                catch{$lastErr=$_; Write-MigLog "[$($c.Split('//')[0].TrimEnd(':'))] fehlgeschlagen" -Lv 'Warning'}
            }
            if($null -ne $lastErr){throw $lastErr}
            $script:SrcApiBase  =Get-SSRSApiBase $usedUrl
            $script:SrcCred     =$cred
            $script:SrcConnected=$true
            $pn=if($info.PSObject.Properties['ProductName']){$info.ProductName} else {'SSRS'}
            Write-MigLog "Quell-Server verbunden: $pn [$usedUrl]" -Lv 'Success'
            Load-ServerTree -ApiBase $script:SrcApiBase -Cred $script:SrcCred -TreeView $tvSrc -StatusLabel $lblSrcStat
            Update-MigBtn
        } catch {
            Write-MigLog "Quell-Server Fehler: $($_.Exception.Message)" -Lv 'Error'
            $script:SrcConnected=$false
        } finally {$btnSrcConn.Enabled=$true; $btnSrcConn.Text='Verbinden'}
    })

    $btnTgtConn.Add_Click({
        $url=$txtTgtSrv.Text.Trim()
        if([string]::IsNullOrWhiteSpace($url)){[System.Windows.Forms.MessageBox]::Show('Ziel-Server URL eingeben.','Fehlt','OK','Warning')|Out-Null; return}
        $btnTgtConn.Enabled=$false; $btnTgtConn.Text='Verbinde ...'
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $cred=Get-MigCred $chkTgtCrd $txtTgtUser $txtTgtPwd
            if($chkTgtCrd.Checked -and $null -eq $cred){return}
            $altUrl=if($url -match '^https://'){$url -replace '^https://','http://'} else {$url -replace '^http://','https://'}
            $usedUrl=$url; $lastErr=$null; $info=$null
            foreach($c in @($url,$altUrl)){
                try{$info=Test-SSRSConnection -ServerUrl $c -Credential $cred; $usedUrl=$c; $lastErr=$null; break}
                catch{$lastErr=$_; Write-MigLog "[$($c.Split('//')[0].TrimEnd(':'))] fehlgeschlagen" -Lv 'Warning'}
            }
            if($null -ne $lastErr){throw $lastErr}
            $script:TgtApiBase  =Get-SSRSApiBase $usedUrl
            $script:TgtCred     =$cred
            $script:TgtConnected=$true
            $pn=if($info.PSObject.Properties['ProductName']){$info.ProductName} else {'SSRS'}
            Write-MigLog "Ziel-Server verbunden: $pn [$usedUrl]" -Lv 'Success'
            Load-ServerTree -ApiBase $script:TgtApiBase -Cred $script:TgtCred -TreeView $tvTgt -StatusLabel $lblTgtStat
            Update-MigBtn
        } catch {
            Write-MigLog "Ziel-Server Fehler: $($_.Exception.Message)" -Lv 'Error'
            $script:TgtConnected=$false
        } finally {$btnTgtConn.Enabled=$true; $btnTgtConn.Text='Verbinden'}
    })

    $tvSrc.Add_AfterSelect({
        if($tvSrc.SelectedNode -and $tvSrc.SelectedNode.Tag){
            $script:SrcFolder=$tvSrc.SelectedNode.Tag.ToString()
            Write-MigLog "Quell-Ordner: $($script:SrcFolder)" -Lv 'Info'
            Update-MigBtn
        }
    })

    $tvTgt.Add_AfterSelect({
        if($tvTgt.SelectedNode -and $tvTgt.SelectedNode.Tag){
            $script:TgtFolder=$tvTgt.SelectedNode.Tag.ToString()
            Write-MigLog "Ziel-Ordner: $($script:TgtFolder)" -Lv 'Info'
            Update-MigBtn
        }
    })

    $txtExpPath.Add_TextChanged({ Update-MigBtn })

    $btnExpBrw.Add_Click({
        $fbd=New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description='Zwischenverzeichnis fuer Export auswaehlen'; $fbd.ShowNewFolderButton=$true
        if($txtExpPath.Text -and (Test-Path $txtExpPath.Text)){$fbd.SelectedPath=$txtExpPath.Text}
        if($fbd.ShowDialog() -eq 'OK'){
            $txtExpPath.Text=$fbd.SelectedPath
            Write-MigLog "Export-Pfad: $($fbd.SelectedPath)" -Lv 'Info'
            Update-MigBtn
        }
    })

    $chkDsMap.Add_CheckedChanged({
        $pnlDsWrap.Height  = if($chkDsMap.Checked){120} else {0}
        $pnlDsWrap.Visible = $chkDsMap.Checked
    })

    $btnMigClrLog.Add_Click({ $rtbMig.Clear(); $lblMigStat.Text='' })
    $btnMigSavLog.Add_Click({
        $sfd=New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter='Textdatei (*.txt)|*.txt|Alle (*.*)|*.*'
        $sfd.FileName="SSRS_Mig_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        if($sfd.ShowDialog() -eq 'OK'){$rtbMig.Text|Out-File -FilePath $sfd.FileName -Encoding UTF8; Write-MigLog "Log gespeichert: $($sfd.FileName)" -Lv 'Info'}
    })

    $btnMig.Add_Click({
        if([string]::IsNullOrWhiteSpace($script:SrcFolder)){[System.Windows.Forms.MessageBox]::Show('Quell-Ordner auswaehlen.','Fehler','OK','Warning')|Out-Null; return}
        if([string]::IsNullOrWhiteSpace($script:TgtFolder)){[System.Windows.Forms.MessageBox]::Show('Ziel-Ordner auswaehlen.','Fehler','OK','Warning')|Out-Null; return}
        $expPath=$txtExpPath.Text.Trim()
        if([string]::IsNullOrWhiteSpace($expPath)){[System.Windows.Forms.MessageBox]::Show('Zwischenverzeichnis angeben.','Fehler','OK','Warning')|Out-Null; return}

        $dsMap=@{}
        if($chkDsMap.Checked){
            foreach($row in $dgvDs.Rows){
                if($row.IsNewRow){continue}
                $s=if($row.Cells['ColSrc'].Value){$row.Cells['ColSrc'].Value.ToString().Trim()} else {''}
                $t=if($row.Cells['ColTgt'].Value){$row.Cells['ColTgt'].Value.ToString().Trim()} else {''}
                if(-not [string]::IsNullOrWhiteSpace($s)){$dsMap[$s]=$t}
            }
            Write-MigLog "DS-Mapping: $($dsMap.Count) Eintraege" -Lv 'Info'
        }

        $btnMig.Enabled=$false; $btnSrcConn.Enabled=$false; $btnTgtConn.Enabled=$false
        $pbMig.Value=0; $pbMig.Maximum=100; $pbMig.Style='Marquee'

        Write-MigLog "========= Migration Start: $($script:SrcFolder) nach $($script:TgtFolder) =========" -Lv 'Header'
        Write-MigLog "Quelle: $($script:SrcApiBase)" -Lv 'Info'
        Write-MigLog "Ziel:   $($script:TgtApiBase)" -Lv 'Info'
        Write-MigLog "Export: $expPath" -Lv 'Info'
        [System.Windows.Forms.Application]::DoEvents()

        try {
            Write-MigLog "========= Phase 1 - Export von Quell-Server =========" -Lv 'Header'
            $result=Export-SSRSContent -ApiBase $script:SrcApiBase -SourceFolderPath $script:SrcFolder `
                        -ExportPath $expPath -DsMappings $dsMap -Credential $script:SrcCred
            foreach($msg in $result.Messages){
                Write-MigLog $msg -Lv $(if($msg -like '[ERR]*'){'Error'} else {'Success'})
            }
            Write-MigLog "Export: $($result.Exported) Dateien, $($result.Errors) Fehler" -Lv $(if($result.Errors-gt 0){'Warning'} else {'Success'})

            Write-MigLog "========= Phase 2 - Deploy auf Ziel-Server =========" -Lv 'Header'
            New-SSRSFolderRecursive -ApiBase $script:TgtApiBase -FolderPath $script:TgtFolder -Credential $script:TgtCred

            $depFiles=@(Get-ChildItem -Path $expPath -Recurse -File -Include '*.rdl','*.rds','*.rsds','*.rsd' -ErrorAction SilentlyContinue)
            $pbMig.Style='Continuous'; $pbMig.Maximum=[math]::Max(1,$depFiles.Count); $pbMig.Value=0
            $dOk=0; $dSk=0; $dErr=0

            foreach($f in $depFiles){
                $relDir=$f.DirectoryName.Substring($expPath.TrimEnd('\').Length).TrimStart('\','/')
                $tgtSub=if($relDir){"$($script:TgtFolder)/$($relDir -replace '\\','/')"} else {$script:TgtFolder}
                New-SSRSFolderRecursive -ApiBase $script:TgtApiBase -FolderPath $tgtSub -Credential $script:TgtCred
                try {
                    $res=switch($f.Extension.ToLower()){
                        '.rdl' {Deploy-Report        -ApiBase $script:TgtApiBase -FilePath $f.FullName -TargetFolder $tgtSub -Credential $script:TgtCred}
                        '.rsd' {Deploy-SharedDataset -ApiBase $script:TgtApiBase -FilePath $f.FullName -TargetFolder $tgtSub -Credential $script:TgtCred}
                        default{Deploy-DataSource    -ApiBase $script:TgtApiBase -FilePath $f.FullName -TargetFolder $tgtSub -Credential $script:TgtCred}
                    }
                    if($res -like 'SKIP*' -or $res -like 'SKIPPED*'){
                        Write-MigLog "$($f.Name) - $res" -Lv 'Skip'; $dSk++
                    } else {
                        Write-MigLog "$($f.Name) - $res" -Lv 'Success'; $dOk++
                    }
                } catch {
                    Write-MigLog "$($f.Name) - FEHLER: $($_.Exception.Message)" -Lv 'Error'; $dErr++
                }
                $pbMig.Value++
                [System.Windows.Forms.Application]::DoEvents()
            }

            $sumMig="Migration: Export $($result.Exported) | Deploy $dOk OK, $dSk Skip, $dErr Fehler"
            Write-MigLog $sumMig -Lv $(if($dErr-gt 0){'Warning'} else {'Header'})
            $lblMigStat.Text=$sumMig; $lblMigStat.ForeColor=if($dErr-gt 0){$cErr} elseif($dSk-gt 0){$cWarn} else{$cOk}
            Load-ServerTree -ApiBase $script:TgtApiBase -Cred $script:TgtCred -TreeView $tvTgt -StatusLabel $lblTgtStat
        } catch {
            Write-MigLog "Schwerwiegender Fehler: $($_.Exception.Message)" -Lv 'Error'
        } finally {
            $btnMig.Enabled=$true; $btnSrcConn.Enabled=$true; $btnTgtConn.Enabled=$true
            $pbMig.Style='Continuous'
            Update-MigBtn
        }
    })

    [void]$form.ShowDialog()
    $form.Dispose()
}

# =============================================================================
# Einstiegspunkt
# =============================================================================
Show-DeploymentTool
