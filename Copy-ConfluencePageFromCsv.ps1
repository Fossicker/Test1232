<#
.SYNOPSIS
    Kopiert eine Confluence-Seite (Vorlage) pro CSV-Zeile und setzt den Titel aus einer CSV-Spalte.

.DESCRIPTION
    - Liest die Vorlagenseite (Storage-Format) per REST-API.
    - Legt pro CSV-Zeile eine neue Seite unter der angegebenen Elternseite an.
    - Titel = Wert der Spalte -TitleColumn.
    - Optional: Platzhalter {{Spaltenname}} im Seiteninhalt werden durch die CSV-Werte ersetzt.
    - Labels der Vorlage werden mitkopiert (-CopyLabels).
    - Existiert eine Seite mit gleichem Titel im Space bereits, wird die Zeile übersprungen.
    - Unterstützt -WhatIf (Probelauf ohne Anlage).
    Kompatibel mit Windows PowerShell 5.1 und PowerShell 7.

.EXAMPLE
    .\Copy-ConfluencePageFromCsv.ps1 -BaseUrl "https://confluence.example.de" `
        -TemplatePageId 123456 -ParentPageId 654321 `
        -CsvPath .\seiten.csv -TitleColumn "Titel" -Token $env:CONFLUENCE_PAT -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)] [string] $BaseUrl,          # z.B. https://confluence.example.de (ohne /rest)
    [Parameter(Mandatory)] [string] $TemplatePageId,   # ID der zu kopierenden Seite
    [string] $ParentPageId,                            # Elternseite der neuen Seiten (Default: Elternseite der Vorlage)
    [string] $SpaceKey,                                # Ziel-Space (Default: Space der Vorlage)
    [Parameter(Mandatory)] [string] $CsvPath,
    [string] $TitleColumn = 'Titel',
    [string] $Delimiter = ';',
    [string] $Token,                                   # Personal Access Token; wenn leer -> Abfrage
    [switch] $ReplacePlaceholders,                     # {{Spalte}} im Inhalt ersetzen
    [switch] $CopyLabels
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$BaseUrl = $BaseUrl.TrimEnd('/')

if (-not $Token) {
    $sec   = Read-Host 'Personal Access Token' -AsSecureString
    $Token = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec))
}
$headers = @{ Authorization = "Bearer $Token"; Accept = 'application/json' }

function Invoke-Confluence {
    param([string]$Method = 'GET', [string]$Path, $Body)
    $params = @{ Method = $Method; Uri = "$BaseUrl$Path"; Headers = $headers }
    if ($null -ne $Body) {
        # UTF-8 explizit als Bytes senden (PS 5.1 würde sonst Umlaute zerstören)
        $json = $Body | ConvertTo-Json -Depth 20 -Compress
        $params.Body        = [Text.Encoding]::UTF8.GetBytes($json)
        $params.ContentType = 'application/json; charset=utf-8'
    }
    Invoke-RestMethod @params
}

function ConvertTo-XmlText([string]$s) {
    # Werte für den Einsatz im Storage-Format (XHTML) escapen
    [Security.SecurityElement]::Escape($s)
}

# --- Vorlage laden ---------------------------------------------------------
Write-Host "Lade Vorlage $TemplatePageId ..." -ForegroundColor Cyan
$template = Invoke-Confluence -Path "/rest/api/content/$TemplatePageId`?expand=body.storage,space,ancestors"

if (-not $SpaceKey)     { $SpaceKey = $template.space.key }
if (-not $ParentPageId) {
    if ($template.ancestors.Count -gt 0) { $ParentPageId = $template.ancestors[-1].id }
}
$templateBody = $template.body.storage.value

$labels = @()
if ($CopyLabels) {
    $labels = (Invoke-Confluence -Path "/rest/api/content/$TemplatePageId/label").results |
              ForEach-Object { @{ prefix = $_.prefix; name = $_.name } }
}

Write-Host "Vorlage: '$($template.title)' | Space: $SpaceKey | Eltern-ID: $ParentPageId" -ForegroundColor Cyan

# --- CSV lesen -------------------------------------------------------------
$rows = Import-Csv -Path $CsvPath -Delimiter $Delimiter -Encoding UTF8
if (-not $rows) { throw "CSV '$CsvPath' ist leer." }
if ($TitleColumn -notin $rows[0].PSObject.Properties.Name) {
    throw "Spalte '$TitleColumn' nicht in CSV gefunden. Vorhanden: $($rows[0].PSObject.Properties.Name -join ', ')"
}

$result = foreach ($row in $rows) {
    $title = "$($row.$TitleColumn)".Trim()
    if (-not $title) { Write-Warning 'Leerer Titel – Zeile übersprungen.'; continue }

    # Existiert der Titel schon im Space?
    $q = "/rest/api/content?spaceKey=$([uri]::EscapeDataString($SpaceKey))&title=$([uri]::EscapeDataString($title))&type=page"
    if ((Invoke-Confluence -Path $q).size -gt 0) {
        Write-Warning "Existiert bereits: '$title' – übersprungen."
        [pscustomobject]@{ Titel = $title; Status = 'Existiert'; Id = $null; Url = $null }
        continue
    }

    $body = $templateBody
    if ($ReplacePlaceholders) {
        foreach ($p in $row.PSObject.Properties) {
            $body = $body.Replace("{{$($p.Name)}}", (ConvertTo-XmlText "$($p.Value)"))
        }
    }

    $newPage = @{
        type  = 'page'
        title = $title
        space = @{ key = $SpaceKey }
        body  = @{ storage = @{ value = $body; representation = 'storage' } }
    }
    if ($ParentPageId) { $newPage.ancestors = @(@{ id = $ParentPageId }) }

    if ($PSCmdlet.ShouldProcess($title, 'Confluence-Seite anlegen')) {
        try {
            $created = Invoke-Confluence -Method POST -Path '/rest/api/content' -Body $newPage
            if ($labels) {
                Invoke-Confluence -Method POST -Path "/rest/api/content/$($created.id)/label" -Body $labels | Out-Null
            }
            $url = "$BaseUrl$($created._links.webui)"
            Write-Host "Angelegt: '$title' -> $url" -ForegroundColor Green
            [pscustomobject]@{ Titel = $title; Status = 'Angelegt'; Id = $created.id; Url = $url }
        }
        catch {
            Write-Warning "Fehler bei '$title': $($_.Exception.Message)"
            [pscustomobject]@{ Titel = $title; Status = "Fehler: $($_.Exception.Message)"; Id = $null; Url = $null }
        }
    }
}

# --- Protokoll -------------------------------------------------------------
if ($result) {
    $log = Join-Path (Split-Path $CsvPath -Parent) ("ConfluenceCopy_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date))
    $result | Export-Csv -Path $log -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8
    $result | Format-Table -AutoSize
    Write-Host "Protokoll: $log" -ForegroundColor Cyan
}
