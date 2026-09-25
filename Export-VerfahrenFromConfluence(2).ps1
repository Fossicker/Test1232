<#
    Export-VerfahrenFromConfluence.ps1

    Liest alle (Unter-)Seiten einer Confluence-Data-Center-Seite aus und exportiert:

      1. die Werte "Name des Verfahrens" und "Kurze Beschreibung des Verfahrens"
         aus der Beschriftungstabelle (Beschriftung Spalte 1, Wert Spalte 2)
      1b. die Tabellen "Allgemeine Information", "Hersteller und Support",
          "Verfuegbarkeitsanforderungen" und "Architektur und Betrieb" -
          gleicher Aufbau wie Tabelle 1, alle Werte landen in der Haupt-CSV
      2. die Matrix-Tabelle "Anwendungsebene | IST | SOLL | Anzahl IST-Server"
         mit einer Zeile je Anwendungsebene. Die Zeilen unterhalb von
         "Zusaetzlich benoetigte Anwendungsebenen (ausserhalb des Standards)"
         gehen in eine eigene CSV, da ihre Anzahl je Seite variiert.
      3. die Schnittstellen-Tabelle "Input/Output | Ziel- bzw. Quellsystem |
         Manuell/Halbautom./Autom. | Kurzbeschreibung zu uebergebende Daten"
         mit beliebig vielen Zeilen je Seite
      4. den Sperrzeiten-Block der Tabelle "Migration": unter der Zeile
         "Sperrzeiten" die Spalten von | bis | Bemerkung/Name
      5. die Tabelle "Kapazitaetsanforderungen Server": Servername | Cluster |
         Anwendungsebene | Betriebssystem | CPU | RAM | Festplatten | Art
      6. die Tabelle "Kapazitaetsanforderungen Datenbank":
         Datenbank | Anwendungsebene | Groesse | Plattform
      7. die Tabelle "Aufbau Anwendungsserver": in Bloecke gegliedert
         (Appliance, Betriebssystem, Build-In Datenbank, Middleware,
         Zusaetzliche Anwendungssoftware) - die Kategorie landet in der
         Spalte "Aufbau"
      8. die Tabelle "Umsysteme" - gleicher Aufbau wie Tabelle 7, Kategorien
         Appliance, Datenbank, Netzwerk Dienste, Sonstige; die Kategorie
         landet in der Spalte "Umsysteme"

    Ausgabe:
      - Verfahren.csv                 eine Zeile je Seite (Tabelle 1 + 2)
      - Verfahren-Schnittstellen.csv  eine Zeile je Schnittstelle (Tabelle 3)
      - Verfahren-Zusatzebenen.csv    eine Zeile je zusaetzlicher Anwendungsebene
      - Verfahren-Sperrzeiten.csv     eine Zeile je Sperrzeit (Tabelle 4)
      - Verfahren-Server.csv          eine Zeile je Server (Tabelle 5)
      - Verfahren-Datenbanken.csv     eine Zeile je Datenbank (Tabelle 6)
      - Verfahren-Aufbau.csv          eine Zeile je Aufbau-Eintrag (Tabelle 7)
      - Verfahren-Umsysteme.csv       eine Zeile je Umsystem (Tabelle 8)
    Erste Spalte ist jeweils der Seitenname.

    Nur Bordmittel (PowerShell 5.1 / 7.x), keine Module, kein NuGet.
    Keine Parameteruebergabe - alles wird im Block "KONFIGURATION" eingestellt.

    Ausfuehren:  .\Export-VerfahrenFromConfluence.ps1
#>

# ===========================================================================
#  KONFIGURATION - ab hier anpassen
# ===========================================================================

# Basis-URL der Confluence-Instanz, ohne abschliessenden Slash.
# Kontextpfad mit angeben, falls vorhanden (z. B. .../confluence).
$BaseUrl = 'https://confluence.intern'

# ID der uebergeordneten Seite ("Dokument"), unterhalb derer gesucht wird.
$ParentPageId = '123456789'

# Alternative zur ID: Space-Key + exakter Seitentitel.
# Wird nur ausgewertet, wenn $ParentPageId leer ist ('').
$SpaceKey    = ''
$ParentTitle = ''

# --- Anmeldung -------------------------------------------------------------
# 'Token'     = Personal Access Token (Bearer)      -> $Token setzen
# 'Basic'     = Benutzername/Passwort               -> Abfrage per Get-Credential
# 'SSO'       = Windows-Anmeldung (Kerberos/NTLM)   -> keine weitere Angabe noetig
# 'FormLogin' = Formular-Login /dologin.action      -> Abfrage per Get-Credential
# 'Anonymous' = ohne Anmeldung
$AuthMode = 'Token'

# Personal Access Token. Besser nicht im Klartext ins Skript schreiben, sondern
# einmalig als Umgebungsvariable setzen:
#   setx CONFLUENCE_PAT "MeinToken"      (danach neue PowerShell-Sitzung oeffnen)
$Token = $env:CONFLUENCE_PAT
# Notloesung fuer einen einmaligen Lauf:
# $Token = 'NDkyMjM4...'

# --- Ausgabe ---------------------------------------------------------------
# Eine Zeile je Seite, jeder Wert in einer eigenen Spalte:
#   Seitenname | Verfahren | Beschreibung | ENTW_IST | ENTW_SOLL | ENTW_Anzahl | ...
$OutputPath = '.\Verfahren.csv'
$Delimiter  = ';'      # ';' passt zum deutschen Excel

# $true = SeitenId und Url zusaetzlich als letzte Spalten ausgeben
$IncludeSeitenIdUndUrl = $true

# --- Tabelle 1: Beschriftung links, Wert rechts ----------------------------
# Der Vergleich ignoriert Gross-/Kleinschreibung, Doppelpunkte und Formatierung.
$NameLabel        = 'Name des Verfahrens'
$DescriptionLabel = 'Kurze Beschreibung des Verfahrens'

# $true = falls die Beschriftungen nicht gefunden werden, stattdessen stur
#         Zeile 1 / Spalte 2 bzw. Zeile 2 / Spalte 2 uebernehmen.
$UsePositionFallback = $false

# --- Tabelle "Allgemeine Information" und weitere Beschriftungszeilen ------
# Gleicher Aufbau wie Tabelle 1: Beschriftung links, Wert rechts. Die Zeilen
# duerfen in einer beliebigen Tabelle der Seite stehen.
# Links der Spaltenname in der CSV, rechts die Beschriftung auf der Seite.
# Weitere Zeilen einfach hier ergaenzen - der Rest des Skripts passt sich an.
#
# Die Beschriftung muss nicht vollstaendig sein, ein eindeutiger Teil reicht.
# Umlaute duerfen als ae/oe/ue/ss geschrieben werden, der Vergleich gleicht das
# aus. Mehrere Schreibweisen als Liste angeben: @('Variante 1', 'Variante 2')
# - dann wird der Reihe nach gesucht, bis ein Wert gefunden ist.
$WeitereFelder = [ordered]@{
    # Allgemeine Information
    AnzahlNutzer              = 'Anzahl der Nutzer'
    Nutzerkreis               = 'Nutzerkreis'
    Fernwartungszugriff       = 'Fernwartungszugriff'
    # Hersteller und Support
    EingesetzteSoftware       = 'Eingesetzte Software'
    EingesetzteVersion        = 'Eingesetzte Version'
    LaufzeitHerstellersupport = 'Laufzeit Herstellersupport'
    # Verfuegbarkeitsanforderungen
    Kritikalitaet             = 'Kritikalitaet'
    Supportzeit               = 'Supportzeit'
    RedundanzBedarf           = @('Grad der Redundanz wird', 'Grad der Redunanz wird')
    RedundanzAktuell          = @('Redundanz hat das Verfahren', 'Redunanz hat das Verfahren')
    Aufbewahrungszeit         = 'Aufbewahrungszeit'
    # Architektur und Betrieb
    Architektur               = 'Architektur'
    ClientZugriff             = 'Greifen Clients'
    StandardOderEigenentw     = 'Standardsoftware oder Eigenentwicklung'
    GradIndividualisierung    = 'Grad der Individualisierung'
    Benutzerverwaltung        = 'Art der Benutzerverwaltung'
    StorageAnforderungen      = 'Anforderungen an den Storage'
    BetriebshandbuchAnwendung = 'Betriebshandbuch (Anwendung)'
    BetriebshandbuchInfra     = 'Betriebshandbuch (Infrastruktur)'
}

# --- Tabelle 2: Matrix der Anwendungsebenen --------------------------------
# Text in der ersten Zelle der Kopfzeile - daran wird die Matrix erkannt.
$MatrixHeaderLabel = 'Anwendungsebene'

# Spaltenreihenfolge der breiten Ausgabe. Kuerzel = Text in der ersten Klammer
# der Zeilenbeschriftung, z. B. "Entwicklung (ENTW)" -> ENTW.
# Auf den Seiten gefundene, hier nicht gelistete Ebenen werden hinten angehaengt.
$Anwendungsebenen = @(
    'ENTW'      # Entwicklung
    'YSIT'      # Subsystem Integrationstest
    'ITGN'      # Integrationstest
    'TEST'      # Freigabetest
    'PNT'       # Produktionsnaher Test
    'PROD'      # Produktion
    'SCHL'      # Produktive Schulungsumgebung
    'TESTS'     # QS Schulungsumgebung
    'PRODVOC'   # Voice Netz
    'SERVICES'  # Zentrale Anwendungsebene
)

# Zeile innerhalb der Matrix, ab der die frei erfassten Ebenen beginnen.
# Alles darunter landet in einer eigenen CSV statt in neuen Spalten der
# Haupt-CSV - sonst wuerde jede Sondereintragung eine Spalte fuer ALLE Seiten
# erzeugen. Ein eindeutiger Teil des Textes reicht, Umlaute als ae/oe/ue/ss.
$MatrixZusatzLabel      = 'Zusaetzlich benoetigte Anwendungsebenen'
$ExportZusatzEbenen     = $true
$OutputPathZusatzEbenen = '.\Verfahren-Zusatzebenen.csv'

# $true = Ja/J/X/Yes werden zu "Ja", Nein/N/No zu "Nein" vereinheitlicht.
$NormalizeJaNein = $true

# --- Tabelle 3: Schnittstellen zu anderen Anwendungen & Systemen -----------
# Da die Zeilenzahl je Seite variiert, landen die Schnittstellen in einer
# eigenen CSV: eine Zeile je Schnittstelle, verknuepft ueber den Seitennamen.
# Erkannt wird die Tabelle an ihrer Kopfzeile - mindestens zwei der vier
# Spalten muessen gefunden werden. Eine Titelzeile darueber wird uebersprungen.
$ExportSchnittstellen     = $true
$OutputPathSchnittstellen = '.\Verfahren-Schnittstellen.csv'

# --- Tabelle 4: Sperrzeiten (Tabelle "Migration") --------------------------
# Unterhalb der Zeile "Sperrzeiten" folgt eine Kopfzeile mit den Spalten
# von | bis | Bemerkung/Name und darunter beliebig viele Zeitraeume.
# Auch hier variiert die Zeilenzahl je Seite -> eigene CSV.
# Die Zeile "Sperrzeiten" ist der Anker; fehlt sie, wird ersatzweise eine
# Kopfzeile gesucht, die sowohl "von" als auch "bis" enthaelt.
$ExportSperrzeiten     = $true
$OutputPathSperrzeiten = '.\Verfahren-Sperrzeiten.csv'
$SperrzeitenLabel      = 'Sperrzeiten'

# --- Tabelle 5: Kapazitaetsanforderungen Server ----------------------------
# Eine Zeile je Server -> eigene CSV. Erkannt wird die Tabelle an ihrer
# Kopfzeile; mindestens drei der acht Spalten muessen gefunden werden, damit
# sie nicht mit der Anwendungsebenen-Matrix verwechselt wird.
$ExportServer     = $true
$OutputPathServer = '.\Verfahren-Server.csv'

# --- Tabelle 6: Kapazitaetsanforderungen Datenbank -------------------------
# Eine Zeile je Datenbank -> eigene CSV. Die Tabelle hat nur vier Spalten,
# deshalb gilt hier: die Spalte "Datenbank" MUSS erkannt werden (sie kommt in
# keiner anderen Tabelle vor) plus mindestens eine weitere. Die Spalte
# "Anwendungsebene" allein wuerde sonst auf Matrix und Server-Tabelle passen.
$ExportDatenbank     = $true
$OutputPathDatenbank = '.\Verfahren-Datenbanken.csv'

# --- Tabelle 7: Aufbau Anwendungsserver ------------------------------------
# Diese Tabelle ist in Bloecke gegliedert: eine Zeile nennt die Kategorie
# (z. B. "Appliance"), darunter stehen deren Datenzeilen. Die Kategorie landet
# in der Spalte "Aufbau", die Datenzeilen kommen darunter in eine eigene CSV.
# Die uebrigen Spalten werden aus der Kopfzeile der Tabelle uebernommen; fehlt
# sie, heissen sie Wert1, Wert2, ...
# Erkannt wird die Tabelle daran, dass mindestens zwei dieser Kategorien in
# der ersten Spalte stehen. Umlaute als ae/oe/ue/ss, ein Teil des Textes reicht.
$ExportAufbau     = $true
$OutputPathAufbau = '.\Verfahren-Aufbau.csv'
# Ueberschrift der Tabelle auf der Seite. Ist sie gesetzt, wird genau diese
# Tabelle ausgewertet - das ist die sicherste Zuordnung. Ein Teil des Textes
# reicht. Leer lassen, um nur ueber die Kategorien zu suchen.
$AufbauTitel      = 'Aufbau Anwendungsserver'
$AufbauKategorien = @(
    'Appliance'
    'Betriebssystem'
    'Build-In Datenbank'
    'Middleware'
    'Zusaetzliche Anwendungssoftware'
)

# --- Tabelle 8: Umsysteme (jeweils einzeln) --------------------------------
# Eigenstaendige Tabelle mit gleichem Aufbau wie Tabelle 7, aber anderen
# Kategorien. Die Kategorie landet in der Spalte "Umsysteme".
# "Appliance" kommt in beiden Tabellen vor - deshalb wird die Tabelle in
# erster Linie ueber ihre Ueberschrift zugeordnet.
$ExportUmsysteme     = $true
$OutputPathUmsysteme = '.\Verfahren-Umsysteme.csv'
$UmsystemeTitel      = 'Umsysteme'
$UmsystemeKategorien = @(
    'Appliance'
    'Datenbank'
    'Netzwerk Dienste'
    'Sonstige'
)

# Gilt fuer Tabelle 7 und 8:
# $true  = Eine Gruppenueberschrift, die NICHT in der Kategorienliste steht,
#          beendet den laufenden Block. Zeilen darunter werden uebersprungen,
#          bis wieder eine gesuchte Kategorie kommt. Verhindert, dass fremde
#          Bloecke derselben Tabelle mit exportiert werden.
# $false = Fremde Ueberschriften werden ignoriert, die zuletzt erkannte
#          Kategorie gilt weiter. Nur noetig, wenn zwischen Kategorie und
#          Daten Hinweiszeilen stehen ("bitte pro Umsystem eine Zeile").
$KategorieBlockStrikt = $true

# --- Suchumfang ------------------------------------------------------------
# $false = alle Ebenen unterhalb der Elternseite
# $true  = nur direkte Unterseiten
$DirectChildrenOnly = $false

# Anzahl Seiten je REST-Aufruf (1-100).
$PageSize = 50

# --- Netzwerk --------------------------------------------------------------
$Proxy                = ''      # z. B. 'http://proxy.intern:8080'; leer = kein Proxy
$SkipCertificateCheck = $false  # $true nur bei internen Zertifikaten ohne Vertrauenskette

# --- Diagnose --------------------------------------------------------------
$ShowVerbose = $false           # $true = ausfuehrliche Ausgabe je REST-Aufruf

# ===========================================================================
#  AB HIER NICHTS MEHR ANPASSEN
# ===========================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($ShowVerbose) { $VerbosePreference = 'Continue' }

# --- Vorbereitung: TLS, Zertifikate ---------------------------------------

if ($PSVersionTable.PSVersion.Major -lt 6) {
    [System.Net.ServicePointManager]::SecurityProtocol =
        [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.ServicePointManager]::SecurityProtocol
    if ($SkipCertificateCheck) {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    }
}

$BaseUrl = $BaseUrl.TrimEnd('/')

$headers = @{ 'Accept' = 'application/json' }

$requestArgs = @{
    Headers         = $headers
    UseBasicParsing = $true
    ErrorAction     = 'Stop'
    Method          = 'Get'
}
if ($Proxy) {
    $requestArgs['Proxy'] = $Proxy
    $requestArgs['ProxyUseDefaultCredentials'] = $true
}
if ($SkipCertificateCheck -and $PSVersionTable.PSVersion.Major -ge 6) {
    $requestArgs['SkipCertificateCheck'] = $true
}

# --- Anmeldung einrichten --------------------------------------------------

switch ($AuthMode) {

    'Token' {
        if ([string]::IsNullOrWhiteSpace($Token)) {
            throw 'Kein Token gesetzt. Variable $Token befuellen oder Umgebungsvariable CONFLUENCE_PAT anlegen.'
        }
        $headers['Authorization'] = "Bearer $Token"
    }

    'Basic' {
        $cred = Get-Credential -Message 'Confluence-Anmeldung (Basic)'
        $pair = '{0}:{1}' -f $cred.UserName, $cred.GetNetworkCredential().Password
        $b64  = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($pair))
        $headers['Authorization'] = "Basic $b64"
    }

    'SSO' {
        $requestArgs['UseDefaultCredentials'] = $true
    }

    'Anonymous' {
        Write-Verbose 'Zugriff ohne Anmeldung - es werden nur anonym lesbare Seiten geliefert.'
    }

    'FormLogin' {
        $cred      = Get-Credential -Message 'Confluence-Anmeldung (Formular-Login)'
        $session   = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $loginArgs = @{
            Uri             = "$BaseUrl/dologin.action"
            Method          = 'Post'
            WebSession      = $session
            UseBasicParsing = $true
            ErrorAction     = 'Stop'
            Body            = @{
                os_username    = $cred.UserName
                os_password    = $cred.GetNetworkCredential().Password
                login          = 'Anmelden'
                os_destination = '/index.action'
            }
        }
        if ($Proxy) {
            $loginArgs['Proxy'] = $Proxy
            $loginArgs['ProxyUseDefaultCredentials'] = $true
        }
        if ($SkipCertificateCheck -and $PSVersionTable.PSVersion.Major -ge 6) {
            $loginArgs['SkipCertificateCheck'] = $true
        }

        Write-Verbose "Formular-Login als '$($cred.UserName)' ..."
        try {
            Invoke-WebRequest @loginArgs | Out-Null
        }
        catch {
            throw "Formular-Login fehlgeschlagen: $($_.Exception.Message)"
        }

        $hasSession = $session.Cookies.GetCookies([uri]$BaseUrl) | Where-Object { $_.Name -match 'JSESSIONID|seraph' }
        if (-not $hasSession) {
            throw 'Formular-Login hat kein Session-Cookie geliefert - Zugangsdaten oder CAPTCHA-Sperre pruefen.'
        }

        $requestArgs['WebSession'] = $session
    }

    default {
        throw "Unbekannter Wert fuer `$AuthMode: '$AuthMode'. Erlaubt: Token, Basic, SSO, FormLogin, Anonymous."
    }
}

Write-Verbose "Anmeldeverfahren: $AuthMode"

# --- REST-Zugriff ----------------------------------------------------------

function Invoke-ConfluenceApi {
    # Ruft die REST-API auf und dekodiert die Antwort explizit als UTF-8.
    # (Windows PowerShell 5.1 verstuemmelt sonst Umlaute.)
    param([Parameter(Mandatory)][string]$Uri)

    Write-Verbose "GET $Uri"
    try {
        $response = Invoke-WebRequest -Uri $Uri @requestArgs
    }
    catch {
        $status = $null
        if ($_.Exception.PSObject.Properties.Name -contains 'Response' -and $_.Exception.Response) {
            try { $status = [int]$_.Exception.Response.StatusCode } catch { }
        }
        switch ($status) {
            401 { throw "401 - Authentifizierung fehlgeschlagen. Token bzw. Zugangsdaten pruefen. ($Uri)" }
            403 { throw "403 - Keine Berechtigung fuer diese Ressource. ($Uri)" }
            404 { throw "404 - Nicht gefunden. BaseUrl/Kontextpfad oder Seiten-ID pruefen. ($Uri)" }
            default { throw "Fehler beim Aufruf von $Uri : $($_.Exception.Message)" }
        }
    }

    $bytes = $response.RawContentStream.ToArray()
    $json  = [System.Text.Encoding]::UTF8.GetString($bytes)
    return ($json | ConvertFrom-Json)
}

function Test-ConfluenceConnection {
    # Prueft die Anmeldung und gibt den erkannten Benutzer aus.
    try {
        $me = Invoke-ConfluenceApi -Uri "$BaseUrl/rest/api/user/current"
        if ($me -and $me.PSObject.Properties.Name -contains 'username') {
            Write-Host "Angemeldet als: $($me.username) ($($me.displayName))" -ForegroundColor DarkGray
        }
    }
    catch {
        if ($AuthMode -eq 'Anonymous') {
            Write-Host 'Zugriff anonym (kein angemeldeter Benutzer).' -ForegroundColor DarkGray
        }
        else {
            throw "Anmeldung nicht erfolgreich: $($_.Exception.Message)"
        }
    }
}

function Resolve-ParentPageId {
    param([string]$Space, [string]$Title)

    $uri = '{0}/rest/api/content?type=page&spaceKey={1}&title={2}&limit=2' -f $BaseUrl, [uri]::EscapeDataString($Space), [uri]::EscapeDataString($Title)

    $result = Invoke-ConfluenceApi -Uri $uri
    if (-not $result.results -or $result.results.Count -eq 0) {
        throw "Elternseite '$Title' im Space '$Space' nicht gefunden."
    }
    if ($result.results.Count -gt 1) {
        Write-Warning "Mehrere Treffer fuer '$Title' - es wird der erste verwendet."
    }
    return $result.results[0].id
}

function Get-DescendantPages {
    # Holt alle Nachfolgeseiten inkl. Storage-Body, seitenweise.
    param([string]$PageId)

    $pages = New-Object System.Collections.Generic.List[object]
    $start = 0

    do {
        if ($DirectChildrenOnly) {
            $uri = '{0}/rest/api/content/{1}/child/page?expand=body.storage,version&limit={2}&start={3}' -f $BaseUrl, $PageId, $PageSize, $start
        }
        else {
            $cql = "type = page and ancestor = $PageId order by title asc"
            $uri = '{0}/rest/api/content/search?cql={1}&expand=body.storage,version&limit={2}&start={3}' -f $BaseUrl, [uri]::EscapeDataString($cql), $PageSize, $start
        }

        $batch = Invoke-ConfluenceApi -Uri $uri
        $count = 0
        if ($batch.PSObject.Properties.Name -contains 'results' -and $batch.results) {
            foreach ($p in $batch.results) {
                $pages.Add($p) | Out-Null
                $count++
            }
        }

        Write-Verbose "  -> $count Seite(n) geladen (start=$start, gesamt=$($pages.Count))"
        $start += $PageSize
    } while ($count -eq $PageSize)

    # Komma verhindert, dass PowerShell die Liste aufloest
    return ,$pages.ToArray()
}

# --- Storage-Format auswerten ---------------------------------------------

function ConvertFrom-StorageHtml {
    # Wandelt einen Ausschnitt des Confluence-Storage-Formats in reinen Text um.
    param([string]$Html)

    if ([string]::IsNullOrWhiteSpace($Html)) { return '' }

    $text = $Html

    # Makro-Parameter und Layout-Metadaten entfernen, bevor Tags gestrippt werden
    $text = $text -replace '(?is)<ac:parameter\b[^>]*>.*?</ac:parameter>', ''
    $text = $text -replace '(?is)<ac:plain-text-body\b[^>]*>(.*?)</ac:plain-text-body>', '$1'
    $text = $text -replace '(?is)<ri:[^>]*/?>', ''

    # Zeilenumbrueche erhalten
    $text = $text -replace '(?is)<\s*br\s*/?>', "`n"
    $text = $text -replace '(?is)</\s*(p|li|div|h[1-6])\s*>', "`n"

    # Restliche Tags entfernen und Entities dekodieren
    $text = $text -replace '(?is)<[^>]+>', ''
    $text = [System.Net.WebUtility]::HtmlDecode($text)

    # Normalisieren: geschuetzte Leerzeichen, Zero-Width, Mehrfach-Whitespace
    $text = $text -replace "[\u00A0\u200B\u200C\uFEFF]", ' '
    $text = $text -replace "`r", ''
    $parts = $text -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
    $text  = ($parts -join ' ')
    $text  = $text -replace '\s{2,}', ' '

    return $text.Trim()
}

function Get-StorageTables {
    # Liefert alle Tabellen einer Seite getrennt zurueck:
    # Array von Tabellen, jede Tabelle ein Array von Zeilen, jede Zeile ein
    # Array von Zellentexten.
    param([string]$Storage)

    $tables = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($Storage)) { return ,$tables.ToArray() }

    foreach ($tableMatch in [regex]::Matches($Storage, '<table\b[^>]*>(.*?)</table>', 'Singleline, IgnoreCase')) {
        $rows = New-Object System.Collections.Generic.List[object]

        foreach ($rowMatch in [regex]::Matches($tableMatch.Groups[1].Value, '<tr\b[^>]*>(.*?)</tr>', 'Singleline, IgnoreCase')) {
            $cells = New-Object System.Collections.Generic.List[string]
            foreach ($cellMatch in [regex]::Matches($rowMatch.Groups[1].Value, '<t[dh]\b[^>]*>(.*?)</t[dh]>', 'Singleline, IgnoreCase')) {
                $cells.Add((ConvertFrom-StorageHtml $cellMatch.Groups[1].Value)) | Out-Null
            }
            if ($cells.Count -gt 0) { $rows.Add($cells.ToArray()) | Out-Null }
        }

        if ($rows.Count -gt 0) { $tables.Add($rows.ToArray()) | Out-Null }
    }

    return ,$tables.ToArray()
}

function Get-StorageTableTitles {
    # Liefert zu jeder Tabelle die Ueberschrift, die unmittelbar davor steht -
    # in derselben Reihenfolge wie Get-StorageTables. Damit laesst sich eine
    # Tabelle ueber ihren Titel identifizieren statt nur ueber ihren Inhalt.
    # Beruecksichtigt <h1>..<h6> und fett gesetzte Absaetze. Gesucht wird nur
    # im Text zwischen der vorherigen und dieser Tabelle, damit ein Titel nicht
    # auf mehrere Tabellen abfaerbt.
    param([string]$Storage)

    $titles = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Storage)) { return ,$titles.ToArray() }

    $letztesEnde = 0

    foreach ($tableMatch in [regex]::Matches($Storage, '<table\b[^>]*>(.*?)</table>', 'Singleline, IgnoreCase')) {

        # nur Tabellen mit Zeilen zaehlen - Get-StorageTables macht es genauso
        if (-not [regex]::IsMatch($tableMatch.Groups[1].Value, '<tr\b', 'IgnoreCase')) { continue }

        $vorText = $Storage.Substring($letztesEnde, $tableMatch.Index - $letztesEnde)
        $titel   = ''

        $muster  = '<(h[1-6])\b[^>]*>(.*?)</\1>|<p\b[^>]*>\s*<strong>(.*?)</strong>\s*</p>'
        $treffer = [regex]::Matches($vorText, $muster, 'Singleline, IgnoreCase')
        if ($treffer.Count -gt 0) {
            $letzter = $treffer[$treffer.Count - 1]
            $roh     = if ($letzter.Groups[2].Success) { $letzter.Groups[2].Value } else { $letzter.Groups[3].Value }
            $titel   = ConvertFrom-StorageHtml $roh
        }

        $titles.Add($titel) | Out-Null
        $letztesEnde = $tableMatch.Index + $tableMatch.Length
    }

    return ,$titles.ToArray()
}

function Get-NormalizedKey {
    # Vergleichsschluessel: Kleinschreibung, Umlaute als ae/oe/ue/ss, danach
    # nur Buchstaben und Ziffern. Dadurch findet die ASCII-Schreibweise
    # 'Kritikalitaet' auch die Zelle 'Kritikalitaet' mit Umlaut.
    param([string]$Value)
    if (-not $Value) { return '' }

    $text = $Value.ToLowerInvariant()
    $text = $text -replace ([string][char]0x00E4), 'ae'   # a-Umlaut
    $text = $text -replace ([string][char]0x00F6), 'oe'   # o-Umlaut
    $text = $text -replace ([string][char]0x00FC), 'ue'   # u-Umlaut
    $text = $text -replace ([string][char]0x00DF), 'ss'   # scharfes s

    return ($text -replace '[^\p{L}\p{Nd}]', '')
}

function Get-CellValueByLabel {
    # Sucht ueber alle Tabellen die Zeile, deren erste Zelle die Beschriftung
    # enthaelt, und liefert den Wert der zweiten Spalte. Erkennt zusaetzlich
    # horizontale Tabellen (Beschriftung als Kopfzeile).
    param(
        [Parameter(Mandatory)]$Tables,
        [Parameter(Mandatory)][string]$Label
    )

    $key = Get-NormalizedKey $Label
    if (-not $key) { return $null }

    # Variante A: Beschriftung in Spalte 1, Wert in Spalte 2 (Standardfall)
    foreach ($table in $Tables) {
        foreach ($row in $table) {
            if ($row.Count -ge 2) {
                $cellKey = Get-NormalizedKey $row[0]
                if ($cellKey -and ($cellKey -eq $key -or $cellKey.Contains($key) -or $key.Contains($cellKey))) {
                    return $row[1]
                }
            }
        }
    }

    # Variante B: Beschriftung als Spaltenkopf, Wert in der Zeile darunter
    foreach ($table in $Tables) {
        if ($table.Count -ge 2) {
            $header = $table[0]
            for ($i = 0; $i -lt $header.Count; $i++) {
                $cellKey = Get-NormalizedKey $header[$i]
                if ($cellKey -and ($cellKey -eq $key -or $cellKey.Contains($key) -or $key.Contains($cellKey))) {
                    if ($table[1].Count -gt $i) { return $table[1][$i] }
                }
            }
        }
    }

    return $null
}

function ConvertTo-JaNein {
    # Vereinheitlicht Ja/Nein-Schreibweisen; alles andere bleibt unveraendert.
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    if (-not $NormalizeJaNein) { return $Value.Trim() }

    $compact = ($Value -replace '\s', '')
    switch -Regex ($compact) {
        '^(?i)(ja|j|x|yes|y|true|wahr)$' { return 'Ja' }
        '^(?i)(nein|n|no|false|falsch)$' { return 'Nein' }
        default                          { return $Value.Trim() }
    }
}

function Get-EbenenCode {
    # "Entwicklung (ENTW)" -> ENTW ; "Freigabetest (TEST) Alte PNT Welt" -> TEST
    param([string]$Label)

    $match = [regex]::Match($Label, '\(([^)]{1,24})\)')
    if ($match.Success) {
        $code = ($match.Groups[1].Value -replace '[^\p{L}\p{Nd}]', '').ToUpperInvariant()
        if ($code) { return $code }
    }

    $fallback = ($Label -replace '[^\p{L}\p{Nd}]', '').ToUpperInvariant()
    if ($fallback.Length -gt 20) { $fallback = $fallback.Substring(0, 20) }
    return $fallback
}

function Get-CellAt {
    # Liefert den Zellwert an der angegebenen Position oder '' bei fehlender Spalte.
    param($Row, [int]$Index)

    if ($Index -lt 0)          { return '' }
    if ($Row.Count -le $Index) { return '' }
    return [string]$Row[$Index]
}

function Get-AnwendungsebenenMatrix {
    # Wertet die Matrix-Tabelle aus:
    #   Anwendungsebene | IST (in Legacy-Umgebung) | SOLL (im RZ-DRV) | Anzahl IST-Server
    # Zusaetzliche Spalten werden ignoriert, fehlende bleiben leer.
    # Liefert Found, Standard (feste Ebenen) und Zusatz (die Zeilen unterhalb
    # von "Zusaetzlich benoetigte Anwendungsebenen").
    param([Parameter(Mandatory)]$Tables)

    $standard   = New-Object System.Collections.Generic.List[object]
    $zusatz     = New-Object System.Collections.Generic.List[object]
    $headerKey  = Get-NormalizedKey $MatrixHeaderLabel
    $zusatzKey  = Get-NormalizedKey $MatrixZusatzLabel

    foreach ($table in $Tables) {
        if ($table.Count -lt 2) { continue }

        # Kopfzeile suchen (muss nicht die erste Zeile sein)
        $headerIndex = -1
        for ($r = 0; $r -lt $table.Count; $r++) {
            if ($table[$r].Count -ge 2) {
                $cellKey = Get-NormalizedKey $table[$r][0]
                if ($cellKey -and $cellKey.Contains($headerKey)) { $headerIndex = $r; break }
            }
        }
        if ($headerIndex -lt 0) { continue }

        $header    = $table[$headerIndex]
        $colIst    = -1
        $colSoll   = -1
        $colAnzahl = -1

        # Reihenfolge wichtig: "Anzahl IST-Server" enthaelt ebenfalls "IST"
        for ($i = 1; $i -lt $header.Count; $i++) {
            $h = Get-NormalizedKey $header[$i]
            if (-not $h) { continue }
            if ($colAnzahl -lt 0 -and $h.StartsWith('anzahl')) { $colAnzahl = $i; continue }
            if ($colSoll   -lt 0 -and $h.StartsWith('soll'))   { $colSoll   = $i; continue }
            if ($colIst    -lt 0 -and $h.StartsWith('ist'))    { $colIst    = $i; continue }
        }

        # Fallback ueber die Position, falls die Kopfzeile abweicht
        if ($colIst    -lt 0 -and $header.Count -ge 2) { $colIst    = 1 }
        if ($colSoll   -lt 0 -and $header.Count -ge 3) { $colSoll   = 2 }
        if ($colAnzahl -lt 0 -and $header.Count -ge 4) { $colAnzahl = 3 }

        Write-Verbose "  Matrix erkannt: IST=Spalte $colIst, SOLL=Spalte $colSoll, Anzahl=Spalte $colAnzahl"

        $inZusatz = $false
        $nr       = 0

        for ($r = $headerIndex + 1; $r -lt $table.Count; $r++) {
            $row = $table[$r]
            if ($row.Count -lt 1) { continue }

            $ebene    = ([string]$row[0]).Trim()
            $ebeneKey = Get-NormalizedKey $ebene

            # Trennzeile: ab hier folgen die frei erfassten Ebenen
            if (-not $inZusatz -and $zusatzKey -and $ebeneKey -and $ebeneKey.Contains($zusatzKey)) {
                $inZusatz = $true
                Write-Verbose "  Zusatzebenen ab Zeile $r"
                continue
            }

            $ist    = ConvertTo-JaNein (Get-CellAt -Row $row -Index $colIst)
            $soll   = ConvertTo-JaNein (Get-CellAt -Row $row -Index $colSoll)
            $anzahl = (Get-CellAt -Row $row -Index $colAnzahl).Trim()

            if ($inZusatz) {
                # Leere Vorlagenzeilen ueberspringen
                if (-not $ebene -and -not $ist -and -not $soll -and -not $anzahl) { continue }
                $nr++
                $zusatz.Add([pscustomobject]@{
                    Nr     = $nr
                    Ebene  = $ebene
                    IST    = $ist
                    SOLL   = $soll
                    Anzahl = $anzahl
                }) | Out-Null
            }
            else {
                if (-not $ebene -or $row.Count -lt 2) { continue }
                $code = Get-EbenenCode $ebene
                if (-not $code) { continue }
                $standard.Add([pscustomobject]@{
                    Ebene  = $ebene
                    Code   = $code
                    IST    = $ist
                    SOLL   = $soll
                    Anzahl = $anzahl
                }) | Out-Null
            }
        }

        # erste passende Tabelle genuegt
        return [pscustomobject]@{
            Found    = $true
            Standard = $standard.ToArray()
            Zusatz   = $zusatz.ToArray()
        }
    }

    return [pscustomobject]@{ Found = $false; Standard = @(); Zusatz = @() }
}

function Get-SchnittstellenTabelle {
    # Wertet die Schnittstellen-Tabelle aus:
    #   Input/Output | Ziel- bzw. Quellsystem | Manuell/Halbautom./Autom. |
    #   Kurzbeschreibung zu uebergebende Daten
    # Liefert ein Objekt mit Found (Tabelle vorhanden) und Rows (Datenzeilen).
    # Leere Vorlagenzeilen werden uebersprungen.
    param([Parameter(Mandatory)]$Tables)

    $rowsOut = New-Object System.Collections.Generic.List[object]

    foreach ($table in $Tables) {
        if ($table.Count -lt 2) { continue }

        # Kopfzeile suchen: erste Zeile, in der mindestens 2 der 4 Spalten erkannt werden
        $headerIndex = -1
        $colRichtung = -1
        $colSystem   = -1
        $colArt      = -1
        $colDaten    = -1

        for ($r = 0; $r -lt $table.Count; $r++) {
            $cRichtung = -1; $cSystem = -1; $cArt = -1; $cDaten = -1
            $row = $table[$r]

            for ($i = 0; $i -lt $row.Count; $i++) {
                $h = Get-NormalizedKey $row[$i]
                if (-not $h) { continue }

                if ($cRichtung -lt 0 -and ($h.StartsWith('input') -or $h.StartsWith('output'))) { $cRichtung = $i; continue }
                if ($cArt      -lt 0 -and ($h.StartsWith('manuell') -or $h.Contains('halbautom'))) { $cArt = $i; continue }
                if ($cSystem   -lt 0 -and ($h.Contains('zielbzw') -or $h.Contains('quellsystem') -or $h.Contains('zielsystem'))) { $cSystem = $i; continue }
                if ($cDaten    -lt 0 -and ($h.StartsWith('kurzbeschreibung') -or $h.Contains('bergebende'))) { $cDaten = $i; continue }
            }

            $treffer = @($cRichtung, $cSystem, $cArt, $cDaten) | Where-Object { $_ -ge 0 }
            if (@($treffer).Count -ge 2) {
                $headerIndex = $r
                $colRichtung = $cRichtung
                $colSystem   = $cSystem
                $colArt      = $cArt
                $colDaten    = $cDaten
                break
            }
        }

        if ($headerIndex -lt 0) { continue }

        Write-Verbose "  Schnittstellen erkannt: Richtung=$colRichtung, System=$colSystem, Art=$colArt, Daten=$colDaten"

        $nr = 0
        for ($r = $headerIndex + 1; $r -lt $table.Count; $r++) {
            $row = $table[$r]

            $richtung = (Get-CellAt -Row $row -Index $colRichtung).Trim()
            $system   = (Get-CellAt -Row $row -Index $colSystem).Trim()
            $art      = (Get-CellAt -Row $row -Index $colArt).Trim()
            $daten    = (Get-CellAt -Row $row -Index $colDaten).Trim()

            # Leere Vorlagenzeilen ueberspringen
            if (-not $richtung -and -not $system -and -not $art -and -not $daten) { continue }

            $nr++
            $rowsOut.Add([pscustomobject]@{
                Nr       = $nr
                Richtung = $richtung
                System   = $system
                Art      = $art
                Daten    = $daten
            }) | Out-Null
        }

        return [pscustomobject]@{ Found = $true; Rows = $rowsOut.ToArray() }
    }

    return [pscustomobject]@{ Found = $false; Rows = @() }
}

function Get-SperrzeitenTabelle {
    # Wertet den Sperrzeiten-Block der Migrations-Tabelle aus:
    #   Zeile "Sperrzeiten"
    #   Kopfzeile   von | bis | Bemerkung/Name
    #   Datenzeilen ...
    # Liefert Found (Block vorhanden) und Rows (Zeitraeume).
    param([Parameter(Mandatory)]$Tables)

    $rowsOut    = New-Object System.Collections.Generic.List[object]
    $ankerKey   = Get-NormalizedKey $SperrzeitenLabel

    # "von" und "bis" sind sehr kurz - hier wird exakt bzw. eng geprueft,
    # damit nicht zufaellig eine andere Zelle als Kopfzeile durchgeht.
    function Test-KurzKey {
        param([string]$Key, [string[]]$Varianten)
        if (-not $Key) { return $false }
        foreach ($v in $Varianten) {
            if ($Key -eq $v) { return $true }
            if ($Key.StartsWith($v) -and $Key.Length -le 12) { return $true }
        }
        return $false
    }

    foreach ($table in $Tables) {
        if ($table.Count -lt 2) { continue }

        # 1. Anker "Sperrzeiten" suchen (irgendeine Zelle der Zeile)
        $ankerIndex = -1
        if ($ankerKey) {
            for ($r = 0; $r -lt $table.Count; $r++) {
                foreach ($cell in $table[$r]) {
                    $k = Get-NormalizedKey $cell
                    if ($k -and $k.Contains($ankerKey)) { $ankerIndex = $r; break }
                }
                if ($ankerIndex -ge 0) { break }
            }
        }

        # 2. Kopfzeile mit von/bis suchen - ab dem Anker, sonst ab Tabellenanfang
        $startIndex  = if ($ankerIndex -ge 0) { $ankerIndex + 1 } else { 0 }
        $headerIndex = -1
        $colVon      = -1
        $colBis      = -1
        $colBemerk   = -1

        for ($r = $startIndex; $r -lt $table.Count; $r++) {
            $row = $table[$r]
            $cVon = -1; $cBis = -1; $cBem = -1

            for ($i = 0; $i -lt $row.Count; $i++) {
                $h = Get-NormalizedKey $row[$i]
                if (-not $h) { continue }
                if ($cVon -lt 0 -and (Test-KurzKey -Key $h -Varianten @('von', 'vom'))) { $cVon = $i; continue }
                if ($cBis -lt 0 -and (Test-KurzKey -Key $h -Varianten @('bis'))) { $cBis = $i; continue }
                if ($cBem -lt 0 -and ($h.StartsWith('bemerkung') -or $h.StartsWith('name') -or $h.Contains('bemerkung'))) { $cBem = $i; continue }
            }

            # von UND bis muessen zusammen auftreten - das ist das eindeutige Signal
            if ($cVon -ge 0 -and $cBis -ge 0) {
                $headerIndex = $r
                $colVon      = $cVon
                $colBis      = $cBis
                $colBemerk   = $cBem
                break
            }
        }

        if ($headerIndex -lt 0) { continue }

        # Bemerkungsspalte notfalls ueber die Position bestimmen
        if ($colBemerk -lt 0) {
            $kandidat = [Math]::Max($colVon, $colBis) + 1
            if ($table[$headerIndex].Count -gt $kandidat) { $colBemerk = $kandidat }
        }

        Write-Verbose "  Sperrzeiten erkannt: von=Spalte $colVon, bis=Spalte $colBis, Bemerkung=Spalte $colBemerk"

        $nr = 0
        for ($r = $headerIndex + 1; $r -lt $table.Count; $r++) {
            $row = $table[$r]

            $von     = (Get-CellAt -Row $row -Index $colVon).Trim()
            $bis     = (Get-CellAt -Row $row -Index $colBis).Trim()
            $bemerk  = (Get-CellAt -Row $row -Index $colBemerk).Trim()

            # Leere Vorlagenzeilen ueberspringen
            if (-not $von -and -not $bis -and -not $bemerk) { continue }

            # Der Sperrzeiten-Block endet, sobald eine Zeile folgt, die keine
            # Zeitangabe mehr enthaelt - typisch fuer die naechste
            # Beschriftungszeile der Migrations-Tabelle ("Ansprechpartner
            # Migration" o. ae.). Erkannt wird das daran, dass weder "von" noch
            # "bis" eine Ziffer enthaelt. Eine Sperrzeit ohne jede Ziffer
            # (z. B. "nach Absprache") wuerde damit ebenfalls den Block
            # beenden - in dem Fall die Zeile auf der Seite datieren.
            if (($von + $bis) -notmatch '\d') {
                Write-Verbose "  Sperrzeiten-Block endet vor Zeile $r ('$von')"
                break
            }

            $nr++
            $rowsOut.Add([pscustomobject]@{
                Nr        = $nr
                Von       = $von
                Bis       = $bis
                Bemerkung = $bemerk
            }) | Out-Null
        }

        return [pscustomobject]@{ Found = $true; Rows = $rowsOut.ToArray() }
    }

    return [pscustomobject]@{ Found = $false; Rows = @() }
}

function Get-ServerTabelle {
    # Wertet die Tabelle "Kapazitaetsanforderungen Server" aus:
    #   Servername | Cluster [Ja/Nein] | Anwendungsebene | Betriebssystem + Version
    #   | # CPU | # RAM | Festplatten in GB | Dediziert / Virtuell
    # Liefert Found (Tabelle vorhanden) und Rows (eine je Server).
    param([Parameter(Mandatory)]$Tables)

    $rowsOut = New-Object System.Collections.Generic.List[object]

    foreach ($table in $Tables) {
        if ($table.Count -lt 2) { continue }

        $headerIndex = -1
        $colName = -1; $colCluster = -1; $colEbene = -1; $colOs = -1
        $colCpu  = -1; $colRam     = -1; $colHdd   = -1; $colArt = -1

        for ($r = 0; $r -lt $table.Count; $r++) {
            $row = $table[$r]
            $cName = -1; $cCluster = -1; $cEbene = -1; $cOs = -1
            $cCpu  = -1; $cRam     = -1; $cHdd   = -1; $cArt = -1

            for ($i = 0; $i -lt $row.Count; $i++) {
                $h = Get-NormalizedKey $row[$i]
                if (-not $h) { continue }

                # StartsWith statt Contains: sonst wuerde "Anzahl IST-Server"
                # aus der Anwendungsebenen-Matrix als Servername durchgehen.
                if ($cName    -lt 0 -and $h.StartsWith('servername'))      { $cName    = $i; continue }
                if ($cCluster -lt 0 -and $h.StartsWith('cluster'))         { $cCluster = $i; continue }
                if ($cEbene   -lt 0 -and $h.StartsWith('anwendungsebene')) { $cEbene   = $i; continue }
                if ($cOs      -lt 0 -and $h.StartsWith('betriebssystem'))  { $cOs      = $i; continue }
                if ($cCpu     -lt 0 -and $h.StartsWith('cpu'))             { $cCpu     = $i; continue }
                if ($cRam     -lt 0 -and $h.StartsWith('ram'))             { $cRam     = $i; continue }
                if ($cHdd     -lt 0 -and $h.StartsWith('festplatte'))      { $cHdd     = $i; continue }
                if ($cArt     -lt 0 -and ($h.StartsWith('dediziert') -or $h.StartsWith('virtuell'))) { $cArt = $i; continue }
            }

            $treffer = @($cName, $cCluster, $cEbene, $cOs, $cCpu, $cRam, $cHdd, $cArt) | Where-Object { $_ -ge 0 }
            if (@($treffer).Count -ge 3) {
                $headerIndex = $r
                $colName = $cName; $colCluster = $cCluster; $colEbene = $cEbene; $colOs = $cOs
                $colCpu  = $cCpu;  $colRam     = $cRam;     $colHdd   = $cHdd;   $colArt = $cArt
                break
            }
        }

        if ($headerIndex -lt 0) { continue }

        Write-Verbose "  Server-Tabelle erkannt: Name=$colName, Cluster=$colCluster, Ebene=$colEbene, OS=$colOs, CPU=$colCpu, RAM=$colRam, HDD=$colHdd, Art=$colArt"

        $nr = 0
        for ($r = $headerIndex + 1; $r -lt $table.Count; $r++) {
            $row = $table[$r]

            $name    = (Get-CellAt -Row $row -Index $colName).Trim()
            $cluster = (Get-CellAt -Row $row -Index $colCluster).Trim()
            $ebene   = (Get-CellAt -Row $row -Index $colEbene).Trim()
            $os      = (Get-CellAt -Row $row -Index $colOs).Trim()
            $cpu     = (Get-CellAt -Row $row -Index $colCpu).Trim()
            $ram     = (Get-CellAt -Row $row -Index $colRam).Trim()
            $hdd     = (Get-CellAt -Row $row -Index $colHdd).Trim()
            $art     = (Get-CellAt -Row $row -Index $colArt).Trim()

            # Leere Vorlagenzeilen ueberspringen
            if (-not $name -and -not $cluster -and -not $ebene -and -not $os -and
                -not $cpu  -and -not $ram     -and -not $hdd   -and -not $art) { continue }

            $nr++
            $rowsOut.Add([pscustomobject]@{
                Nr            = $nr
                Servername    = $name
                Cluster       = (ConvertTo-JaNein $cluster)
                Ebene         = $ebene
                Betriebssystem = $os
                CPU           = $cpu
                RAM           = $ram
                Festplatten   = $hdd
                Art           = $art
            }) | Out-Null
        }

        return [pscustomobject]@{ Found = $true; Rows = $rowsOut.ToArray() }
    }

    return [pscustomobject]@{ Found = $false; Rows = @() }
}

function Get-DatenbankTabelle {
    # Wertet die Tabelle "Kapazitaetsanforderungen Datenbank" aus:
    #   Datenbank | Anwendungsebene | Groesse | Plattform
    # Liefert Found (Tabelle vorhanden) und Rows (eine je Datenbank).
    param([Parameter(Mandatory)]$Tables)

    $rowsOut = New-Object System.Collections.Generic.List[object]

    foreach ($table in $Tables) {
        if ($table.Count -lt 2) { continue }

        $headerIndex = -1
        $colDb = -1; $colEbene = -1; $colGroesse = -1; $colPlattform = -1

        for ($r = 0; $r -lt $table.Count; $r++) {
            $row = $table[$r]
            $cDb = -1; $cEbene = -1; $cGroesse = -1; $cPlattform = -1

            for ($i = 0; $i -lt $row.Count; $i++) {
                $h = Get-NormalizedKey $row[$i]
                if (-not $h) { continue }

                if ($cDb        -lt 0 -and $h.StartsWith('datenbank'))       { $cDb        = $i; continue }
                if ($cEbene     -lt 0 -and $h.StartsWith('anwendungsebene')) { $cEbene     = $i; continue }
                if ($cGroesse   -lt 0 -and $h.StartsWith('groesse'))         { $cGroesse   = $i; continue }
                if ($cPlattform -lt 0 -and ($h.StartsWith('plattform') -or $h.StartsWith('platform'))) { $cPlattform = $i; continue }
            }

            # "Datenbank" ist Pflicht - diese Spalte gibt es in keiner anderen
            # Tabelle. Dazu mindestens eine weitere erkannte Spalte.
            $treffer = @($cDb, $cEbene, $cGroesse, $cPlattform) | Where-Object { $_ -ge 0 }
            if ($cDb -ge 0 -and @($treffer).Count -ge 2) {
                $headerIndex  = $r
                $colDb        = $cDb
                $colEbene     = $cEbene
                $colGroesse   = $cGroesse
                $colPlattform = $cPlattform
                break
            }
        }

        if ($headerIndex -lt 0) { continue }

        Write-Verbose "  Datenbank-Tabelle erkannt: DB=$colDb, Ebene=$colEbene, Groesse=$colGroesse, Plattform=$colPlattform"

        $nr = 0
        for ($r = $headerIndex + 1; $r -lt $table.Count; $r++) {
            $row = $table[$r]

            $db        = (Get-CellAt -Row $row -Index $colDb).Trim()
            $ebene     = (Get-CellAt -Row $row -Index $colEbene).Trim()
            $groesse   = (Get-CellAt -Row $row -Index $colGroesse).Trim()
            $plattform = (Get-CellAt -Row $row -Index $colPlattform).Trim()

            # Leere Vorlagenzeilen ueberspringen
            if (-not $db -and -not $ebene -and -not $groesse -and -not $plattform) { continue }

            $nr++
            $rowsOut.Add([pscustomobject]@{
                Nr        = $nr
                Datenbank = $db
                Ebene     = $ebene
                Groesse   = $groesse
                Plattform = $plattform
            }) | Out-Null
        }

        return [pscustomobject]@{ Found = $true; Rows = $rowsOut.ToArray() }
    }

    return [pscustomobject]@{ Found = $false; Rows = @() }
}

function Get-KategorieTabelle {
    # Wertet eine in Bloecke gegliederte Tabelle aus - so aufgebaut sind
    # "Aufbau Anwendungsserver" und "Umsysteme":
    #     Appliance                 <- Kategoriezeile
    #     <Datenzeile>              <- gehoert zu Appliance
    #     <Datenzeile>
    #     Datenbank                 <- naechste Kategoriezeile
    #     ...
    # Zwei Schreibweisen werden unterstuetzt:
    #   A) Kategoriezeile steht allein, Datenzeilen nutzen alle Spalten
    #   B) Kategorie steht in Spalte 1, die Werte daneben
    # Welche Tabelle gemeint ist, entscheidet die Liste der Kategorien.
    # Liefert Found, Spalten (Ueberschriften) und Rows (Kategorie + Werte).
    param(
        [Parameter(Mandatory)]$Tables,
        [Parameter(Mandatory)][string[]]$Kategorien,
        $Titel = @(),          # Ueberschriften der Tabellen (gleiche Reihenfolge)
        [string]$TitelFilter = ''   # wenn gesetzt: nur diese Tabelle auswerten
    )

    $rowsOut = New-Object System.Collections.Generic.List[object]
    $spalten = New-Object System.Collections.Generic.List[string]

    # Steht ein Titel fest, wird genau die Tabelle mit dieser Ueberschrift
    # ausgewertet. Nur wenn keine passt, wird ueber die Kategorien gesucht.
    $nurIndex = -1
    if ($TitelFilter) {
        $filterKey = Get-NormalizedKey $TitelFilter
        for ($t = 0; $t -lt $Tables.Count; $t++) {
            if ($Titel.Count -le $t) { break }
            $tk = Get-NormalizedKey $Titel[$t]
            if ($tk -and $filterKey -and $tk.Contains($filterKey)) { $nurIndex = $t; break }
        }
        if ($nurIndex -ge 0) {
            Write-Verbose "  Tabelle ueber Ueberschrift '$($Titel[$nurIndex])' gefunden"
        }
        else {
            Write-Verbose "  Keine Tabelle mit Ueberschrift '$TitelFilter' - Suche ueber die Kategorien"
        }
    }

    # Liefert den Kategorienamen, wenn der Text einer Kategorie entspricht
    function Resolve-Kategorie {
        param([string]$Text, [string[]]$Liste)
        $key = Get-NormalizedKey $Text
        if (-not $key) { return $null }
        foreach ($kat in $Liste) {
            $katKey = Get-NormalizedKey $kat
            if (-not $katKey) { continue }
            if ($key -eq $katKey) { return $kat }
            if ($key.StartsWith($katKey)) { return $kat }
            # Abkuerzungen nur ab vier Zeichen, sonst wuerde ein einzelnes
            # "A" in einer Datenzelle schon als "Appliance" durchgehen.
            if ($key.Length -ge 4 -and $katKey.StartsWith($key)) { return $kat }
        }
        return $null
    }

    for ($tIdx = 0; $tIdx -lt $Tables.Count; $tIdx++) {

        # Wurde die Tabelle ueber die Ueberschrift bestimmt, nur diese ansehen
        if ($nurIndex -ge 0 -and $tIdx -ne $nurIndex) { continue }

        $table = $Tables[$tIdx]
        if ($table.Count -lt 2) { continue }

        # Kategoriezeilen der ersten Spalte suchen
        $katZeilen = New-Object System.Collections.Generic.List[int]
        for ($r = 0; $r -lt $table.Count; $r++) {
            if ($table[$r].Count -lt 1) { continue }
            if (Resolve-Kategorie -Text $table[$r][0] -Liste $Kategorien) { $katZeilen.Add($r) | Out-Null }
        }

        # Ohne Titeltreffer muessen mindestens zwei Kategorien zusammenkommen,
        # damit nicht die falsche Tabelle erwischt wird. Steht die Tabelle ueber
        # ihre Ueberschrift fest, genuegt eine.
        $mindest = if ($nurIndex -ge 0) { 1 } else { 2 }
        if ($katZeilen.Count -lt $mindest) { continue }

        # Modus bestimmen: stehen neben den Kategorienamen noch Werte?
        $mitWertenDaneben = 0
        foreach ($r in $katZeilen) {
            $rest = @($table[$r] | Select-Object -Skip 1 | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($rest.Count -gt 0) { $mitWertenDaneben++ }
        }
        # Variante B, wenn die Mehrheit der Kategoriezeilen Werte daneben hat
        $offset = if ($mitWertenDaneben * 2 -gt $katZeilen.Count) { 1 } else { 0 }

        # Kopfzeile: eine Zeile vor der ersten Kategorie mit mehreren Werten
        $ersteKat  = $katZeilen[0]
        $startRow  = 0
        if ($ersteKat -gt 0) {
            $kandidat = $table[$ersteKat - 1]
            $gefuellt = @($kandidat | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($gefuellt.Count -ge 2) {
                for ($i = $offset; $i -lt $kandidat.Count; $i++) {
                    $name = ([string]$kandidat[$i]).Trim()
                    if (-not $name) { $name = "Wert$($i - $offset + 1)" }
                    $spalten.Add($name) | Out-Null
                }
                $startRow = $ersteKat
            }
        }

        Write-Verbose "  Kategorie-Tabelle erkannt: $($katZeilen.Count) Kategorien, Offset $offset"

        $aktuelleKategorie = ''
        $nr = 0

        for ($r = $startRow; $r -lt $table.Count; $r++) {
            $row = $table[$r]
            if ($row.Count -lt 1) { continue }

            $erste = ([string]$row[0]).Trim()
            $kat   = Resolve-Kategorie -Text $erste -Liste $Kategorien
            $rest  = @($row | Select-Object -Skip 1 | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

            if ($offset -eq 0) {
                # Variante A: eine Zeile mit nur einer gefuellten Zelle ist eine
                # Gruppenueberschrift. Steht sie in der Kategorienliste, gelten
                # die folgenden Zeilen fuer sie. Ist sie eine FREMDE Ueberschrift
                # (ein Block, der nicht exportiert werden soll), endet der
                # bisherige Block - sonst wuerden dessen Zeilen faelschlich
                # unter der zuletzt erkannten Kategorie landen.
                if ($rest.Count -eq 0 -and $erste) {
                    if ($kat) {
                        $aktuelleKategorie = $erste
                    }
                    elseif ($KategorieBlockStrikt) {
                        Write-Verbose "  Fremder Block '$erste' - Zeilen darunter werden uebersprungen"
                        $aktuelleKategorie = ''
                    }
                    continue
                }
            }
            else {
                # Variante B: jede Zeile traegt ihre Kategorie in Spalte 1.
                # Zeilen mit einer anderen Beschriftung gehoeren nicht dazu.
                if (-not $kat) {
                    Write-Verbose "  Zeile '$erste' ist keine gesuchte Kategorie - uebersprungen"
                    continue
                }
                $aktuelleKategorie = $erste
            }

            # Werte ab Offset einsammeln
            $werte = New-Object System.Collections.Generic.List[string]
            for ($i = $offset; $i -lt $row.Count; $i++) {
                $werte.Add((([string]$row[$i]).Trim())) | Out-Null
            }

            # Leere Vorlagenzeilen ueberspringen
            $hatInhalt = @($werte | Where-Object { $_ -ne '' }).Count -gt 0
            if (-not $hatInhalt) { continue }

            # Zeilen vor der ersten Kategorie oder nach einem fremden Block
            # gehoeren zu keiner gesuchten Kategorie
            if (-not $aktuelleKategorie) { continue }

            $nr++
            $rowsOut.Add([pscustomobject]@{
                Nr        = $nr
                Kategorie = $aktuelleKategorie
                Werte     = $werte.ToArray()
            }) | Out-Null
        }

        # Spaltennamen notfalls generisch ergaenzen
        $maxWerte = 0
        foreach ($z in $rowsOut) { if ($z.Werte.Count -gt $maxWerte) { $maxWerte = $z.Werte.Count } }
        while ($spalten.Count -lt $maxWerte) { $spalten.Add("Wert$($spalten.Count + 1)") | Out-Null }

        return [pscustomobject]@{
            Found   = $true
            Spalten = $spalten.ToArray()
            Rows    = $rowsOut.ToArray()
        }
    }

    return [pscustomobject]@{ Found = $false; Spalten = @(); Rows = @() }
}

function Export-CsvUtf8Bom {
    # CSV mit UTF-8-BOM schreiben, damit Excel Umlaute korrekt anzeigt.
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$Path
    )

    if (-not [System.IO.Path]::IsPathRooted($Path)) {
        $Path = Join-Path -Path (Get-Location).ProviderPath -ChildPath $Path
    }
    $dir = Split-Path -Path $Path -Parent
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }

    $lines = $Data | ConvertTo-Csv -NoTypeInformation -Delimiter $Delimiter
    [System.IO.File]::WriteAllLines($Path, $lines, (New-Object System.Text.UTF8Encoding($true)))

    return $Path
}

# ===========================================================================
#  HAUPTABLAUF
# ===========================================================================

Test-ConfluenceConnection

if ([string]::IsNullOrWhiteSpace($ParentPageId)) {
    if ([string]::IsNullOrWhiteSpace($SpaceKey) -or [string]::IsNullOrWhiteSpace($ParentTitle)) {
        throw 'Weder $ParentPageId noch $SpaceKey/$ParentTitle sind gesetzt.'
    }
    Write-Verbose "Ermittle Seiten-ID fuer '$ParentTitle' im Space '$SpaceKey' ..."
    $ParentPageId = Resolve-ParentPageId -Space $SpaceKey -Title $ParentTitle
    Write-Host "Elternseite '$ParentTitle' hat die ID $ParentPageId" -ForegroundColor DarkGray
}

Write-Host "Lade Seiten unterhalb von Seiten-ID $ParentPageId ..." -ForegroundColor Cyan
$pages = Get-DescendantPages -PageId $ParentPageId
Write-Host "$($pages.Count) Seite(n) gefunden." -ForegroundColor Cyan

if ($pages.Count -eq 0) {
    Write-Warning 'Keine Unterseiten gefunden - Abbruch.'
    return
}

$pageData     = New-Object System.Collections.Generic.List[object]
$noLabels     = New-Object System.Collections.Generic.List[string]
$noMatrix     = New-Object System.Collections.Generic.List[string]
$noSchnitt    = New-Object System.Collections.Generic.List[string]
$noSperr      = New-Object System.Collections.Generic.List[string]
$noServer     = New-Object System.Collections.Generic.List[string]
$noDatenbank  = New-Object System.Collections.Generic.List[string]
$noAufbau     = New-Object System.Collections.Generic.List[string]
$aufbauCols   = New-Object System.Collections.Generic.List[string]
$noUmsysteme  = New-Object System.Collections.Generic.List[string]
$umsysCols    = New-Object System.Collections.Generic.List[string]
$codeOrder    = New-Object System.Collections.Generic.List[string]
$codeToLabel  = @{}
$processed    = 0

# Vorgegebene Spaltenreihenfolge uebernehmen
foreach ($c in $Anwendungsebenen) {
    $upper = $c.ToUpperInvariant()
    if (-not $codeOrder.Contains($upper)) { $codeOrder.Add($upper) | Out-Null }
}

foreach ($page in $pages) {
    $processed++
    $progressArgs = @{
        Activity        = 'Seiten werden ausgewertet'
        Status          = "$processed von $($pages.Count): $($page.title)"
        PercentComplete = [int](($processed / $pages.Count) * 100)
    }
    Write-Progress @progressArgs

    $storage = ''
    if ($page.PSObject.Properties.Name -contains 'body' -and $page.body -and
        $page.body.PSObject.Properties.Name -contains 'storage' -and $page.body.storage) {
        $storage = [string]$page.body.storage.value
    }

    $tables      = Get-StorageTables -Storage $storage
    $tableTitles = Get-StorageTableTitles -Storage $storage

    # --- Tabelle 1: Verfahren / Beschreibung ---
    $verfahren    = Get-CellValueByLabel -Tables $tables -Label $NameLabel
    $beschreibung = Get-CellValueByLabel -Tables $tables -Label $DescriptionLabel

    if ($UsePositionFallback -and $tables.Count -ge 1) {
        $first = $tables[0]
        if (-not $verfahren    -and $first.Count -ge 1 -and $first[0].Count -ge 2) { $verfahren    = $first[0][1] }
        if (-not $beschreibung -and $first.Count -ge 2 -and $first[1].Count -ge 2) { $beschreibung = $first[1][1] }
    }

    if (-not $verfahren -and -not $beschreibung) {
        $noLabels.Add($page.title) | Out-Null
        Write-Verbose "  ! Keine Beschriftungstabelle auf '$($page.title)'"
    }

    # --- Tabelle "Allgemeine Information" und weitere Beschriftungszeilen ---
    $weitere = [ordered]@{}
    foreach ($spalte in $WeitereFelder.Keys) {

        $wert = ''
        foreach ($label in @($WeitereFelder[$spalte])) {
            $treffer = Get-CellValueByLabel -Tables $tables -Label $label
            if (-not [string]::IsNullOrWhiteSpace($treffer)) { $wert = $treffer; break }
        }

        if (-not $wert) {
            Write-Verbose "  ! Spalte '$spalte' nicht gefunden auf '$($page.title)'"
        }
        $weitere[$spalte] = $wert
    }

    # --- Tabelle 2: Anwendungsebenen-Matrix ---
    $matrixResult = Get-AnwendungsebenenMatrix -Tables $tables
    $matrix       = $matrixResult.Standard
    $zusatzEbenen = $matrixResult.Zusatz
    $matrixMap    = @{}
    foreach ($entry in $matrix) {
        if (-not $matrixMap.ContainsKey($entry.Code)) { $matrixMap[$entry.Code] = $entry }
        if (-not $codeOrder.Contains($entry.Code))    { $codeOrder.Add($entry.Code) | Out-Null }
        if (-not $codeToLabel.ContainsKey($entry.Code)) { $codeToLabel[$entry.Code] = $entry.Ebene }
    }

    if (-not $matrixResult.Found) {
        $noMatrix.Add($page.title) | Out-Null
        Write-Verbose "  ! Keine Anwendungsebenen-Matrix auf '$($page.title)'"
    }

    # --- Tabelle 3: Schnittstellen ---
    $schnittstellen = @()
    if ($ExportSchnittstellen) {
        $schnittResult  = Get-SchnittstellenTabelle -Tables $tables
        $schnittstellen = $schnittResult.Rows
        if (-not $schnittResult.Found) {
            $noSchnitt.Add($page.title) | Out-Null
            Write-Verbose "  ! Keine Schnittstellen-Tabelle auf '$($page.title)'"
        }
    }

    # --- Tabelle 4: Sperrzeiten (Migration) ---
    $sperrzeiten = @()
    if ($ExportSperrzeiten) {
        $sperrResult = Get-SperrzeitenTabelle -Tables $tables
        $sperrzeiten = $sperrResult.Rows
        if (-not $sperrResult.Found) {
            $noSperr.Add($page.title) | Out-Null
            Write-Verbose "  ! Kein Sperrzeiten-Block auf '$($page.title)'"
        }
    }

    # --- Tabelle 5: Kapazitaetsanforderungen Server ---
    $server = @()
    if ($ExportServer) {
        $serverResult = Get-ServerTabelle -Tables $tables
        $server       = $serverResult.Rows
        if (-not $serverResult.Found) {
            $noServer.Add($page.title) | Out-Null
            Write-Verbose "  ! Keine Server-Tabelle auf '$($page.title)'"
        }
    }

    # --- Tabelle 6: Kapazitaetsanforderungen Datenbank ---
    $datenbanken = @()
    if ($ExportDatenbank) {
        $dbResult    = Get-DatenbankTabelle -Tables $tables
        $datenbanken = $dbResult.Rows
        if (-not $dbResult.Found) {
            $noDatenbank.Add($page.title) | Out-Null
            Write-Verbose "  ! Keine Datenbank-Tabelle auf '$($page.title)'"
        }
    }

    # --- Tabelle 7: Aufbau Anwendungsserver ---
    $aufbau = @()
    if ($ExportAufbau) {
        $aufbauResult = Get-KategorieTabelle -Tables $tables -Kategorien $AufbauKategorien -Titel $tableTitles -TitelFilter $AufbauTitel
        $aufbau       = $aufbauResult.Rows
        if ($aufbauResult.Found) {
            # Spaltennamen ueber alle Seiten vereinheitlichen
            for ($i = 0; $i -lt $aufbauResult.Spalten.Count; $i++) {
                if ($aufbauCols.Count -le $i) { $aufbauCols.Add($aufbauResult.Spalten[$i]) | Out-Null }
            }
        }
        else {
            $noAufbau.Add($page.title) | Out-Null
            Write-Verbose "  ! Keine Aufbau-Tabelle auf '$($page.title)'"
        }
    }

    # --- Tabelle 8: Umsysteme ---
    $umsysteme = @()
    if ($ExportUmsysteme) {
        $umsysResult = Get-KategorieTabelle -Tables $tables -Kategorien $UmsystemeKategorien -Titel $tableTitles -TitelFilter $UmsystemeTitel
        $umsysteme   = $umsysResult.Rows
        if ($umsysResult.Found) {
            for ($i = 0; $i -lt $umsysResult.Spalten.Count; $i++) {
                if ($umsysCols.Count -le $i) { $umsysCols.Add($umsysResult.Spalten[$i]) | Out-Null }
            }
        }
        else {
            $noUmsysteme.Add($page.title) | Out-Null
            Write-Verbose "  ! Keine Umsysteme-Tabelle auf '$($page.title)'"
        }
    }

    $webui = ''
    if ($page.PSObject.Properties.Name -contains '_links' -and $page._links -and
        $page._links.PSObject.Properties.Name -contains 'webui') {
        $webui = "$BaseUrl$($page._links.webui)"
    }

    $pageData.Add([pscustomobject]@{
        SeitenId     = $page.id
        Seitentitel  = $page.title
        Verfahren    = $verfahren
        Beschreibung = $beschreibung
        Weitere      = $weitere
        Url          = $webui
        Matrix       = $matrix
        MatrixMap      = $matrixMap
        ZusatzEbenen   = $zusatzEbenen
        Schnittstellen = $schnittstellen
        Sperrzeiten    = $sperrzeiten
        Server         = $server
        Datenbanken    = $datenbanken
        Aufbau         = $aufbau
        Umsysteme      = $umsysteme
    }) | Out-Null
}

Write-Progress -Activity 'Seiten werden ausgewertet' -Completed

# --- Ausgabe: eine Zeile je Seite, ein Wert je Spalte ---------------------

$rows = New-Object System.Collections.Generic.List[object]

foreach ($d in $pageData) {

    $record = [ordered]@{
        Seitenname   = $d.Seitentitel
        Verfahren    = $d.Verfahren
        Beschreibung = $d.Beschreibung
    }

    # Allgemeine Information und weitere Beschriftungszeilen
    foreach ($spalte in $WeitereFelder.Keys) {
        $record[$spalte] = $d.Weitere[$spalte]
    }

    # Je Anwendungsebene drei Spalten: IST, SOLL, Anzahl
    foreach ($code in $codeOrder) {
        $ist = ''; $soll = ''; $anzahl = ''
        if ($d.MatrixMap.ContainsKey($code)) {
            $entry  = $d.MatrixMap[$code]
            $ist    = $entry.IST
            $soll   = $entry.SOLL
            $anzahl = $entry.Anzahl
        }
        $record["${code}_IST"]    = $ist
        $record["${code}_SOLL"]   = $soll
        $record["${code}_Anzahl"] = $anzahl
    }

    if ($IncludeSeitenIdUndUrl) {
        $record['SeitenId'] = $d.SeitenId
        $record['Url']      = $d.Url
    }

    $rows.Add([pscustomobject]$record) | Out-Null
}

$writtenPath = Export-CsvUtf8Bom -Data $rows -Path $OutputPath

# --- Ausgabe Schnittstellen: eine Zeile je Schnittstelle -------------------

$schnittRows = New-Object System.Collections.Generic.List[object]
$schnittPath = $null

if ($ExportSchnittstellen) {

    foreach ($d in $pageData) {
        foreach ($s in $d.Schnittstellen) {
            $record = [ordered]@{
                Seitenname                 = $d.Seitentitel
                Verfahren                  = $d.Verfahren
                Nr                         = $s.Nr
                InputOutput                = $s.Richtung
                ZielQuellsystem            = $s.System
                ManuellHalbautomAutom      = $s.Art
                KurzbeschreibungDaten      = $s.Daten
            }
            if ($IncludeSeitenIdUndUrl) {
                $record['SeitenId'] = $d.SeitenId
                $record['Url']      = $d.Url
            }
            $schnittRows.Add([pscustomobject]$record) | Out-Null
        }
    }

    if ($schnittRows.Count -gt 0) {
        $schnittPath = Export-CsvUtf8Bom -Data $schnittRows -Path $OutputPathSchnittstellen
    }
    else {
        Write-Warning 'Keine Schnittstellen-Zeilen gefunden - die Schnittstellen-CSV wurde nicht geschrieben.'
    }
}

# --- Ausgabe Zusatzebenen: eine Zeile je zusaetzlicher Anwendungsebene -----

$zusatzRows = New-Object System.Collections.Generic.List[object]
$zusatzPath = $null

if ($ExportZusatzEbenen) {

    foreach ($d in $pageData) {
        foreach ($z in $d.ZusatzEbenen) {
            $record = [ordered]@{
                Seitenname      = $d.Seitentitel
                Verfahren       = $d.Verfahren
                Nr              = $z.Nr
                Anwendungsebene = $z.Ebene
                IST             = $z.IST
                SOLL            = $z.SOLL
                AnzahlIstServer = $z.Anzahl
            }
            if ($IncludeSeitenIdUndUrl) {
                $record['SeitenId'] = $d.SeitenId
                $record['Url']      = $d.Url
            }
            $zusatzRows.Add([pscustomobject]$record) | Out-Null
        }
    }

    if ($zusatzRows.Count -gt 0) {
        $zusatzPath = Export-CsvUtf8Bom -Data $zusatzRows -Path $OutputPathZusatzEbenen
    }
    else {
        Write-Warning 'Keine zusaetzlichen Anwendungsebenen gefunden - diese CSV wurde nicht geschrieben.'
    }
}

# --- Ausgabe Sperrzeiten: eine Zeile je Zeitraum ---------------------------

$sperrRows = New-Object System.Collections.Generic.List[object]
$sperrPath = $null

if ($ExportSperrzeiten) {

    foreach ($d in $pageData) {
        foreach ($s in $d.Sperrzeiten) {
            $record = [ordered]@{
                Seitenname   = $d.Seitentitel
                Verfahren    = $d.Verfahren
                Nr           = $s.Nr
                Von          = $s.Von
                Bis          = $s.Bis
                BemerkungName = $s.Bemerkung
            }
            if ($IncludeSeitenIdUndUrl) {
                $record['SeitenId'] = $d.SeitenId
                $record['Url']      = $d.Url
            }
            $sperrRows.Add([pscustomobject]$record) | Out-Null
        }
    }

    if ($sperrRows.Count -gt 0) {
        $sperrPath = Export-CsvUtf8Bom -Data $sperrRows -Path $OutputPathSperrzeiten
    }
    else {
        Write-Warning 'Keine Sperrzeiten gefunden - die Sperrzeiten-CSV wurde nicht geschrieben.'
    }
}

# --- Ausgabe Server: eine Zeile je Server ----------------------------------

$serverRows = New-Object System.Collections.Generic.List[object]
$serverPath = $null

if ($ExportServer) {

    foreach ($d in $pageData) {
        foreach ($sv in $d.Server) {
            $record = [ordered]@{
                Seitenname          = $d.Seitentitel
                Verfahren           = $d.Verfahren
                Nr                  = $sv.Nr
                Servername          = $sv.Servername
                Cluster             = $sv.Cluster
                Anwendungsebene     = $sv.Ebene
                BetriebssystemVersion = $sv.Betriebssystem
                CPU                 = $sv.CPU
                RAM                 = $sv.RAM
                FestplattenGB       = $sv.Festplatten
                DediziertVirtuell   = $sv.Art
            }
            if ($IncludeSeitenIdUndUrl) {
                $record['SeitenId'] = $d.SeitenId
                $record['Url']      = $d.Url
            }
            $serverRows.Add([pscustomobject]$record) | Out-Null
        }
    }

    if ($serverRows.Count -gt 0) {
        $serverPath = Export-CsvUtf8Bom -Data $serverRows -Path $OutputPathServer
    }
    else {
        Write-Warning 'Keine Server-Zeilen gefunden - die Server-CSV wurde nicht geschrieben.'
    }
}

# --- Ausgabe Datenbanken: eine Zeile je Datenbank --------------------------

$dbRows = New-Object System.Collections.Generic.List[object]
$dbPath = $null

if ($ExportDatenbank) {

    foreach ($d in $pageData) {
        foreach ($db in $d.Datenbanken) {
            $record = [ordered]@{
                Seitenname      = $d.Seitentitel
                Verfahren       = $d.Verfahren
                Nr              = $db.Nr
                Datenbank       = $db.Datenbank
                Anwendungsebene = $db.Ebene
                Groesse         = $db.Groesse
                Plattform       = $db.Plattform
            }
            if ($IncludeSeitenIdUndUrl) {
                $record['SeitenId'] = $d.SeitenId
                $record['Url']      = $d.Url
            }
            $dbRows.Add([pscustomobject]$record) | Out-Null
        }
    }

    if ($dbRows.Count -gt 0) {
        $dbPath = Export-CsvUtf8Bom -Data $dbRows -Path $OutputPathDatenbank
    }
    else {
        Write-Warning 'Keine Datenbank-Zeilen gefunden - die Datenbank-CSV wurde nicht geschrieben.'
    }
}

# --- Ausgabe der Kategorie-Tabellen (Aufbau, Umsysteme) --------------------

function Build-KategorieRows {
    # Baut die Zeilen einer Kategorie-Tabelle: Seitenname, Verfahren, Nr,
    # die Kategoriespalte und danach die gesammelten Wertespalten.
    param(
        [Parameter(Mandatory)]$PageData,
        [Parameter(Mandatory)][string]$Property,        # Feld in $pageData
        [Parameter(Mandatory)][string]$KategorieSpalte, # Ueberschrift der Kategoriespalte
        [Parameter(Mandatory)]$Spalten                  # gesammelte Wertespalten
    )

    # Spaltennamen eindeutig machen (Excel mag keine doppelten Kopfzeilen)
    $namen = New-Object System.Collections.Generic.List[string]
    foreach ($name in $Spalten) {
        $kandidat = $name
        $z = 2
        while ($namen.Contains($kandidat) -or $kandidat -eq $KategorieSpalte) { $kandidat = "$name$z"; $z++ }
        $namen.Add($kandidat) | Out-Null
    }

    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($d in $PageData) {
        foreach ($e in $d.$Property) {
            $record = [ordered]@{
                Seitenname = $d.Seitentitel
                Verfahren  = $d.Verfahren
                Nr         = $e.Nr
            }
            $record[$KategorieSpalte] = $e.Kategorie

            for ($i = 0; $i -lt $namen.Count; $i++) {
                $wert = ''
                if ($e.Werte.Count -gt $i) { $wert = $e.Werte[$i] }
                $record[$namen[$i]] = $wert
            }
            if ($IncludeSeitenIdUndUrl) {
                $record['SeitenId'] = $d.SeitenId
                $record['Url']      = $d.Url
            }
            $rows.Add([pscustomobject]$record) | Out-Null
        }
    }

    return [pscustomobject]@{ Rows = $rows.ToArray(); Spalten = $namen.ToArray() }
}

$aufbauRows   = @()
$aufbauPath   = $null
$aufbauNamen  = @()

if ($ExportAufbau) {
    $b = Build-KategorieRows -PageData $pageData -Property 'Aufbau' -KategorieSpalte 'Aufbau' -Spalten $aufbauCols
    $aufbauRows  = $b.Rows
    $aufbauNamen = $b.Spalten

    if ($aufbauRows.Count -gt 0) {
        $aufbauPath = Export-CsvUtf8Bom -Data $aufbauRows -Path $OutputPathAufbau
    }
    else {
        Write-Warning 'Keine Aufbau-Zeilen gefunden - die Aufbau-CSV wurde nicht geschrieben.'
    }
}

$umsysRows  = @()
$umsysPath  = $null
$umsysNamen = @()

if ($ExportUmsysteme) {
    $b = Build-KategorieRows -PageData $pageData -Property 'Umsysteme' -KategorieSpalte 'Umsysteme' -Spalten $umsysCols
    $umsysRows  = $b.Rows
    $umsysNamen = $b.Spalten

    if ($umsysRows.Count -gt 0) {
        $umsysPath = Export-CsvUtf8Bom -Data $umsysRows -Path $OutputPathUmsysteme
    }
    else {
        Write-Warning 'Keine Umsysteme-Zeilen gefunden - die Umsysteme-CSV wurde nicht geschrieben.'
    }
}

# --- Zusammenfassung ------------------------------------------------------

$spaltenzahl = 3 + $WeitereFelder.Count + ($codeOrder.Count * 3)
if ($IncludeSeitenIdUndUrl) { $spaltenzahl += 2 }

Write-Host ''
Write-Host "CSV geschrieben:  $writtenPath" -ForegroundColor Green
Write-Host "Zeilen (Seiten):  $($rows.Count)" -ForegroundColor Green
Write-Host "Spalten:          $spaltenzahl" -ForegroundColor Green
Write-Host "Anwendungsebenen: $($codeOrder -join ', ')" -ForegroundColor DarkGray

if ($schnittPath) {
    $seitenMitSchnitt = @($pageData | Where-Object { @($_.Schnittstellen).Count -gt 0 }).Count
    Write-Host ''
    Write-Host "CSV geschrieben:  $schnittPath" -ForegroundColor Green
    Write-Host "Schnittstellen:   $($schnittRows.Count) auf $seitenMitSchnitt Seite(n)" -ForegroundColor Green
}

if ($zusatzPath) {
    $seitenMitZusatz = @($pageData | Where-Object { @($_.ZusatzEbenen).Count -gt 0 }).Count
    Write-Host ''
    Write-Host "CSV geschrieben:  $zusatzPath" -ForegroundColor Green
    Write-Host "Zusatzebenen:     $($zusatzRows.Count) auf $seitenMitZusatz Seite(n)" -ForegroundColor Green
}

if ($sperrPath) {
    $seitenMitSperr = @($pageData | Where-Object { @($_.Sperrzeiten).Count -gt 0 }).Count
    Write-Host ''
    Write-Host "CSV geschrieben:  $sperrPath" -ForegroundColor Green
    Write-Host "Sperrzeiten:      $($sperrRows.Count) auf $seitenMitSperr Seite(n)" -ForegroundColor Green
}

if ($serverPath) {
    $seitenMitServer = @($pageData | Where-Object { @($_.Server).Count -gt 0 }).Count
    Write-Host ''
    Write-Host "CSV geschrieben:  $serverPath" -ForegroundColor Green
    Write-Host "Server:           $($serverRows.Count) auf $seitenMitServer Seite(n)" -ForegroundColor Green
}

if ($dbPath) {
    $seitenMitDb = @($pageData | Where-Object { @($_.Datenbanken).Count -gt 0 }).Count
    Write-Host ''
    Write-Host "CSV geschrieben:  $dbPath" -ForegroundColor Green
    Write-Host "Datenbanken:      $($dbRows.Count) auf $seitenMitDb Seite(n)" -ForegroundColor Green
}

if ($aufbauPath) {
    $seitenMitAufbau = @($pageData | Where-Object { @($_.Aufbau).Count -gt 0 }).Count
    Write-Host ''
    Write-Host "CSV geschrieben:  $aufbauPath" -ForegroundColor Green
    Write-Host "Aufbau-Zeilen:    $($aufbauRows.Count) auf $seitenMitAufbau Seite(n)" -ForegroundColor Green
    Write-Host "Wertespalten:     $($aufbauNamen -join ', ')" -ForegroundColor DarkGray
}

if ($umsysPath) {
    $seitenMitUmsys = @($pageData | Where-Object { @($_.Umsysteme).Count -gt 0 }).Count
    Write-Host ''
    Write-Host "CSV geschrieben:  $umsysPath" -ForegroundColor Green
    Write-Host "Umsysteme-Zeilen: $($umsysRows.Count) auf $seitenMitUmsys Seite(n)" -ForegroundColor Green
    Write-Host "Wertespalten:     $($umsysNamen -join ', ')" -ForegroundColor DarkGray
}

# Ebenen, die als Standard erkannt wurden, aber nicht in $Anwendungsebenen
# stehen - meist ein Zeichen dafuer, dass die Trennzeile anders formuliert ist
# und Sondereintraege faelschlich Spalten in der Haupt-CSV erzeugt haben.
$bekannt   = @($Anwendungsebenen | ForEach-Object { $_.ToUpperInvariant() })
$unbekannt = @($codeOrder | Where-Object { $bekannt -notcontains $_ })

if ($unbekannt.Count -gt 0) {
    Write-Warning "Nicht vorgesehene Anwendungsebenen als Spalten in der Haupt-CSV: $($unbekannt -join ', ')"
    Write-Warning 'Tipp: Stehen diese unter "Zusaetzlich benoetigte Anwendungsebenen"? Dann $MatrixZusatzLabel'
    Write-Warning '      an den Wortlaut auf der Seite anpassen. Sonst in $Anwendungsebenen aufnehmen.'
}

if ($noLabels.Count -gt 0) {
    Write-Warning "Auf $($noLabels.Count) Seite(n) fehlt die Beschriftungstabelle (Verfahren/Beschreibung):"
    $noLabels | Select-Object -First 10 | ForEach-Object { Write-Warning "  - $_" }
    if ($noLabels.Count -gt 10) { Write-Warning "  ... und $($noLabels.Count - 10) weitere." }
    Write-Warning 'Tipp: $NameLabel / $DescriptionLabel anpassen oder $UsePositionFallback = $true setzen.'
}

if ($noMatrix.Count -gt 0) {
    Write-Warning "Auf $($noMatrix.Count) Seite(n) fehlt die Anwendungsebenen-Matrix:"
    $noMatrix | Select-Object -First 10 | ForEach-Object { Write-Warning "  - $_" }
    if ($noMatrix.Count -gt 10) { Write-Warning "  ... und $($noMatrix.Count - 10) weitere." }
    Write-Warning 'Tipp: $MatrixHeaderLabel an den tatsaechlichen Text der ersten Kopfzelle anpassen.'
}

if ($ExportSchnittstellen -and $noSchnitt.Count -gt 0) {
    Write-Warning "Auf $($noSchnitt.Count) Seite(n) fehlt die Schnittstellen-Tabelle:"
    $noSchnitt | Select-Object -First 10 | ForEach-Object { Write-Warning "  - $_" }
    if ($noSchnitt.Count -gt 10) { Write-Warning "  ... und $($noSchnitt.Count - 10) weitere." }
    Write-Warning 'Hinweis: Seiten mit vorhandener, aber leerer Tabelle zaehlen hier nicht - die sind nur ohne Schnittstellen.'
}

if ($ExportSperrzeiten -and $noSperr.Count -gt 0) {
    Write-Warning "Auf $($noSperr.Count) Seite(n) fehlt der Sperrzeiten-Block:"
    $noSperr | Select-Object -First 10 | ForEach-Object { Write-Warning "  - $_" }
    if ($noSperr.Count -gt 10) { Write-Warning "  ... und $($noSperr.Count - 10) weitere." }
    Write-Warning 'Tipp: $SperrzeitenLabel pruefen; erkannt wird die Kopfzeile mit "von" und "bis".'
}

if ($ExportServer -and $noServer.Count -gt 0) {
    Write-Warning "Auf $($noServer.Count) Seite(n) fehlt die Server-Tabelle:"
    $noServer | Select-Object -First 10 | ForEach-Object { Write-Warning "  - $_" }
    if ($noServer.Count -gt 10) { Write-Warning "  ... und $($noServer.Count - 10) weitere." }
    Write-Warning 'Hinweis: Es muessen mindestens drei der acht Spaltenueberschriften erkannt werden.'
}

if ($ExportDatenbank -and $noDatenbank.Count -gt 0) {
    Write-Warning "Auf $($noDatenbank.Count) Seite(n) fehlt die Datenbank-Tabelle:"
    $noDatenbank | Select-Object -First 10 | ForEach-Object { Write-Warning "  - $_" }
    if ($noDatenbank.Count -gt 10) { Write-Warning "  ... und $($noDatenbank.Count - 10) weitere." }
    Write-Warning 'Hinweis: Die Spaltenueberschrift "Datenbank" muss vorhanden sein, dazu eine weitere.'
}

if ($ExportAufbau -and $noAufbau.Count -gt 0) {
    Write-Warning "Auf $($noAufbau.Count) Seite(n) fehlt die Aufbau-Tabelle:"
    $noAufbau | Select-Object -First 10 | ForEach-Object { Write-Warning "  - $_" }
    if ($noAufbau.Count -gt 10) { Write-Warning "  ... und $($noAufbau.Count - 10) weitere." }
    Write-Warning 'Hinweis: Mindestens zwei Kategorien aus $AufbauKategorien muessen in Spalte 1 stehen.'
}

if ($ExportUmsysteme -and $noUmsysteme.Count -gt 0) {
    Write-Warning "Auf $($noUmsysteme.Count) Seite(n) fehlt die Umsysteme-Tabelle:"
    $noUmsysteme | Select-Object -First 10 | ForEach-Object { Write-Warning "  - $_" }
    if ($noUmsysteme.Count -gt 10) { Write-Warning "  ... und $($noUmsysteme.Count - 10) weitere." }
    Write-Warning 'Hinweis: Mindestens zwei Kategorien aus $UmsystemeKategorien muessen in Spalte 1 stehen.'
}
