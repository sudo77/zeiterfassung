# Konfiguration: Pfad zur CSV-Datei (anpassen!)
$csvPath = "C:\zeiterfassung\Zeiterfassung.csv"

# Pause in Minuten (anpassen!)
$pause = 30

# Aktuelles Datum und Zeit
$heute = Get-Date -Format "dd.MM.yyyy"
$jetzt = Get-Date

# Falls die CSV noch nicht existiert, lege sie mit Header an.
if (-not (Test-Path $csvPath)) {
    "Start_____Zeitpunkt;End_____Zeitpunkt;SummeTag;SummeWoche;KW" | Out-File -FilePath $csvPath -Encoding UTF8
}

# CSV einlesen
$data = Import-Csv -Path $csvPath -Delimiter ';'

# Funktion zur Formatierung der Zeitspanne in Stunden:Minuten:Sekunden
function Format-Dauer($timespan) {
    $hours = $timespan.Hours + ($timespan.Days * 24)  # Berücksichtigt die Tage in Stunden
    $minutes = $timespan.Minutes
    $seconds = $timespan.Seconds
    return "{0:D2}:{1:D2}:{2:D2}" -f $hours, $minutes, $seconds
}

# Funktion, um ISO-Wochenzahl zu ermitteln (ISO 8601: Montag als erster Tag der Woche)
function Get-IsoWeekNumber([datetime]$date) {
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    return $culture.Calendar.GetWeekOfYear($date, [System.Globalization.CalendarWeekRule]::FirstFourDayWeek, [DayOfWeek]::Monday)
}

# Funktion zur Umwandlung der Zeit (hh:mm:ss) in Sekunden
function ConvertToSeconds($timeString) {
    $parts = $timeString.Split(":")
    if ($parts.Length -eq 3) {
        return [int]$parts[0] * 3600 + [int]$parts[1] * 60 + [int]$parts[2]
    }
    return 0
}

# Funktion zur Umwandlung von Sekunden in hh:mm:ss
function Format-DauerInSeconds($seconds) {
    $ts = [TimeSpan]::FromSeconds($seconds)
    $hours = $ts.Hours + ($ts.Days * 24)  # Berücksichtigt die Tage in Stunden
    $minutes = $ts.Minutes
    $seconds = $ts.Seconds
    return "{0:D2}:{1:D2}:{2:D2}" -f $hours, $minutes, $seconds
}

# Sucht nach einem offenen Eintrag (ohne Ende) für heute
$offenerEintrag = $data | Where-Object { $_.Start_____Zeitpunkt -like "$heute*" -and ([string]::IsNullOrWhiteSpace($_.End_____Zeitpunkt)) }

if ($offenerEintrag) {
    # Bestehender Eintrag für heute schließen
    $startString = $offenerEintrag.Start_____Zeitpunkt  # z.B. "23.02.2025 08:30:15"
    try {
        $startDateTime = [datetime]::ParseExact($startString, "dd.MM.yyyy HH:mm:ss", $null)
    }
    catch {
        Write-Error "Fehler beim Parsen des Start-Zeitpunkts: $startString"
        exit 1
    }
    $endeZeit = $jetzt
    $differenz = $endeZeit - $startDateTime

    # Subtrahiere die Pause von der Gesamtdauer (Pause wird in Minuten abgezogen und in TimeSpan umgerechnet)
    $differenz = $differenz - [TimeSpan]::FromMinutes($pause)

    # Überprüfung, ob die berechnete Dauer negativ ist
    if ($differenz.TotalSeconds -lt 0) {
        Write-Warning "Die berechnete Zeitspanne ist negativ, daher wird sie ignoriert."
        $summeTag = "00:00:00"  # Setze auf 0, wenn die Zeitspanne negativ ist
    }
    else {
        # Umrechnung in Stunden:Minuten:Sekunden
        $summeTag = Format-Dauer($differenz)
    }

    # Aktualisieren des Eintrags: Endzeit und SummeTag eintragen
    $data | ForEach-Object {
        if ($_.Start_____Zeitpunkt -eq $startString -and ([string]::IsNullOrWhiteSpace($_.End_____Zeitpunkt))) {
            $_.End_____Zeitpunkt = $endeZeit.ToString("dd.MM.yyyy HH:mm:ss")
            $_.SummeTag = $summeTag
        }
    }

    # Um KW zu berechnen und hinzufügen
    foreach ($entry in $data) {
        if ($entry.Start_____Zeitpunkt -ne $null -and $entry.Start_____Zeitpunkt -match "\d{2}\.\d{2}\.\d{4}") {
            $entryDate = [datetime]::ParseExact($entry.Start_____Zeitpunkt.Substring(0,10), "dd.MM.yyyy", $null)
            $entry | Add-Member -NotePropertyName KW -NotePropertyValue (Get-IsoWeekNumber($entryDate)) -Force
        }
    }

    # Berechnung der SummeWoche nach KW
    $groups = $data | Group-Object -Property KW

    foreach ($grp in $groups) {
        $totalSeconds = 0
        foreach ($item in $grp.Group) {
            if (-not [string]::IsNullOrWhiteSpace($item.SummeTag)) {
                $totalSeconds += ConvertToSeconds($item.SummeTag)
            }
        }

        # Gesamtzeit als TimeSpan (in Stunden, Minuten, Sekunden)
        $summeWocheStr = Format-DauerInSeconds($totalSeconds)

        # Aktualisieren der SummeWoche für alle Einträge der Gruppe (gleiche KW)
        foreach ($item in $grp.Group) {
            $item.SummeWoche = $summeWocheStr
        }
    }

    # Schreibe alle Daten zurück in die CSV (ohne zusätzlichen Header)
    $data | Select-Object Start_____Zeitpunkt, End_____Zeitpunkt, SummeTag, SummeWoche, KW | Export-Csv -Path $csvPath -NoTypeInformation -Delimiter ';'

    Write-Output "Session beendet: Ende = $($endeZeit.ToString("dd.MM.yyyy HH:mm:ss")); SummeTag = $summeTag"
    Write-Output "Wöchentliche Summe aktualisiert."
}
else {
    # Neuer Eintrag: Session starten
    $startZeit = $jetzt.ToString("dd.MM.yyyy HH:mm:ss")
    $newEntry = [PSCustomObject]@{
        Start_____Zeitpunkt = $startZeit
        End_____Zeitpunkt   = ""
        SummeTag            = ""
        SummeWoche          = ""
        KW                  = ""
    }

    # Füge den neuen Eintrag hinzu
    $newEntry | Export-Csv -Path $csvPath -Append -NoTypeInformation -Delimiter ';'

    Write-Output "Session gestartet: $startZeit"
}
