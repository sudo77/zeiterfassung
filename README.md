# Zeiterfassung

Dieses Projekt enthält ein PowerShell-Skript zur Zeiterfassung, das die Arbeitszeiten eines Benutzers aufzeichnet und in einer CSV-Datei speichert.

## Dateien

- `timecollector.ps1`: Das Hauptskript zur Zeiterfassung.
- `Zeiterfassung.csv`: Die CSV-Datei, in der die erfassten Zeiten gespeichert werden.
- `zeiterfassung.xml`: Die XML-Datei zur Konfiguration des Task Schedulers.

## Installation

1. Klonen Sie das Repository oder laden Sie die Dateien herunter.
2. Stellen Sie sicher, dass PowerShell auf Ihrem System installiert ist.

## Verwendung

### Starten einer neuen Session

Führen Sie das Skript `timecollector.ps1` aus, um eine neue Session zu starten. Das Skript erstellt einen neuen Eintrag in der CSV-Datei mit dem aktuellen Startzeitpunkt.

```powershell
powershell -File C:\zeiterfassung\timecollector.ps1
```

### Beenden einer bestehenden Session

Führen Sie das Skript erneut aus, um die aktuelle Session zu beenden. Das Skript aktualisiert den Eintrag in der CSV-Datei mit dem Endzeitpunkt und berechnet die Dauer der Session.

```powershell
powershell -File C:\zeiterfassung\timecollector.ps1
```

## CSV-Datei

Die CSV-Datei `Zeiterfassung.csv` enthält die folgenden Spalten:

- `Start_____Zeitpunkt`: Der Startzeitpunkt der Session.
- `End_____Zeitpunkt`: Der Endzeitpunkt der Session.
- `SummeTag`: Die Dauer der Session in Stunden, Minuten und Sekunden.
- `SummeWoche`: Die kumulierte Dauer aller Sessions in der aktuellen Woche.
- `KW`: Die Kalenderwoche des Eintrags.

## Funktionen im Skript

### Format-Dauer

Formatiert eine `TimeSpan`-Objekt in das Format `hh:mm:ss`.

```powershell
function Format-Dauer($timespan) {
    $hours = $timespan.Hours + ($timespan.Days * 24)
    $minutes = $timespan.Minutes
    $seconds = $timespan.Seconds
    return "{0:D2}:{1:D2}:{2:D2}" -f $hours, $minutes, $seconds
}
```

### Get-IsoWeekNumber

Ermittelt die ISO-Wochenzahl (ISO 8601: Montag als erster Tag der Woche).

```powershell
function Get-IsoWeekNumber([datetime]$date) {
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    return $culture.Calendar.GetWeekOfYear($date, [System.Globalization.CalendarWeekRule]::FirstFourDayWeek, [DayOfWeek]::Monday)
}
```

### ConvertToSeconds

Konvertiert eine Zeitangabe im Format `hh:mm:ss` in Sekunden.

```powershell
function ConvertToSeconds($timeString) {
    $parts = $timeString.Split(":")
    if ($parts.Length -eq 3) {
        return [int]$parts[0] * 3600 + [int]$parts[1] * 60 + [int]$parts[2]
    }
    return 0
}
```

### Format-DauerInSeconds

Konvertiert eine Zeitangabe in Sekunden in das Format `hh:mm:ss`.

```powershell
function Format-DauerInSeconds($seconds) {
    $ts = [TimeSpan]::FromSeconds($seconds)
    $hours = $ts.Hours + ($ts.Days * 24)
    $minutes = $ts.Minutes
    $seconds = $ts.Seconds
    return "{0:D2}:{1:D2}:{2:D2}" -f $hours, $minutes, $seconds
}
```

## Task Scheduler

Um das Skript automatisch beim Anmelden und Abmelden auszuführen, können Sie den Task Scheduler verwenden.

### Task Scheduler Konfiguration

1. Öffnen Sie den Task Scheduler (`taskschd.msc`).
2. Klicken Sie auf `Aktion` > `Importieren...`.
3. Wählen Sie die Datei `zeiterfassung.xml` aus und klicken Sie auf `Öffnen`.
4. Passen Sie die Einstellungen nach Bedarf an und klicken Sie auf `OK`.

Die XML-Datei `zeiterfassung.xml` enthält die Konfiguration für den Task Scheduler, um das Skript `timecollector.ps1` beim Anmelden und Abmelden auszuführen.

## Lizenz

Dieses Projekt ist unter der MIT-Lizenz lizenziert. Weitere Informationen finden Sie in der [LICENSE](LICENSE)-Datei.
