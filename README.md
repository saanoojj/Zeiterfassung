# Pythonskript zur Erstellung einer Zeiterfassungsliste (output in .xlsx)

## Clicks Optionen

- "-d" oder "--dateiname" fügt nachfolgenden Text in den Dateinamen (Zeiterfassung_*_2024.xlsx) ein
- "-z" oder "--zuschlag" aktiviert die Feiertagszuschlagsrechung (Standardwert bei 50%, hardcoded)
- "-g" oder "--gehalt" verändert den Stundenlohn entsprechend, Standard liegt bei 12,50€

<details>
<summary>Beispiele für die Verwendung (klicken zum Ausklappen)</summary>

### Standardverwendung
```bash
Zeiterfassung.exe
```
Erstellt eine Excel-Datei mit Standardstundenlohn (12,50€) ohne Feiertagszuschlag

### Mit angepasstem Stundenlohn
```bash
Zeiterfassung.exe --gehalt 14.50
```
Verwendet 14,50€ als Stundenlohn

### Mit Feiertagszuschlag
```bash
Zeiterfassung.exe --zuschlag
```
Aktiviert die Feiertagszuschlagsberechnung (50%)

### Mit eigenem Dateinamen
```bash
Zeiterfassung.exe --dateiname MeineZeiterfassung
```
Erstellt "Zeitplan_MeineZeiterfassung_2024.xlsx"

### Alle Optionen kombiniert
```bash
Zeiterfassung.exe -d MeineZeiterfassung -g 14.50 -z
```
Verwendet alle verfügbaren Optionen
</details>

## Lizenz

[Lizenzinformationen hier]