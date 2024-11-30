# Pythonskript zur Erstellung einer Zeiterfassungsliste (output in .xlsx)

## Clicks Optionen

- "-d" oder "--dateiname" fügt nachfolgenden Text in den Dateinamen (Zeiterfassung_*_2024.xlsx) ein
- "-z" oder "--zuschlag" aktiviert die Feiertagszuschlagsrechung (Standardwert bei 50%, hardcoded)
- "-g" oder "--gehalt" verändert den Stundenlohn entsprechend, Standard liegt bei 12,50€

<details>
<summary>Beispiele für die Verwendung (klicken zum Ausklappen)</summary>

### Standardverwendung
```bash
py main.py
```
Erstellt eine Excel-Datei mit Standardstundenlohn (12,50€) ohne Feiertagszuschlag

### Mit angepasstem Stundenlohn
```bash
py main.py --gehalt 14.50
```

```bash
py main.py -g 14.50
```
Verwendet 14,50€ als Stundenlohn

### Mit Feiertagszuschlag
```bash
py main.py --zuschlag
```

```bash
py main.py -z
```
Aktiviert die Feiertagszuschlagsberechnung (50%)

### Mit eigenem Dateinamen
```bash
py main.py --dateiname MeineZeiterfassung
```

```bash
py main.py -d MeineZeiterfassung
```
Erstellt "Zeitplan_MeineZeiterfassung_2024.xlsx"

### Alle Optionen kombiniert
```bash
py main.py -d MeineZeiterfassung -g 14.50 -z
```
Verwendet alle verfügbaren Optionen
</details>

