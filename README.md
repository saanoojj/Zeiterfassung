# Pythonskript zur Erstellung einer Zeiterfassungsliste (output in .xlsx)

## Download for precompiled application

- macOS: https://drive.proton.me/urls/8SFQJ2EZ28#tRvsK9k9B3B8 (extract .zip) </br>
  compiled with: pyinstaller main.spec </br>
  Icon used: https://tinyurl.com/3hckak56 (if you are the righhtsowner and want it to be removed, please dont hesitate to contact me!)

### Dateiname

Einschub in Standardbenennung, notwendig um schon erstellte/verwendete nicht zu überschreiben

### Zuschlag

Momentan defekt/nicht funktionsfahig!

### Gehalt

Standard liegt bei 12.50€ ("Punkt" nicht "Komma" zur Trennung verwenden!)


### Bundesländer

Für die Erkennung der Feiertage wird das Module 'holidays' verwendet!

<details>
   <summary> Mögliche Optionen</summary>
</br>
   
   1. BB (Brandenburg)
   2. BE (Berlin)
   3. BW (Baden-Württemberg)
   4. BY (Bayern)
   5. HB (Bremen)
   6. HE (Hessen)
   7. HH (Hamburg)
   8. MV (Mecklenburg-Vorpommern)
   9. NI (Niedersachsen)
   10. NW (Nordrhein-Westfalen)
   11. RP (Rheinland-Pfalz)
   12. SH (Schleswig-Holstein)
   13. SL (Saarland)
   14. SN (Sachsen)
   15. ST (Sachsen-Anhalt)
   16. TH (Thüringen)

</details>

### Excel Datei

Die erstellte Excel-Datei enthält 13 'Data-Sheets'

   1. Das erste 'Data-Sheet' enthält eine Gesamtübersicht der geleisteten Stunden und des gesamt erwirtschafteten Lohn
   2. 'Data-Sheet' 2-13 enthält die Tabelle des jeweiligen Monats (Januar - Dezember)
   3. Das Skript geht (für die Feiertage) vom Bundesland "Sachsen" aus (hardcoded, bisher keine Clicks-Option)
   4. Feiertage sind gelb markiert


<details>
   <summary> To-Do's</summary>
</br>
   
   1. DONE! Clicks-Option zur Auswahl des Bundeslandes (für korrekte Eintragung Feiertage)
   2. DONE! Grundgehalt als feste Konstante auf 'Data-Sheet' 1 (zur einfachen Anpassung)
   3. --no click anymore-- Clicks-Option zur genauen Definition des Feiertagzuschlages & Wochenend- & Nachtzuschläge (entsprechende Clicks-Option)
   4. DONE! GUI für nutzerfreundlichere Bedienung (Custom TKinter)
      
   
</details>


