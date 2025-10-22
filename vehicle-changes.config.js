
/*
===============================================================
  Anleitung: Fahrzeug-Änderungen eintragen 
===============================================================
- Eine Änderung pro Zeile.
- Leere Zeilen sind OK.
- Zeilen, die mit # beginnen, sind Kommentare (werden ignoriert).

Format je Zeile:
  <Datum> <Aktion> <Fahrzeug-ID> [Fahrzeugname]

Erlaubte Datumsformate:
  - DD.MM.YYYY   z. B. 09.03.2025
  - YYYY-MM-DD   z. B. 2025-03-09

Erlaubte Aktionen:
  - "hinzufügen" oder "add"     → Fahrzeug ab diesem Datum in die Liste aufnehmen
  - "entfernen"  oder "remove"  → Fahrzeug ab diesem Datum aus der Liste entfernen

Beispiel für entfernen:
01.01.2025 entfernen 222

Beispiel für hinzufügen:
01.01.2025 hinzufügen 222 Opel Zafira
oder:
01.01.2025 add 222 Opel Zafira

Hinweise:
  - Fahrzeug-ID: nur Zahlen (2–4 Ziffern). Führende Nullen werden entfernt.
  - Der Fahrzeugname ist optional und wird nur bei "hinzufügen" genutzt.
  - Ungültige/fehlerhafte Zeilen werden ignoriert.
*/

window.VEHICLE_CHANGES = `













































`;

