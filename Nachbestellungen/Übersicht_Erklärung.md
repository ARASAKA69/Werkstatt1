### Übersicht Nachbestellung-Rote Hütchen (Code kurz erklärt)

Das Skript ist die **Schaltzentrale für die Werkstatt-Nachbestellungen**. Es sorgt dafür, dass die verschiedenen Tabellenblätter
(Input-Sheets, Nachbestellung, Dashboard und Archiv) automatisch miteinander sprechen, damit keiner Daten doppelt eintippen muss.

---


### Die wichtigsten Funktionen im Detail:

#### 1. Automatisches Erzeugen von Entry IDs und Datum
Sobald jemand in den Input-Sheets (`Input Q-Check`, `Input Mechanik`, `Input Lack`, `Input Exit`) eine neue Stock-ID einträgt:
* Generiert das Skript im Hintergrund eine eindeutige ID (UUID) in der Spalte `entry_id` (Spalte 27 bzw. AA).
* Setzt automatisch das aktuelle Datum in Spalte A, falls das noch leer war.

#### 2. Nachbestellungen automatisch anlegen & löschen
* Wenn im Input-Sheet in Spalte C ein **"ja"** steht (und die Zeile ausgefüllt ist), wandert der Eintrag automatisch rüber ins Sheet `Nachbestellung`.
* Steht da **"nein"** oder **"diagnose"**, wird die Zeile aus `Nachbestellung` wieder gelöscht.

#### 3. Zwei-Wege-Synchronisation (NB <--> Input Exit)
* **NB -> Exit**: Wenn im Sheet `Nachbestellung` bei "Exit Mechanik"-Zeilen Preise, Cost Gate, Carol-Status oder der Status geändert werden, schreibt das Skript diese Infos automatisch rüber ins Sheet `Input Exit`.
* **Exit -> NB**: Wenn im Sheet `Input Exit` bestimmte Felder geändert werden, spiegelt das Skript die Werte zurück in die `Nachbestellung`. (Bisher nur Fertiggestellt, B2A1 & nicht notwendig) NB -> Input Exit hat immer Priorität.
* **Ziel-Sheet**: Sobald in `Input Exit` in Spalte M ein Wert eingetragen wird, wird dieser direkt in das externe Ziel-Sheet `Repair Status` geschoben.

#### 4. Die Event-Warteschlange (EventQueue)
* Damit das Sheet nicht abstürzt oder sperrt, wenn mehrere Leute gleichzeitig arbeiten, wirft das Skript alle Änderungen erst mal in den Reiter `EventQueue`.
* Im Hintergrund läuft regelmäßig ein Trigger, der diese Warteschlange abarbeitet und das Dashboard updatet. (Deshalb dauert alles auch etwas länger, sheets ist einfach langsamer als ein eigener server damit muss man leben)

#### 5. Dashboard & Automatisches Archivieren
* Das **Dashboard** zeigt alle offenen Fälle an und berechnet die Tage seit dem letzten Eintrag.
* **Archivieren**: Sobald du auf dem Dashboard in der Spalte "archivieren" den Haken setzt, verschiebt das Skript den Fall automatisch aus allen aktiven Reitern (Dashboard, Nachbestellung, Inputs) in die jeweiligen Archiv-Reiter (z.B. `Dashboard_Archiv`, `Nachbestellung_Archiv`, etc.).

#### 6. Schutz der Stock-ID
* Das Skript sperrt die Bearbeitung der Stock-ID für normale User, damit keiner aus Versehen IDs zerschießt.

#### 7. Kaskadierendes Löschen
* Wenn eine Stock-ID gelöscht wird, sorgt das Skript dafür, dass dieser Eintrag überall (Dashboard, Nachbestellung und Inputs) verschwindet und eventuell noch wartende Aktionen in der Queue abgebrochen werden.

---