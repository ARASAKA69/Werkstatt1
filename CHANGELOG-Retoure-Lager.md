# Retoure Lager / Retoure Scan — Changelog

Projekt: `Retoure Lager` (Apps Script: App.js, Index.html = Retoure Lager, Aussen.html = Retoure Scan, carol-retoure-bridge.user.js = Tampermonkey)

Hinweis: Diese Datei liegt bewusst außerhalb des Ordners `Retoure Lager`, weil dort beim Apps-Script-Sync (clasp) alles gelöscht wird, was nicht .js/.html/.json ist.

## 1.2.14 — 2026-09-17
PDF nutzt die ganze Seite: die Karten liegen jetzt in einem Raster (3 Spalten auf A4, `auto-fill` ab 56 mm) statt einspaltig zentriert. Karten etwas kompakter, Barcode 26 mm hoch, Gruppen-Überschriften brechen nicht mehr vom Block ab.

## 1.2.13 — 2026-09-17
Der Carol-Auftrag wird jetzt immer geöffnet, auch wenn der Badge schon in der Trefferzeile der Suche steht — dadurch enthält der gespeicherte Badge-Text das Datum („Flagged for Return to Auto1 on 16 Sep 2026“ statt nur „Als B2A1 markiert“). Öffnen wird alle 12 Versuche erneut probiert; klappt es gar nicht, wird als Fallback der Badge aus der Trefferzeile gemeldet statt „kein Auftrag“. Tampermonkey 2.1.

## 1.2.12 — 2026-09-17
Falsches „kein Auftrag“ behoben: der Auftrag wird über die Auftrags-URL geöffnet (`location.assign`) statt über einen Klick, der in Carol ins Leere ging. „Kein Carol-Auftrag“ nur noch, wenn die Suche wirklich keine Zeile liefert. Der Carol-Button am Fahrzeug erzwingt eine erneute Prüfung (`retoure_force=1`). Tampermonkey 2.0.

## 1.2.11 — 2026-09-17
„Return to Auto1“ landet nicht mehr im Badge-Text — das ist der Aktions-Button; Buttons, Links und Menüs werden beim Badge-Lesen übersprungen. Im HUD steht statt BEHALTEN „Refurbishment gestartet“ in Violett, wenn Carol nur „Refurbishment Started“ zeigt; Liste, Detail und PDF nutzen dasselbe Label. Tampermonkey 1.9.

## 1.2.10 — 2026-09-17
Badges werden gezielt aus dem Carol-Badge-Container gelesen (`refurbishmentBadges`, `data-qa-selector*="efurbishmentStatus"`) statt aus der ganzen Seite. Sobald der Container da ist, wird sofort entschieden. Ohne B2A1/Fertig-Badge werden die gefundenen Badge-Texte gespeichert und im HUD gezeigt. Stock-ID aus dem Vehicle-Header. Tampermonkey 1.8.

## 1.2.9 — 2026-09-17
Listen-Status macht kein Fertiggestellt und kein B2A1 mehr — nur Carol-Badge, B2A1-Mail oder Nachbestellung. „Herausgegeben“ aus der Refurbishment List ist nur noch Info. Badge-Ergebnis getrennt vom Listen-Status gespeichert (Spalte „Carol fertig“ = B2A1 / FERTIG / NEIN plus Badge-Text), Aktion wird beim Laden neu berechnet, manuelle Auswahl bleibt. Jede Stock-ID wird in Carol geprüft. Tampermonkey 1.7.

## 1.2.8 — 2026-09-17
Loop behoben: die Stock-ID kommt nur noch aus der offenen Auftragsseite, nicht mehr aus sessionStorage — vorher wurde immer die erste ID erneut gemeldet und wieder geöffnet. Jede ID wird pro Dump nur einmal gemeldet. Stock-IDs ohne Carol-Auftrag (Reifen, Lagerbox) werden abgehakt statt endlos neu geöffnet. Tampermonkey 1.6 mit Doppelstart-Sperre und Versionsanzeige im Toast.

## 1.2.7 — 2026-09-17
Carol-Badges nur noch aus sichtbaren Chips (kein GraphQL). „Refurbishment Started“ ist kein B2A1. Report läuft auch ohne JSON-Antwort weiter.

## 1.2.6 — 2026-09-17
Copy-Button über der Stock-ID im Retoure Scan.

## 1.2.5 — 2026-09-16
Carol öffnet erst nach dem Check, nicht schon beim Eintippen der Stock-IDs.

## 1.2.4 — 2026-09-16
Badges werden erst auf der offenen Auftragsseite gelesen, nicht auf der Suchliste.

## 1.2.3 — 2026-09-16
Carol-Check läuft automatisch nach dem Dump: ein Fenster, nächste ID, Fenster zu am Ende, Ton wenn fertig.

## 1.2.2 — 2026-09-16
Carol-Badges (Als B2A1 markiert / Fertiggestellt) über Tampermonkey statt Refurbishment-Listen-Spalte.

## 1.2.1 — 2026-09-16
Carol Retoure Bridge Userscript, Queue- und Report-Endpoints (`?page=carolq`, `?page=carolreport`).

## 1.2.0 — 2026-09-16
Retoure Scan Inventur: Dump aller Stock-IDs (Reifen + Lagerbox), B2A1-Mail, Nachbestellung, Tagesliste, PDF B2A1 + Fertig. Reifen Kontrolle unverändert.
