# Erledigt

- [x] StockID `NA.jpg` / NA-Pattern → Kommentar als **Nachbestellung 1/1**, **2/2**, usw. je nach wie viele Bilder in der Ladung sind (Carol Comment).
- [x] StockID + `StockID (2)` Kram → **Ausgabe 1/1**, **2/2** usw. im Comment, passt sich an die Anzahl Bilder an.
- [x] **Retoure**-Dateien (`…Retoure…`) → **Retoure 1/1**, **2/2** usw. im Comment.
- [x] Duplicate-Thema: gleicher Dateiname / gleicher Zeitstempel wie letzter erfolgreicher Upload → skip statt doppelt hochladen; Batch in Drive nimmt bei gleichem Namen die **neueste** Datei (mtime), Rest fliegt raus.
- [x] Carol-Seite: wenn der **Dateiname schon in der Liste** steht, aber die Datei in Drive **neuer** ist als der letzte gespeicherte Upload → trotzdem hochladen (Late Shift / anderer Kollege); nur wenn wirklich **gleicher** Timestamp wie gespeichert → skip (kein Fake-Duplikat mehr).
- [x] Logs fetter gemacht: skipped counts, warum skipped, moveFile-Zeilen; Summary bei Skips auch im Ordner **Kisten Duplicate Uploads** + normales Log bleibt wo es hingehört.
- [x] Bridge läuft per **POST + JSON** (kein HTML-Fehler mehr von zu langen URLs), `getBatch` schickt **modifiedTime** + **lastUploadedStored** fürs Timestamp-Vergleichen.
- [x] HUD ohne Tech-Sprech für User — alles Detaillierte nur noch in der **Konsole** wenn `ARASAKA_DEBUG` auf true; default false damit’s clean bleibt.
- [x] Fixing Duplicate issue, when any of these uploaded with the same name, it shpuld skip them.
- [ ] 
# Offen

