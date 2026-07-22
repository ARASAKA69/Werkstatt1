# N4Parts StockID Suche – Installation (Chrome)

Kurz und knackig: Tampermonkey rein, Script drauf, fertig. Dann kannst du auf der Warenkörbe-Seite (`#/cart`) nach StockID suchen ohne durch 11 Seiten zu klicken.

## 1. Tampermonkey in Chrome installieren

1. Chrome öffnen
2. Hier den Extension-Store aufmachen:  
   https://chromewebstore.google.com/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo
3. **„Hinzufügen“ / „Add to Chrome“** klicken
4. Bestätigen → Extension ist drin

Falls Chrome meckert wegen „Entwicklermodus“ oder so: einfach erlauben, sonst läuft das Script nicht.

## 2. Unser Script installieren

1. Diesen Link **im Chrome** öffnen (nicht runterladen und speichern, einfach anklicken):  
   https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/N4Parts/n4parts-stockid-suche.user.js
2. Tampermonkey poppt auf → **Installieren** klicken
3. Fertig

Alternativ: Tampermonkey-Icon → Dashboard → „+“ / neues Script → Inhalt vom Link reinpasten → speichern. Aber der Direktlink oben ist easier.

## 3. Nutzen

1. Bei N4Parts einloggen
2. Warenkörbe öffnen: https://www.n4parts.net/#/cart
3. Oben links (oder wo du das Fenster hingeschoben hast): **StockID suchen**
4. StockID eintippen (z.B. `ZG11817`) → **Suchen** / Enter
5. Script sucht über alle Seiten und öffnet den Warenkorb

Fenster lässt sich am dunklen Titelbalken verschieben – Position bleibt gespeichert.

## Tipps / Trouble

- Panel nur sichtbar auf **Warenkörbe** (`#/cart`), nicht auf Bestellungen oder woanders
- Script nicht da? Seite neu laden, prüfen ob Tampermonkey aktiv ist (Icon → Script enabled)
- Update: Tampermonkey synced manchmal selbst; sonst nochmal den Install-Link öffnen und updaten
- Läuft nur wenn du bei N4Parts eingeloggt bist (nutzt deine Session)

by ARASAKA
