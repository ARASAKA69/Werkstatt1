# Erledigt

- [x] Kommentar + Regal als ein Flow: Kommentar rot im Sheet, Status auf Teilweise angeliefert, Regal setzen
- [x] Kommentar- und Regal-Buttons zu einem gemeinsamen Button zusammenführen für unvollständige Aufträge
- [x] Kommentar beim gemeinsamen Regal-Speichern automatisch mitspeichern
- [x] Wortfilter: Öl + WM + KNOLL + STA + ALFAH ebenfalls grün markieren
- [x] Kleinen Button oben links bauen, der alle Stock-IDs mit Regalplatz sortiert anzeigt
- [x] Reifen annahme system, erkennt das tabellenblatt nicht, bzw. nicht die StockID wenn ein leerzeichen im namen vom Tabellenblatt ist.
- [x] Add WSS + VA + HA + SENG or Seng in die Wort filter Liste die Grün markiert werden im HUD
- [x] Kommentar sollte auch OHNE ein Regalplatz auszuwählen abgespeichert werden können, jedoch eine kleine info angezeigt werden mit OK das kein Regal vergeben worden ist.
- [x] Wenn nachbestellung Exit Bestellung dann Exit liste ebenfalls auf Komplett ANgeliefert markieren im dropdown wenn auf Angeliefert im HUD gesetzt wird
- [x] Link zum sheet --> https://docs.google.com/spreadsheets/d/1OrSRkB8xdMk0uGvTGUVA_J8Q3IPX0GYF7eOXf6af1GI/edit?gid=1722963315#gid=1722963315
- [x] Nachbestellung, Pfeile beim Etiketten druck Grösser machen wo man Stückzahl von etiketten ändern kann auch bei allen anderen feldern wo man etiketten drucken kann
- [x] Übersicht + Reifen Annahme + Nachbestellungen Button als button Cards einrichten (Bigger)
- [x] Add descripütion on Kommentar + Regal Speichern button over it (Nur Verwenden wenn Auftrag nicht fertig und eingelagert werden muss !)
- [x] Move (Inside Regal xx are x Boxes) badge in the middle
- [x] Make Regal x.x Dropdown font and Aktuell: Regal x.x badge bigger to be more present and visible.
- [x] Move Auftrag Beenden & Carold Starten Button Centered also Vertically inside its card.
- [x] Inside Übersicht Popup adding Copy button for the StockID beside of Im WMS öffnen button
- [x] Update Regal Places inside dropdown automatic every 2 seconds, when the dropdown is not in use. to make sure the places are updated even someone uses any place, without affecting the dropdown while choosing a place.
- [x] Auto add the Nachbestellung Beschreibung and StockID based on my Script down below, into my Werkstattauftrag.
- [x] Nachbestellungen is filtering out everything not only the open orders, we need only the open orders in the list nit the ended oney, ofc. when i set one to delivered it should not be removed by itself cause i still need the infos. maybe we add a refresh button ontop to filter them out too when done.
- [x] Also its better to add a minimize button which ends in a small toolbar on the bottom to make sure i can reopen it, or just minimize into the button which means when click the button aslong the session is active i can reopen itt without need to load all orders from that sheet again until i dont hit the refresh button manually, and see all from before again, without need to re enter it and wait for new filtering of all the orders. Only Status bestellt inside the Dropdown.
- [x] Reifen Annahme: Minimize + Refresh Button (gleiche Logik wie Nachbestellungen)
- [x] When hit delivered (in my case) i need to use my other google sheet code to auto implement the stock id and the description into my Werkstattauftrag, this way i dont have to place it manually (for now only for myself cause i have it in my own sheet, the others are using 1 for all which is useless then.
- [x] It should be also able to search for the Artikel number not only the stock, which shows me the part number row with its stock id, so change search bar placeholder text then too STOCK-ID SUCHEN... into STOCK-ID & Artikel Nummer Suchen...
- [x] Inside the main HUD where we scan fr stock id. we should also be able to search for the order number, they are always inside "Kommentar Ersatzteile Bestellung" well not always, but most times they are inside there, this way we can give in eg. N4P1793025, ofc. also should be possible without the N4P so only the number itserlf, this way we search this number, he scans the sheet on that row and when found he uses this line with xx stockid and shows me the stockid results. that way the step going into emails and check for the stock id is then no longer needed. Also for extern orders like alfah we have the order number eg. 1196247 and then we should be able to search it by this number, it should pick out the line where this number is inside (it exist only  time inside the sheet on "Kommentar Ersatzteile Bestellung" then we work with this lines stockid and continue the main steps like when we load in a strockid
- [x] make Reifen annahme popup bigger also its font inside to be better visible and readable.
- [x] In reifen Annahme we need a feature to search for not set to JA tires which i can also only search by its size eg. 255/60 R18 and then i get a list with all the tires from all tables inside that sheet with this size and stock id + stückzahl + Lastindex + GW Index, these tires always have "per Paketdienst 2-3 Werktage" inside the row from xx stockid which we can filter out first all with this hint then filter out all which are already delivered then we show only the ones which are not delivered so set to JA and there should be a button to like use this tires to bookin and after that we do the same as usually, print stockid xx times, set all ready inside the other sheets table and set the tires table to green on that line and on Ja for delivered. this way i dont even need to enter the sheet for searching only size tires. + we need to have Sheet date on that info to know which one we see in the found list of tires and stockids.
- [x] Nachbestellungen list was cutting off at 04.03.2026 even tho we had like 2 more weeks of entries on the sheet. turned out the sort was buggy af, the return 0 fallback was messing up the order so newer rows werent showing up on top. fixed it by using row number as tiebreaker so now all open ones show up properly from latest to oldest.
- [x] Copy button for the description in Nachbestellungen wasnt copying anything, only the stockid one worked. fixed it with a proper clipboard fallback and now both boxes flash green with a Kopiert! feedback so u actually know it copied.
- [x] Made the Nachbestellungen detail sections color coded so its not all the same boring look. Kopieren section is now purple, Etiketten section golden and centered, and the Status box goes red when its nicht bestellt and green when angeliefert. also updates live when u change the dropdown no refresh needed.
- [x] Built the Heutige Reifenausgabe feature, new button in topbar that opens a modal showing only todays and tomorrows tires from the Tagesliste sheet. only shows entries with 2 Reifen or 4 Reifen dropdown (the yellow/orange ones), ignores gestellt/fehlbestand/lagerbestand since those are already handled. color coded badges for 2 vs 4 tires, dates tagged as Heute/Morgen, yes/no columns green/red. minimize + refresh just like the other modals.

-----

# Offen


- [ ] Add Language change inside Settings and make sure everything inside the HUD get translated, maybe we use Google Translate for this case, and making sure also the 2 first cards with values input get Translated. For now oonly German and English. German is Main language, english will use the translator before load in all data of any order.
- [ ] When set any Nachbestellkung into Teilweise angelierfert, it should be able to set a Regal place to store this first box.
- [ ] Reifen annahme, after click buchen and popup for checking the order, i click yes and it just stays at reifne annahme window and dont sends me into the HUD for checking the StockID automatically as it should do.
- [ ] The regal counting is not correctly, because we also have boxes from Tagesliste/Nachbestellungen which are stored inside a regal and these boxes dont get counted inside the system, we need to take care of these ones too.
- [ ] .
