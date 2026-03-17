- [x] Kommentar + Regal als ein Flow: Kommentar rot im Sheet, Status auf Teilweise angeliefert, Regal setzen
- [x] Kommentar- und Regal-Buttons zu einem gemeinsamen Button zusammenführen für unvollständige Aufträge
- [x] Kommentar beim gemeinsamen Regal-Speichern automatisch mitspeichern
- [x] Wortfilter: Öl + WM + KNOLL + STA + ALFAH ebenfalls grün markieren
- [x] Kleinen Button oben links bauen, der alle Stock-IDs mit Regalplatz sortiert anzeigt
- [ ] Reifen annahme system, erkennt das tabellenblatt nicht, bzw. nicht die StockID wenn ein leerzeichen im namen vom Tabellenblatt ist.
- [ ] Add WSS + VA + HA + SENG or Seng in die Wort filter Liste die Grün markiert werden im HUD
- [ ] Kommentar sollte auch OHNE ein Regalplatz auszuwählen abgespeichert werden können, jedoch eine kleine info angezeigt werden mit OK das kein Regal vergeben worden ist.
- [ ] Wenn nachbestellung Exit Bestellung dann Exit liste ebenfalls auf Komplett ANgeliefert markieren im dropdown wenn auf Angeliefert im HUD gesetzt wird
- [ ] Link zum sheet --> https://docs.google.com/spreadsheets/d/1OrSRkB8xdMk0uGvTGUVA_J8Q3IPX0GYF7eOXf6af1GI/edit?gid=1722963315#gid=1722963315
- [ ] Nachbestellung, Pfeile beim Etiketten druck Grösser machen wo man Stückzahl von etiketten ändern kann auch bei allen anderen feldern wo man etiketten drucken kann
- [ ] Übersicht + Reifen Annahme + Nachbestellungen Button als button Cards einrichten (Bigger)
- [ ] Add settings button top right, only with login ARASAKA (random short pass) to setup word filters, which words are in filters and or remove them/ add as seperate section inside Settings.
- [ ] Add Language change inside Settings and make sure everything inside the HUB get translated, maybe we use Google Translate for this case, and making sure also the 2 first cards with values input get Translated. For now oonly German and English. German is Main language, english will use the translator before load in all data of any order.
- [ ] Add descripütion on Kommentar + Regal Speichern button over it (Nur Verwenden wenn Auftrag nicht fertig und eingelagert werden muss !)
- [ ] Move (Inside Regal xx are x Boxes) badge in the middle
- [ ] Make Regal x.x Dropdown font and Aktuell: Regal x.x badge bigger to be more present and visible.
- [ ] Move Auftrag Beenden & Carold Starten Button Centered also Vertically inside its card.
- [ ] Inside Übersicht Popup adding Copy button for the StockID beside of Im WMS öffnen button
- [ ] When set any Nachbestellkung into Teilweise angelierfert, it should be able to set a Regal place to store this first box.
- [ ] Nachbestellungen is filtering out everything not only the open orders, we need only the open orders in the list nit the ended oney, ofc. when i set one to delivered it should not be removed by itself cause i still need the infos. maybe we add a refresh button ontop to filter them out too when done.
- [ ] Also its better to add a minimize button which ends in a small toolbar on the bottom to make sure i can reopen it and see all from before again, without need to re enter it and wait for new filtering of all the orders. Only Status bestellt inside the Dropdown.

- [ ] When hit delivered (in my case) i need to use my other google sheet code to auto implement the stock id and the description into my Werkstattauftrag, this way i dont have to place it manually (for now only for myself cause i have it in my own sheet, the others are using 1 for all which is useless then. (CODE FOR THE AUTO APPLY OF STOCK AND DESCRIPTION NEEDED FROM GOOGLE APPS SCRIPTS)

```ruby
const ZIEL_TABELLEN_ID = "1nE6SErc1-jmZYd_Ydviw28Pa5qdJmwNepXCiVbsdsVo";
const ZIEL_TABELLENBLATT_NAME = "BLANCO Reparaturauftrag";
const MEINE_EMAIL = "francesco.berger@auto1.com";

function autoFillAuftrag(e) {
  if (!e || !e.range) return;
  
  var bearbeiter = Session.getActiveUser().getEmail();
  
  if (bearbeiter !== MEINE_EMAIL) return;
  
  var sheet = e.range.getSheet();
  if (sheet.getName() !== "Nachbestellungen") return;
  
  var col = e.range.getColumn();
  var row = e.range.getRow();
  
  if (col === 12 && row > 1) {
    var newValue = String(e.range.getValue());
    
    if (newValue.indexOf("Angeliefert/Bereit für Tagesliste") !== -1) {
      var stockId = sheet.getRange(row, 2).getValue();
      var beschreibung = sheet.getRange(row, 6).getValue();
      
      try {
        var zielSs = SpreadsheetApp.openById(ZIEL_TABELLEN_ID);
        var zielSheet = zielSs.getSheetByName(ZIEL_TABELLENBLATT_NAME);
        
        if (zielSheet) {
          zielSheet.getRange("D10").setValue(stockId);
          zielSheet.getRange("D18").setValue(beschreibung);
        }
      } catch (err) {
      }
    }
  }
}

```

- [ ] It should be also able to search for the Artikel number not only the stock, which shows me the part number row with its stock id, so change search bar placeholder text then too STOCK-ID SUCHEN... into STOCK-ID & Artikel Nummer Suchen...
- [ ] In reifen Annahme we need a feature to search for not set to JA tires which i can only search by its size eg. 255/60 R18 and then i get a list with all the tires from all tables inside that sheet with this size and stock id + stückzahl + Lastindex	+ GW Index, these tires always have "per Paketdienst 2-3 Werktage" inside the row from xx stockid which we can filter out first all with this hint then filter out all which are already delivered then we show only the ones which are not delivered so set to JA and there should be a button to like use this tires to bookin and after that we do the same as usually, print stockid xx times, set all ready inside the other sheets table and set the tires table to green on that line and on Ja for delivered. this way i dont even need to enter the sheet for searching only size tires. + we need to have Sheet date on that info to know which one we see in the found list of tires and stockids.
- [ ] make Reifen annahme popup bigger also its font inside to be better visible and readable.
- [ ] Inside the main HUD where we scan fr stock id. we should also be able to search for the order number, they are always inside "Kommentar Ersatzteile Bestellung" well not always, but most times they are inside there, this way we can give in eg. N4P1793025, ofc. also should be possible without the N4P so only the number itserlf, this way we search this number, he scans the sheet on that row and when found he uses this line with xx stockid and shows me the stockid results. that way the step going into emails and check for the stock id is then no longer needed. Also for extern orders like alfah we have the order number eg. 1196247 and then we should be able to search it by this number, it should pick out the line where this number is inside (it exist only  time inside the sheet on "Kommentar Ersatzteile Bestellung" then we work with this lines stockid and continue the main steps like when we load in a strockid
- [ ] 
-----

Check daily list table for tires which need to get out on that shift. maybe another new button for Check Tires for today (maybe Heutige Reifenausgabe) we only show the current date of these open tires and the oney from next day. the rest are not important, inside row A they all have a date so we should be able to filter them out easly, and then we scan this table:
Inside the dropdown we only show the "ffe5a0 and ffe8aa" colors these are the two options which need to giving out for this shift. if they have the other dropdown colors it should be ignored because that means they are already gaved out.

<img width="564" height="994" alt="image" src="https://github.com/user-attachments/assets/f4d0803f-6f2f-41a9-a018-8c8e69d2a528" />

Exit Table example for generating code using the correct columns and rows.:
Link to it:
https://docs.google.com/spreadsheets/d/1OrSRkB8xdMk0uGvTGUVA_J8Q3IPX0GYF7eOXf6af1GI/edit?gid=1722963315#gid=1722963315

<img width="1641" height="518" alt="image" src="https://github.com/user-attachments/assets/26cc0482-4dd3-4507-aac5-b8223b3bca4e" />
