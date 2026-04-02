  // S.Oelze@knoll-online.com,         - KNOLL einsetzen wenn bereit.
  // Regensburg@wm.de,                 - WM einsetzen wenn bereit.
  // verkauf-autohero@stahlgruber.de,  - STA einsetzen wenn bereit.
  // Unten senden wir jetzt nur zu mir und dir, zum testen, alles andere ist erstmal rausgenommen bis alles bereit ist. Erst dann die mails oben, unten in die "var" einsetzen.
var EMPFAENGER = {
  "KNOLL":       "tobias.nehring@auto1.com,francesco.berger@auto1.com",
  "WM":          "tobias.nehring@auto1.com,francesco.berger@auto1.com",
  "Stahlgruber": "tobias.nehring@auto1.com,francesco.berger@auto1.com"
};

var BETREFF = "Bestellung KVA-Teile";

var BLATTNAME = "Lieferantenauftrag";

var EMAIL_TEXT = "Sehr geehrte Damen und Herren,\n\n"
  + "wir bitten um Bestellung der angekreuzten Teile wie im Anhang hinterlegt,\n"
  + "sowie eine Zusendung des Warenkorbs inkl. der Einkaufspreise\n"
  + "Bitte die günstigsten, vorhandenen Teile zusenden.\n"
  + "Bitte senden Sie die Rückantwort unter der Angabe der Stock ID an tobias.nehring@auto1.com\n\n"
  + "Mit freundlichen Grüßen\n\n"
  + "Autohero GmbH\n"
  + "Nordhausstraße 1\n"
  + "Hemau";


function zeigeVersandDialog() {
  var namen = Object.keys(EMPFAENGER);
  var buttons = "";
  for (var i = 0; i < namen.length; i++) {
    buttons += '<button class="btn" onclick="senden(\'' + namen[i] + '\')">'
             + namen[i] + '</button>\n';
  }

  var htmlString = ''
    + '<style>'
    + '  * { box-sizing: border-box; margin: 0; padding: 0; }'
    + '  html, body { height: 100%; width: 100%; }'
    + '  body { font-family: "Segoe UI", sans-serif; text-align: center;'
    + '         background: linear-gradient(180deg, #090d13 0%, #0d1117 100%); color: #c9d1d9;'
    + '         display: flex; align-items: center; justify-content: center; }'
    + '  .wrap { width: 100%; padding: 28px 24px; }'
    + '  h2 { margin-bottom: 6px; color: #f2cc60; font-size: 18px; font-weight: 800;'
    + '       text-transform: uppercase; letter-spacing: 1.2px; }'
    + '  p  { color: #8b949e; margin-bottom: 22px; font-size: 15px; }'
    + '  .btn { display: block; width: 84%; margin: 10px auto; padding: 15px 0;'
    + '         font-size: 14px; font-weight: 800; color: #0d1117;'
    + '         background: linear-gradient(180deg, #f2cc60 0%, #e3b341 100%);'
    + '         border: none; border-radius: 14px; cursor: pointer;'
    + '         text-transform: uppercase; letter-spacing: 1px;'
    + '         box-shadow: 0 8px 16px rgba(227, 179, 65, 0.2);'
    + '         transition: transform 0.2s, box-shadow 0.2s; }'
    + '  .btn:hover { transform: translateY(-1px); box-shadow: 0 10px 22px rgba(227, 179, 65, 0.3); }'
    + '  .btn:disabled { background: #21262d; color: #484f58; cursor: not-allowed;'
    + '                   box-shadow: none; transform: none; }'
    + '  #status { margin-top: 22px; font-size: 13px; font-weight: 700; display: none;'
    + '            padding: 10px 14px; border-radius: 999px; }'
    + '  #status.sending { background: rgba(88,166,255,0.08); border: 1px solid rgba(88,166,255,0.15);'
    + '                     color: #58a6ff; }'
    + '  #status.success { background: rgba(46,160,67,0.12); border: 1px solid rgba(63,185,80,0.3);'
    + '                     color: #56d364; }'
    + '  #status.error   { background: rgba(248,81,73,0.1); border: 1px solid rgba(248,81,73,0.3);'
    + '                     color: #f85149; }'
    + '  #status.cancelled { background: rgba(242,204,96,0.08); border: 1px solid rgba(242,204,96,0.2);'
    + '                       color: #f2cc60; }'
    + '  .spinner { display: inline-block; width: 16px; height: 16px;'
    + '             border: 2px solid rgba(242,204,96,0.22); border-top-color: #f2cc60;'
    + '             border-radius: 50%; animation: spin .72s linear infinite;'
    + '             vertical-align: middle; margin-right: 8px; }'
    + '  @keyframes spin { to { transform: rotate(360deg); } }'
    + '  #cancelBtn { display: none; margin: 14px auto 0; padding: 10px 28px;'
    + '               font-size: 12px; font-weight: 700; color: #f85149;'
    + '               background: rgba(248,81,73,0.08); border: 1px solid rgba(248,81,73,0.25);'
    + '               border-radius: 10px; cursor: pointer; text-transform: uppercase;'
    + '               letter-spacing: 0.8px; transition: all 0.2s; }'
    + '  #cancelBtn:hover { background: rgba(248,81,73,0.15); }'
    + '  #cancelBtn.too-late { color: #484f58; background: #21262d; border-color: #30363d;'
    + '                         cursor: not-allowed; }'
    + '  #cancelBtn.too-late:hover { background: #21262d; }'
    + '</style>'
    + '<div class="wrap">'
    + '<h2>Bestellung versenden</h2>'
    + '<p>An welchen Lieferanten soll die Bestellung gesendet werden?</p>'
    + buttons
    + '<div id="status"></div>'
    + '<button id="cancelBtn" onclick="abbrechen()">Abbrechen</button>'
    + '</div>'
    + '<script>'
    + '  var cancelled = false;'
    + '  var tempFileId = null;'
    + ''
    + '  function senden(name) {'
    + '    cancelled = false;'
    + '    var btns = document.querySelectorAll(".btn");'
    + '    for (var i = 0; i < btns.length; i++) { btns[i].disabled = true; btns[i].style.opacity = 0.5; }'
    + '    var s = document.getElementById("status");'
    + '    var cb = document.getElementById("cancelBtn");'
    + '    s.style.display = "block";'
    + '    s.className = "sending";'
    + '    s.innerHTML = \'<span class="spinner"></span> PDF wird vorbereitet...\';'
    + '    cb.style.display = "block";'
    + '    cb.className = "";'
    + '    cb.textContent = "Abbrechen";'
    + ''
    + '    google.script.run'
    + '      .withSuccessHandler(function(result) {'
    + '        tempFileId = result.fileId;'
    + '        if (cancelled) {'
    + '          google.script.run.tempDateiLoeschen(tempFileId);'
    + '          s.className = "cancelled";'
    + '          s.innerHTML = "⚠️ Abgebrochen — E-Mail wurde nicht gesendet.";'
    + '          cb.style.display = "none";'
    + '          setTimeout(function(){'
    + '            for (var i = 0; i < btns.length; i++) { btns[i].disabled = false; btns[i].style.opacity = 1; }'
    + '          }, 1000);'
    + '          return;'
    + '        }'
    + '        s.innerHTML = \'<span class="spinner"></span> Wird an \' + name + \' gesendet...\';'
    + '        cb.className = "too-late";'
    + '        cb.textContent = "Zu sp\\u00e4t";'
    + '        cb.onclick = null;'
    + ''
    + '        google.script.run'
    + '          .withSuccessHandler(function(n) {'
    + '            s.className = "success";'
    + '            s.innerHTML = "\\u2705 Bestellung erfolgreich an " + n + " gesendet!";'
    + '            cb.style.display = "none";'
    + '            setTimeout(function(){ google.script.host.close(); }, 2000);'
    + '          })'
    + '          .withFailureHandler(function(e) {'
    + '            s.className = "error";'
    + '            s.innerHTML = "\\u274c Fehler: " + e.message;'
    + '            cb.style.display = "none";'
    + '            for (var i = 0; i < btns.length; i++) { btns[i].disabled = false; btns[i].style.opacity = 1; }'
    + '          })'
    + '          .mailAbsenden(result.fileId, name);'
    + '      })'
    + '      .withFailureHandler(function(e) {'
    + '        s.className = "error";'
    + '        s.innerHTML = "\\u274c Fehler: " + e.message;'
    + '        cb.style.display = "none";'
    + '        for (var i = 0; i < btns.length; i++) { btns[i].disabled = false; btns[i].style.opacity = 1; }'
    + '      })'
    + '      .pdfVorbereiten(name);'
    + '  }'
    + ''
    + '  function abbrechen() {'
    + '    cancelled = true;'
    + '    var s = document.getElementById("status");'
    + '    var cb = document.getElementById("cancelBtn");'
    + '    s.className = "cancelled";'
    + '    s.innerHTML = "\\u23f3 Wird abgebrochen...";'
    + '    cb.className = "too-late";'
    + '    cb.textContent = "Wird abgebrochen...";'
    + '    cb.onclick = null;'
    + '  }'
    + '<\/script>';

  var html = HtmlService.createHtmlOutput(htmlString)
    .setWidth(420)
    .setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(html, "Bestellung versenden");
}


function pdfVorbereiten(name) {
  var empfaenger = EMPFAENGER[name];
  if (!empfaenger) {
    throw new Error("Unbekannter Lieferant: " + name);
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var blatt = ss.getSheetByName(BLATTNAME);

  var kopie = blatt.copyTo(ss);
  kopie.setName("_tmp_pdf_export");
  var drawings = kopie.getDrawings();
  for (var d = 0; d < drawings.length; d++) {
    drawings[d].remove();
  }
  SpreadsheetApp.flush();

  var url = ss.getUrl().replace(/\/edit.*$/, "")
    + "/export?format=pdf"
    + "&gid=" + kopie.getSheetId()
    + "&portrait=true"
    + "&size=A4"
    + "&fitw=true"
    + "&gridlines=false"
    + "&top_margin=0.3"
    + "&bottom_margin=0.3"
    + "&left_margin=0.3"
    + "&right_margin=0.3"
    + "&r1=0&c1=0"
    + "&r2=46&c2=15";

  var token    = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + token }
  });
  var pdfBlob = response.getBlob().setName(BLATTNAME + ".pdf");

  ss.deleteSheet(kopie);

  var tempFile = DriveApp.createFile(pdfBlob);
  return { fileId: tempFile.getId(), name: name };
}


function mailAbsenden(fileId, name) {
  var empfaenger = EMPFAENGER[name];
  if (!empfaenger) {
    throw new Error("Unbekannter Lieferant: " + name);
  }

  var file    = DriveApp.getFileById(fileId);
  var pdfBlob = file.getBlob().setName(BLATTNAME + ".pdf");

  MailApp.sendEmail({
    to: empfaenger,
    subject: BETREFF,
    body: EMAIL_TEXT,
    attachments: [pdfBlob]
  });

  file.setTrashed(true);
  return name;
}


function tempDateiLoeschen(fileId) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (e) {
  }
}
