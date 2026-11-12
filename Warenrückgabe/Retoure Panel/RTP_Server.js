function doGet() {
  if (!rtpAllowed_()) {
    return HtmlService.createHtmlOutput("<p style=\"font-family:sans-serif;padding:24px\">Kein Zugriff.</p>").setTitle("Retoure Panel");
  }
  var t = HtmlService.createTemplateFromFile("RTP_App");
  t.appVersion = RTP_VERSION;
  t.appIcon = Utilities.base64Encode(RTP_ICON_SVG);
  return t
    .evaluate()
    .setTitle("Retoure Panel")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function rtpUser_() {
  try {
    return String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  } catch (e) {
    return "";
  }
}

function rtpEffectiveUser_() {
  try {
    return String(Session.getEffectiveUser().getEmail() || "").trim().toLowerCase();
  } catch (e) {
    return "";
  }
}

function rtpDefaultWho_() {
  var u = rtpUser_();
  if (!u) return "";
  var local = u.split("@")[0] || "";
  if (local === "ersatzteile.hemau") return "";
  var parts = local.split(".");
  var last = parts[parts.length - 1] || local;
  if (!last) return "";
  return last.charAt(0).toUpperCase() + last.slice(1);
}

function rtpAllowed_() {
  var u = rtpUser_();
  if (!u) return false;
  var lists = RTP_ALLOWED.concat(RTP_PROTECT_EDITORS);
  for (var i = 0; i < lists.length; i++) {
    if (String(lists[i]).trim().toLowerCase() === u) return true;
  }
  return rtpSheetEditors_().indexOf(u) >= 0;
}

function rtpSheetEditors_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get("rtp_editors");
  if (hit) return JSON.parse(hit);
  var list = [];
  try {
    var ss = rtpSS_();
    list = ss.getEditors().map(function (x) { return String(x.getEmail() || "").trim().toLowerCase(); });
    var owner = ss.getOwner();
    if (owner) list.push(String(owner.getEmail() || "").trim().toLowerCase());
    list = list.filter(Boolean);
    cache.put("rtp_editors", JSON.stringify(list), 600);
  } catch (e) {
    list = [];
  }
  return list;
}

function rtpSS_() {
  return SpreadsheetApp.openById(RTP_SHEET_ID);
}

function rtpTabConfig_(key) {
  for (var i = 0; i < RTP_TABS.length; i++) {
    if (RTP_TABS[i].key === key) return RTP_TABS[i];
  }
  return null;
}

function rtpFindSheet_(ss, cfg) {
  if (!cfg) return null;
  if (cfg.preferName) {
    var byNameFirst = ss.getSheetByName(cfg.sheet);
    if (byNameFirst) return byNameFirst;
  }
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === cfg.gid) return sheets[i];
  }
  return ss.getSheetByName(cfg.sheet);
}

function rtpNorm_(s) {
  return String(s == null ? "" : s)
    .replace(/\u00a0/g, " ")
    .replace(/[\u200b-\u200d\ufeff]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function rtpIsDate_(v) {
  return Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime());
}

function rtpToMs_(dateVal, timeVal) {
  var baseMs = 0;
  if (rtpIsDate_(dateVal)) {
    baseMs = dateVal.getTime();
  } else if (typeof dateVal === "number" && dateVal > 20000 && dateVal < 80000) {
    baseMs = Math.round((dateVal - 25569) * 864e5);
  } else {
    var s = String(dateVal == null ? "" : dateVal).trim();
    var m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (m) {
      var y = Number(m[3]);
      if (y < 100) y += 2000;
      var hh = Number(m[4] || 0);
      var mm = Number(m[5] || 0);
      var tm = String(timeVal == null ? "" : timeVal).trim().match(/^(\d{1,2}):(\d{2})/);
      if (!m[4] && tm) {
        hh = Number(tm[1]);
        mm = Number(tm[2]);
      }
      baseMs = new Date(y, Number(m[2]) - 1, Number(m[1]), hh, mm, Number(m[6] || 0)).getTime();
    }
  }
  if (baseMs && rtpIsDate_(dateVal) && timeVal) {
    var tm2 = String(timeVal == null ? "" : timeVal).trim().match(/^(\d{1,2}):(\d{2})/);
    if (tm2) {
      var d = new Date(baseMs);
      d.setHours(Number(tm2[1]), Number(tm2[2]), 0, 0);
      baseMs = d.getTime();
    }
  }
  return baseMs;
}

function rtpDayIndex_(ms) {
  if (!ms) return null;
  var k = Utilities.formatDate(new Date(ms), RTP_TZ, "yyyy-MM-dd").split("-");
  return Math.floor(Date.UTC(Number(k[0]), Number(k[1]) - 1, Number(k[2])) / 864e5);
}

function rtpDaysSince_(ms, todayIdx) {
  var d = rtpDayIndex_(ms);
  return d == null ? null : todayIdx - d;
}

function rtpChecked_(value) {
  if (value === true) return true;
  var t = String(value == null ? "" : value).trim().toUpperCase();
  return t === "TRUE" || t === "WAHR" || t === "1";
}

function rtpGmailUrl_(hash) {
  var target = "https://mail.google.com/mail/u/0/?authuser=" + encodeURIComponent(RTP_INBOX) + hash;
  return "https://accounts.google.com/AccountChooser?Email=" + encodeURIComponent(RTP_INBOX) + "&continue=" + encodeURIComponent(target);
}

function rtpThreadUrl_(threadId) {
  var id = String(threadId || "").trim();
  if (!id) return "";
  return rtpGmailUrl_("#all/" + id);
}

function rtpSearchUrl_(rNumber) {
  var r = String(rNumber || "").trim().toUpperCase();
  if (!r) return "";
  var q = encodeURIComponent("\"" + r + "\"");
  return rtpGmailUrl_("#search/" + q);
}

function rtpRichLink_(rich) {
  if (!rich) return "";
  var url = "";
  try { url = String(rich.getLinkUrl() || ""); } catch (e) { url = ""; }
  if (url) return url;
  try {
    var runs = rich.getRuns();
    for (var i = 0; i < runs.length; i++) {
      var u = runs[i].getLinkUrl();
      if (u) return String(u);
    }
  } catch (e2) {}
  return "";
}

function rtpFormulaUrl_(formula) {
  var f = String(formula || "");
  var m = f.match(/HYPERLINK\s*\(\s*"([^"]+)"/i);
  return m ? m[1] : "";
}

function rtpClassifyAlfah_(status, info, bgHex, tabKey) {
  var st = rtpNorm_(status);
  var inf = rtpNorm_(info);
  var bg = String(bgHex || "").trim().toLowerCase();
  if (tabKey === "abgelehnt" || st === "abgelehnt" || bg === RTP_ALFAH_RED) return "abgelehnt";
  if (st === "storno") return "storno";
  if (tabKey === "geliefert") return "geliefert";
  if (bg === RTP_ALFAH_GREEN) return "bereit";
  if (bg === RTP_ALFAH_ORANGE) return "veranlasst";
  if (bg === RTP_ALFAH_YELLOW || st === "vorbereitet") return "vorbereitet_mail";
  if (st === "rücksendung" || st === "rucksendung" || st === "rücksendung eingeleitet") {
    if (inf.indexOf("veranlasst") >= 0 || inf.indexOf("eingeleitet") >= 0) return "veranlasst";
    return "bereit";
  }
  if (!st) return "offen";
  return "sonstige";
}

function rtpReadAlfahTab_(ss, cfg, todayIdx) {
  var out = {
    key: cfg.key,
    kind: cfg.kind,
    label: cfg.label,
    code: cfg.code,
    sheet: cfg.sheet,
    color: cfg.color,
    archive: !!cfg.archive,
    rows: []
  };
  var sh = rtpFindSheet_(ss, cfg);
  if (!sh) {
    out.missing = true;
    return out;
  }
  var last = sh.getLastRow();
  if (last < 3) return out;
  var count = last - 2;
  var width = Math.max(16, Math.min(sh.getLastColumn(), 16));
  var rng = sh.getRange(3, 1, count, width);
  var vals = rng.getValues();
  var disp = rng.getDisplayValues();
  var bgs = sh.getRange(3, 1, count, 1).getBackgrounds();
  var richLinks = sh.getRange(3, 13, count, 2).getRichTextValues();
  var formulas = sh.getRange(3, 13, count, 2).getFormulas();
  for (var i = 0; i < count; i++) {
    var d = disp[i];
    var orderId = String(d[0] || "").trim();
    var stockId = String(d[1] || "").trim().toUpperCase();
    var rNumber = String(d[2] || "").trim().toUpperCase();
    var article = String(d[3] || "").trim();
    var name = String(d[4] || "").trim();
    if (!orderId && !stockId && !rNumber && !article && !name) continue;
    var supplier = String(d[5] || "").trim();
    var brand = String(d[6] || "").trim();
    var date = String(d[7] || "").trim();
    var time = String(d[8] || "").trim();
    var who = String(d[9] || "").trim();
    var status = String(d[10] || "").trim();
    var info = String(d[11] || "").trim();
    var mailUrl = rtpRichLink_(richLinks[i][0]) || rtpFormulaUrl_(formulas[i][0]);
    var searchUrl = rtpRichLink_(richLinks[i][1]) || rtpFormulaUrl_(formulas[i][1]) || (rNumber ? rtpSearchUrl_(rNumber) : "");
    var vorbereitet = rtpChecked_(vals[i][14]);
    var geliefert = rtpChecked_(vals[i][15]);
    var bg = String(bgs[i][0] || "").trim().toLowerCase();
    var ms = rtpToMs_(vals[i][7], d[8]);
    out.rows.push({
      row: 3 + i,
      tab: cfg.key,
      kind: "alfah",
      orderId: orderId,
      stockId: stockId,
      rNumber: rNumber,
      article: article,
      name: name,
      supplier: supplier,
      brand: brand,
      date: date,
      time: time,
      who: who,
      status: status,
      info: info,
      mailUrl: mailUrl,
      searchUrl: searchUrl,
      vorbereitet: vorbereitet,
      geliefert: geliefert,
      bg: bg,
      state: rtpClassifyAlfah_(status, info, bg, cfg.key),
      datumMs: ms,
      daysOpen: rtpDaysSince_(ms, todayIdx)
    });
  }
  return out;
}

function rtpReadSupplierTab_(ss, cfg, todayIdx) {
  var out = {
    key: cfg.key,
    kind: cfg.kind,
    label: cfg.label,
    code: cfg.code,
    sheet: cfg.sheet,
    color: cfg.color,
    defaultSupplier: cfg.defaultSupplier || cfg.label,
    rows: []
  };
  var sh = rtpFindSheet_(ss, cfg);
  if (!sh) {
    out.missing = true;
    return out;
  }
  var last = sh.getLastRow();
  if (last < 3) return out;
  var count = last - 2;
  var width = Math.max(10, Math.min(sh.getLastColumn(), 10));
  var rng = sh.getRange(3, 1, count, width);
  var vals = rng.getValues();
  var disp = rng.getDisplayValues();
  for (var i = 0; i < count; i++) {
    var d = disp[i];
    var orderId = String(d[0] || "").trim();
    var stockId = String(d[1] || "").trim().toUpperCase();
    var article = String(d[2] || "").trim();
    var name = String(d[3] || "").trim();
    if (!orderId && !stockId && !article && !name) continue;
    var supplier = String(d[4] || "").trim();
    var brand = String(d[5] || "").trim();
    var date = String(d[6] || "").trim();
    var time = String(d[7] || "").trim();
    var who = String(d[8] || "").trim();
    var info = String(d[9] || "").trim();
    var ms = rtpToMs_(vals[i][6], d[7]);
    out.rows.push({
      row: 3 + i,
      tab: cfg.key,
      kind: "supplier",
      orderId: orderId,
      stockId: stockId,
      rNumber: "",
      article: article,
      name: name,
      supplier: supplier,
      brand: brand,
      date: date,
      time: time,
      who: who,
      status: "In der Liste",
      info: info,
      mailUrl: "",
      searchUrl: "",
      vorbereitet: true,
      geliefert: false,
      bg: "",
      state: "done",
      datumMs: ms,
      daysOpen: rtpDaysSince_(ms, todayIdx)
    });
  }
  return out;
}

function getPanelData() {
  if (!rtpAllowed_()) return { ok: false, message: "Kein Zugriff." };
  try {
    var ss = rtpSS_();
    var now = new Date();
    var todayIdx = rtpDayIndex_(now.getTime());
    var tabs = {};
    var tabOrder = [];
    var suppliersMap = {};
    var brandsMap = {};
    var whoMap = {};
    for (var i = 0; i < RTP_TABS.length; i++) {
      var cfg = RTP_TABS[i];
      var data = cfg.kind === "alfah"
        ? rtpReadAlfahTab_(ss, cfg, todayIdx)
        : rtpReadSupplierTab_(ss, cfg, todayIdx);
      tabs[cfg.key] = data;
      tabOrder.push(cfg.key);
      for (var r = 0; r < data.rows.length; r++) {
        var row = data.rows[r];
        if (row.supplier) suppliersMap[row.supplier] = true;
        if (row.brand) brandsMap[row.brand] = true;
        if (row.who && !/^\d+$/.test(row.who)) whoMap[row.who] = true;
      }
    }
    var defWho = rtpDefaultWho_();
    if (defWho) whoMap[defWho] = true;
    return {
      ok: true,
      version: RTP_VERSION,
      user: rtpUser_(),
      defaultWho: defWho,
      effectiveUser: rtpEffectiveUser_(),
      loadedAt: Utilities.formatDate(now, RTP_TZ, "dd.MM.yyyy HH:mm"),
      loadedClock: Utilities.formatDate(now, RTP_TZ, "HH:mm:ss"),
      sheetUrl: ss.getUrl(),
      tabOrder: tabOrder,
      tabs: tabs,
      lists: {
        suppliers: Object.keys(suppliersMap).sort(function (a, b) { return a.localeCompare(b, "de"); }),
        brands: Object.keys(brandsMap).sort(function (a, b) { return a.localeCompare(b, "de"); }),
        workers: Object.keys(whoMap).sort(function (a, b) { return a.localeCompare(b, "de"); })
      }
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

