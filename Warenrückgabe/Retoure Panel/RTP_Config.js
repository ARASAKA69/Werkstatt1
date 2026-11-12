var RTP_VERSION = "1.1.0";
var RTP_ICON_SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><defs><linearGradient id="bg" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#0f172a"/><stop offset="0.55" stop-color="#131b36"/><stop offset="1" stop-color="#1e1538"/></linearGradient><linearGradient id="gm" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#38bdf8"/><stop offset="0.5" stop-color="#818cf8"/><stop offset="1" stop-color="#e879a8"/></linearGradient><linearGradient id="gl" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#60a5fa" stop-opacity=".32"/><stop offset="1" stop-color="#a855f7" stop-opacity=".08"/></linearGradient></defs><rect width="64" height="64" rx="16" fill="url(#bg)"/><rect x="1.25" y="1.25" width="61.5" height="61.5" rx="14.75" fill="none" stroke="url(#gm)" stroke-width="2.2" stroke-opacity=".85"/><path d="M30 14 46 22.5v15L30 46 14 37.5v-15Z" fill="url(#gl)"/><path d="M30 14 46 22.5v15L30 46 14 37.5v-15Z" fill="none" stroke="url(#gm)" stroke-width="2.4" stroke-linejoin="round"/><path d="M14 22.5 30 31l16-8.5M30 31v15M21.8 18.3l16 8.5" fill="none" stroke="#e2e8f0" stroke-width="2.2" stroke-linejoin="round" stroke-linecap="round" stroke-opacity=".9"/><circle cx="46.5" cy="46.5" r="11" fill="#11182d" stroke="url(#gm)" stroke-width="2.2"/><path d="M50.2 43.2a4.9 4.9 0 1 0 .8 5.4" fill="none" stroke="#f8fafc" stroke-width="2.2" stroke-linecap="round"/><path d="M50.8 40.3 51.4 44l-3.6.7" fill="none" stroke="#38bdf8" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/><path d="M49 11.5c.5 2.6 1.4 3.5 4 4-2.6.5-3.5 1.4-4 4-.5-2.6-1.4-3.5-4-4 2.6-.5 3.5-1.4 4-4Z" fill="url(#gm)"/></svg>';

var RTP_SHEET_ID = "1P8s15IFPyQb3OFkjW__YZZ9Q1-Qui9b4Y5k4RJMJtpc";
var RTP_TZ = "Europe/Berlin";
var RTP_INBOX = "ersatzteile.hemau@autohero.com";
var RTP_LOCK_MS = 20000;

var RTP_ALLOWED = [
  "francesco.berger@auto1.com",
  "ersatzteile.hemau@autohero.com"
];

var RTP_PROTECT_EDITORS = [
  "francesco.berger@auto1.com",
  "bastian.frankhauser@auto1.com",
  "joerg.eichenseher@auto1.com",
  "tobias.nehring@auto1.com",
  "helmut.mirbeth@auto1.com"
];

var RTP_ALFAH_BLUE = "#3c78d8";
var RTP_ALFAH_GREEN = "#00b050";
var RTP_ALFAH_ORANGE = "#e69138";
var RTP_ALFAH_YELLOW = "#f1c232";
var RTP_ALFAH_RED = "#cc0000";
var RTP_WM_GREEN = "#18884a";
var RTP_AAG_NAVY = "#00244e";
var RTP_STA_BLUE = "#1d4ed8";
var RTP_FONT = 8;

var RTP_ALFAH_HEADERS = [
  "Auftrag", "Stock-ID", "R-Nummer", "Teilenummer", "Name",
  "Lieferant", "Hersteller", "Datum", "Uhrzeit", "Wer",
  "Status", "Info", "Mail", "Suche", "Vorbereitet", "Geliefert"
];
var RTP_ALFAH_WIDTHS = [60, 62, 64, 74, 100, 84, 68, 58, 38, 48, 80, 70, 36, 40, 48, 44];

var RTP_SUPPLIER_HEADERS = [
  "Auftrag", "Stock-ID", "Teilenummer", "Name", "Lieferant",
  "Hersteller", "Datum", "Uhrzeit", "Wer", "Info"
];
var RTP_SUPPLIER_WIDTHS = [72, 78, 132, 120, 72, 84, 92, 76, 72, 96, 92];
var RTP_STA_WIDTHS = [110, 90, 140, 180, 90, 110, 92, 76, 90, 160, 92];

var RTP_TABS = [
  {
    key: "alfah",
    kind: "alfah",
    gid: 786084326,
    sheet: "Alfah",
    title: "Alfah Retouren",
    label: "Alfah",
    code: "AL",
    color: RTP_ALFAH_BLUE,
    archive: false
  },
  {
    key: "geliefert",
    kind: "alfah",
    gid: 1870949588,
    sheet: "Alfah Geliefert",
    title: "Alfah Geliefert",
    label: "Alfah Geliefert",
    code: "GL",
    color: RTP_ALFAH_GREEN,
    archive: true
  },
  {
    key: "abgelehnt",
    kind: "alfah",
    gid: 1505373902,
    sheet: "Alfah Abgelehnt",
    title: "Alfah Abgelehnt",
    label: "Alfah Abgelehnt",
    code: "AB",
    color: RTP_ALFAH_RED,
    archive: true
  },
  {
    key: "wm",
    kind: "supplier",
    gid: 891953845,
    sheet: "WM",
    title: "WM Retouren",
    label: "WM",
    code: "WM",
    color: RTP_WM_GREEN,
    defaultSupplier: "WM",
    widths: RTP_SUPPLIER_WIDTHS
  },
  {
    key: "aag",
    kind: "supplier",
    gid: 0,
    sheet: "KNOLL / AAG",
    title: "KNOLL / AAG Retouren",
    label: "KNOLL / AAG",
    code: "AG",
    color: RTP_AAG_NAVY,
    defaultSupplier: "AAG",
    preferName: true,
    widths: RTP_SUPPLIER_WIDTHS
  },
  {
    key: "sta",
    kind: "supplier",
    gid: 441953307,
    sheet: "STA",
    title: "STA Retouren",
    label: "STA",
    code: "ST",
    color: RTP_STA_BLUE,
    defaultSupplier: "STAKIS",
    widths: RTP_STA_WIDTHS
  }
];
