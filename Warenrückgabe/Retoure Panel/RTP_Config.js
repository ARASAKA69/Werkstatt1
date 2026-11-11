var RTP_VERSION = "1.0.1";
var RTP_ICON_SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><defs><linearGradient id="b" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#3FA0DB"/><stop offset="1" stop-color="#1E56B8"/></linearGradient></defs><rect width="64" height="64" rx="16" fill="url(#b)"/><path d="M32 14 49 23v16L32 48 15 39V23Z" fill="#fff" fill-opacity=".14"/><path d="M15 23 32 32l17-9M32 32v16M23.2 18.6l17 9" fill="none" stroke="#fff" stroke-width="2.6" stroke-linejoin="round" stroke-linecap="round"/><circle cx="46" cy="46" r="11" fill="#F27420"/><path d="M50 42.5a5.2 5.2 0 1 0 .9 5.7" fill="none" stroke="#fff" stroke-width="2.3" stroke-linecap="round"/><path d="M50.6 39.6 51.2 43.4 47.6 44.2" fill="none" stroke="#fff" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"/></svg>';

var RTP_SHEET_ID = "1P8s15IFPyQb3OFkjW__YZZ9Q1-Qui9b4Y5k4RJMJtpc";
var RTP_TZ = "Europe/Berlin";
var RTP_INBOX = "ersatzteile.hemau@autohero.com";
var RTP_LOCK_MS = 20000;
var RTP_MAIL_BRIDGE_URL = "";
var RTP_BRIDGE_PROP = "rtp_mail_bridge_url";
var RTP_MAIL_PRIMED_PROP = "alfah_mail_primed_v3";

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
