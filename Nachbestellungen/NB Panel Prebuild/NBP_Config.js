var NBP_VERSION = "0.1.5";
var NBP_ICON = "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA2NCA2NCI+PGRlZnM+PGxpbmVhckdyYWRpZW50IGlkPSJnIiB4MT0iMCIgeTE9IjAiIHgyPSIxIiB5Mj0iMSI+PHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjRkY4QTNEIi8+PHN0b3Agb2Zmc2V0PSIxIiBzdG9wLWNvbG9yPSIjRDk1MzFFIi8+PC9saW5lYXJHcmFkaWVudD48bGluZWFyR3JhZGllbnQgaWQ9ImIiIHgxPSIwIiB5MT0iMCIgeDI9IjEiIHkyPSIxIj48c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiM1QkI0RUEiLz48c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMyQTdBQjgiLz48L2xpbmVhckdyYWRpZW50PjwvZGVmcz48cmVjdCB3aWR0aD0iNjQiIGhlaWdodD0iNjQiIHJ4PSIxNiIgZmlsbD0idXJsKCNnKSIvPjxwYXRoIGQ9Ik0zMCAxMiA0NyAyMS41djE5TDMwIDUwIDEzIDQwLjV2LTE5WiIgZmlsbD0iI2ZmZiIgZmlsbC1vcGFjaXR5PSIuMTYiIHN0cm9rZT0iI2ZmZiIgc3Ryb2tlLXdpZHRoPSIzLjIiIHN0cm9rZS1saW5lam9pbj0icm91bmQiLz48cGF0aCBkPSJNMTMgMjEuNSAzMCAzMWwxNy05LjVNMzAgMzF2MTlNMjEuNSAxNi44bDE3IDkuNSIgZmlsbD0ibm9uZSIgc3Ryb2tlPSIjZmZmIiBzdHJva2Utd2lkdGg9IjMuMiIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIi8+PGNpcmNsZSBjeD0iNDciIGN5PSI0NyIgcj0iMTIiIGZpbGw9InVybCgjYikiIHN0cm9rZT0iI2ZmZiIgc3Ryb2tlLXdpZHRoPSIzIi8+PHBhdGggZD0iTTUyLjIgNDUuMmE1LjYgNS42IDAgMSAwLTEuMSA1LjMiIGZpbGw9Im5vbmUiIHN0cm9rZT0iI2ZmZiIgc3Ryb2tlLXdpZHRoPSIyLjgiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIvPjxwYXRoIGQ9Ik01MyA0MC4ydjUuNGgtNS40IiBmaWxsPSJub25lIiBzdHJva2U9IiNmZmYiIHN0cm9rZS13aWR0aD0iMi44IiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiLz48L3N2Zz4=";
var NBP_MAIN_ID = "1VGCAHUbOPgsInQICA1GnrtKg1EPK1d1zWB-GkLi6iVE";
var NBP_ZIEL_ID = "18mx0dJupfas4w8xuTvNdcX3TrGjx0idYfLhSBt7DHYY";
var NBP_ZIEL_TAB = "Repair Status";
var NBP_TZ = "Europe/Berlin";
var NBP_ALLOWED = ["francesco.berger@auto1.com"];
var NBP_NB_TAB = "Nachbestellung";
var NBP_DASH_TAB = "Dashboard";
var NBP_QUEUE_TAB = "EventQueue";
var NBP_EXIT_TAB = "Input Exit";
var NBP_INPUT_EID_FALLBACK = 27;
var NBP_LOCK_MS = 20000;

var NBP_SECTIONS = [
  { key: "qc", sheet: "Input Q-Check", code: "QC", label: "Q-Check" },
  { key: "me", sheet: "Input Mechanik", code: "ME", label: "Mechanik" },
  { key: "la", sheet: "Input Lack", code: "LA", label: "Lack" },
  { key: "ex", sheet: "Input Exit", code: "EX", label: "Exit" }
];

var NBP_PROTECT_EDITORS = [
  "francesco.berger@auto1.com",
  "bastian.frankhauser@auto1.com",
  "joerg.eichenseher@auto1.com",
  "tobias.nehring@auto1.com"
];

var NBP_NB_FIELDS = [
  { key: "datum", aliases: ["datum"], col: 1 },
  { key: "stockId", aliases: ["stock id"], col: 2 },
  { key: "art", aliases: ["art der nachbestellung"], col: 3 },
  { key: "besteller", aliases: ["besteller"], col: 4 },
  { key: "ersatzteil", aliases: ["ersatzteil"], col: 5 },
  { key: "preis", aliases: ["preis"], col: 6, edit: "text" },
  { key: "artikelnummer", aliases: ["artikelnummer", "artikelnr"], col: 7, edit: "text" },
  { key: "haendler", aliases: ["händler", "haendler", "handler"], col: 8, edit: "list" },
  { key: "costGate", aliases: ["durch senior genehmigt", "cost gate", "cost calc"], col: 9, edit: "list" },
  { key: "carol", aliases: ["carol eingetragen"], col: 10, edit: "list" },
  { key: "status", aliases: ["status"], col: 11, edit: "list" },
  { key: "statusAenderung", aliases: ["statusänderung", "statusanderung"], col: 12 },
  { key: "lagerort", aliases: ["lagerort"], col: 13, edit: "list" },
  { key: "bearbeitetVon", aliases: ["bearbeitet von"], col: 14, edit: "list" },
  { key: "entryId", aliases: ["entryid", "entry id", "entry_id"], col: 15 },
  { key: "kommentar", aliases: ["kommentar"], col: 16, edit: "text" }
];

var NBP_DASH_FIELDS = [
  { key: "stockId", aliases: ["stock id"] },
  { key: "letzterTagesliste", aliases: ["letzer tageslisteneintrag", "letzter tageslisteneintrag"] },
  { key: "tageSeitLetztem", aliases: ["tage seit letzten eintrag", "tage seit letztem eintrag"] },
  { key: "eintraege", aliases: ["einträge", "eintrage"] },
  { key: "herkunft", aliases: ["herkunft"] },
  { key: "tageSeitInput", aliases: ["tage seit input"] },
  { key: "bemerkung", aliases: ["bemerkung/verlauf", "bemerkung / verlauf"] },
  { key: "entryId", aliases: ["entryid", "entry id", "entry_id"] }
];
