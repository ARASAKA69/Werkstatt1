var NBP_VERSION = "0.1.1";
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
