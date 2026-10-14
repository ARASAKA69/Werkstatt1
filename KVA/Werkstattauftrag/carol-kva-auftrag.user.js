// ==UserScript==
// @name         Carol KVA Werkstattauftrag Automation
// @namespace    arasaka-kva
// @version      1.4.0
// @description  In-Carol KVA HUD: Vollautomatischer 1:1 Abgleich aller Carol-Schäden, Services, Lack/Delle (1-6) & HU-Status mit dem Sheet, erstellt PDF im Drive, lädt PDF in Carol hoch und setzt Ersatzteile auf Bestellt
// @match        *://carol.autohero.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        unsafeWindow
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @run-at       document-start
// ==/UserScript==

(function () {
  'use strict';

  var VER = '1.4.0';
  if (window.__kvaCarolHudVer) return;
  window.__kvaCarolHudVer = VER;

  var DEFAULT_WEB_APP_URL = '';
  var BRIDGE_SECRET = 'ARASAKA69';
  var MSG_MARKER = '__kvaBridgeGraphql';
  var STORE_URL_KEY = 'kva_wa_webapp_url_v1';
  var STORE_META_KEY = 'kva_wa_meta_cache_v3';
  var STORE_AUTO_PARTS_KEY = 'kva_wa_auto_parts_sid';

  var FALLBACK_META = {
    workItems: [
      { row: 11, label: 'Inspektion' },
      { row: 12, label: 'Motorölwechsel inkl. Filterersatz' },
      { row: 14, label: 'passendenes Motorenöl' },
      { row: 15, label: 'Bremsflüssigkeit wechseln' },
      { row: 16, label: 'Innenraum Filter ersetzen' },
      { row: 17, label: 'Luftfilter (Motor) ersetzen' },
      { row: 18, label: 'Luftfilter Hochvoltbatterie ersetzen' },
      { row: 19, label: 'Kraftstofffilter ersetzen' },
      { row: 20, label: 'Zündkerzenwechsel' },
      { row: 21, label: 'Schlüsselbatterie ersetzen' },
      { row: 22, label: 'Getriebeöl ersetzen' },
      { row: 23, label: 'Haldex Öl ersetzen' },
      { row: 24, label: 'Differnzialölwechsel' },
      { row: 25, label: 'Schiebe-/Panoramadach warten' },
      { row: 26, label: 'Keilriemen ersetzen' },
      { row: 27, label: 'Kühlflüssigkeitswechsel' },
      { row: 28, label: 'Zahnriemen ersetzen' }
    ],
    oils: [
      { row: 29, label: '7000 AV 0W-20', nr: '1' },
      { row: 30, label: '7000 VO 0W-20', nr: '2' },
      { row: 31, label: '7000 CP 0W-30', nr: '3' },
      { row: 32, label: '7000 FJ 0W-30', nr: '4' },
      { row: 33, label: '7000 LL 0W-30', nr: '5' },
      { row: 34, label: '5000 FR 5W-20', nr: '6' },
      { row: 35, label: '5000 AV 5W-30', nr: '7' },
      { row: 36, label: '5000 XS 5W-30', nr: '8' },
      { row: 37, label: '5000 CP 5W-30', nr: '9' },
      { row: 38, label: '5000 DM 5W-30', nr: '10' },
      { row: 39, label: '3000 FR 5W-30', nr: '11' },
      { row: 40, label: '3000 E 5W-40', nr: '12' },
      { row: 41, label: 'Öl wird mitgeliefert', nr: '' }
    ],
    dItems: [
      { row: 2, label: 'Batterie laden und prüfen' },
      { row: 3, label: 'Batterie ersetzen' },
      { row: 4, label: 'Radsatz-Wechsel' },
      { row: 5, label: 'Reifen vorne rechts erneuern' },
      { row: 6, label: 'Reifen vorne links erneuern' },
      { row: 7, label: 'Reifen hinten rechts erneuern' },
      { row: 8, label: 'Reifen hinten links erneuern' },
      { row: 9, label: 'Bremsscheiben und -beläge vorne ersetzen' },
      { row: 10, label: 'Bremsscheiben und -beläge hinten ersetzen' },
      { row: 11, label: 'Antenne ersetzen' },
      { row: 12, label: 'Ozonbehandlung' },
      { row: 13, label: 'Diagnose' },
      { row: 14, label: 'Ölverlust' },
      { row: 15, label: 'Smart Repair Windschutzscheibe' },
      { row: 16, label: 'Fehlerspeicher löschen' },
      { row: 17, label: 'Scheibenwischer links' },
      { row: 18, label: 'Scheibewischer rechts' },
      { row: 19, label: 'Heckwischer' },
      { row: 20, label: 'Marderschutz entfernen' },
      { row: 21, label: 'Windschutzscheibe ersetzen' },
      { row: 22, label: 'Moterhaubendämpfer' },
      { row: 38, label: 'Differnzialöl Bestellung' },
      { row: 39, label: 'Hinterachsdifferenzial Öl' },
      { row: 40, label: 'Differnzial Öl' },
      { row: 41, label: 'Sperrdifferenzial Öl' }
    ],
    actions: [
      { nr: 1, label: 'ganzes Bauteil lackieren' },
      { nr: 2, label: 'Beilackieren' },
      { nr: 3, label: 'Delle drücken' },
      { nr: 4, label: 'Andere (Spaltmaß einstellen)' },
      { nr: 5, label: 'Andere (Reparatur)' },
      { nr: 6, label: 'Polieren gesamtes Fahrzeug' }
    ],
    parts: [
      { side: 1, row: 2, label: 'A-Säule links' },
      { side: 1, row: 3, label: 'A-Säule rechts' },
      { side: 1, row: 4, label: 'Anhängerkupplung' },
      { side: 1, row: 5, label: 'Antenne' },
      { side: 1, row: 6, label: 'Auspuffrohr' },
      { side: 1, row: 7, label: 'Außenspiegel links' },
      { side: 1, row: 8, label: 'Außenspiegel rechts' },
      { side: 1, row: 9, label: 'Außenspiegelgehäuse links' },
      { side: 1, row: 10, label: 'Außenspiegelgehäuse rechts' },
      { side: 1, row: 11, label: 'B-Säule links' },
      { side: 1, row: 12, label: 'B-Säule rechts' },
      { side: 1, row: 13, label: 'Blinker hinten links' },
      { side: 1, row: 14, label: 'Blinker hinten rechts' },
      { side: 1, row: 15, label: 'Blinker linke Seite' },
      { side: 1, row: 16, label: 'Blinker rechts Seite' },
      { side: 1, row: 17, label: 'Blinker vorne links' },
      { side: 1, row: 18, label: 'Blinker vorne rechts' },
      { side: 1, row: 19, label: 'C-Säule links' },
      { side: 1, row: 20, label: 'C-Säule rechts' },
      { side: 1, row: 21, label: 'Cabrioverdeck' },
      { side: 1, row: 22, label: 'Cabrioverdeck Fenster' },
      { side: 1, row: 23, label: 'D-Säule links' },
      { side: 1, row: 24, label: 'D-Säule rechts' },
      { side: 1, row: 25, label: 'Dach' },
      { side: 1, row: 26, label: 'Dachreling' },
      { side: 1, row: 27, label: 'Dachträger' },
      { side: 1, row: 28, label: 'Diffusor' },
      { side: 1, row: 29, label: 'Drittes Bremslicht hinten' },
      { side: 1, row: 30, label: 'Fensterrahmen hinten' },
      { side: 1, row: 31, label: 'Frensterrahmen vorne' },
      { side: 1, row: 32, label: 'Frontspoiler' },
      { side: 1, row: 33, label: 'Hardtop' },
      { side: 1, row: 34, label: 'Heckklappe' },
      { side: 1, row: 35, label: 'Heckspoiler' },
      { side: 1, row: 36, label: 'Heckteil' },
      { side: 1, row: 37, label: 'Heckwischer' },
      { side: 1, row: 38, label: 'Kamera' },
      { side: 1, row: 39, label: 'Kotflügel hinten links' },
      { side: 1, row: 40, label: 'Kotflügel hinten rechts' },
      { side: 1, row: 41, label: 'Kotflügel vorne links' },
      { side: 1, row: 42, label: 'Kotflügel vorne rechts' },
      { side: 2, row: 2, label: 'Kühler' },
      { side: 2, row: 3, label: 'Kuhlergrill' },
      { side: 2, row: 4, label: 'Licht hinten links' },
      { side: 2, row: 5, label: 'Licht hinten rechts' },
      { side: 2, row: 6, label: 'Motorhaube' },
      { side: 2, row: 7, label: 'Nebelscheinwerfer vorne links' },
      { side: 2, row: 8, label: 'Nebelscheinwerfer vorne rechts' },
      { side: 2, row: 9, label: 'Nebelschlussleuchte hinten links' },
      { side: 2, row: 10, label: 'Nebelschlussleuchte hinten rechts' },
      { side: 2, row: 11, label: 'Nummernschild - hinten' },
      { side: 2, row: 12, label: 'Nummernschild - vorne' },
      { side: 2, row: 13, label: 'Panoramadach' },
      { side: 2, row: 14, label: 'Parksensoren hinten' },
      { side: 2, row: 15, label: 'Parksensoren vorne' },
      { side: 2, row: 16, label: 'Scheibe hinten' },
      { side: 2, row: 17, label: 'Scheibe vorne links' },
      { side: 2, row: 18, label: 'Scheibe vorne rechts' },
      { side: 2, row: 19, label: 'Scheibenwischer links' },
      { side: 2, row: 20, label: 'Scheibenwischer rechts' },
      { side: 2, row: 21, label: 'Scheinwerfer links' },
      { side: 2, row: 22, label: 'Scheinwerfer rechts' },
      { side: 2, row: 23, label: 'Schiebedach' },
      { side: 2, row: 24, label: 'Seitenscheibe hinten links' },
      { side: 2, row: 25, label: 'Seitenscheibe hinten rechts' },
      { side: 2, row: 26, label: 'Seitenschweller links' },
      { side: 2, row: 27, label: 'Seitenschweller rechts' },
      { side: 2, row: 28, label: 'Stoßsstange hinten' },
      { side: 2, row: 29, label: 'Stoßstange vorne' },
      { side: 2, row: 30, label: 'Tankdeckel links' },
      { side: 2, row: 31, label: 'Tankdeckel rechts' },
      { side: 2, row: 32, label: 'Tür hinten links' },
      { side: 2, row: 33, label: 'Tür hinten rechts' },
      { side: 2, row: 34, label: 'Tür vorne links' },
      { side: 2, row: 35, label: 'Tür vorne rechts' },
      { side: 2, row: 36, label: 'Türgriff hinten links' },
      { side: 2, row: 37, label: 'Türgriff hinten rechts' },
      { side: 2, row: 38, label: 'Türgriff vorne links' },
      { side: 2, row: 39, label: 'Türgriff vorne rechts' },
      { side: 2, row: 40, label: 'Türscheibe hinten links' },
      { side: 2, row: 41, label: 'Türscheibe hinten rechts' },
      { side: 2, row: 42, label: 'Windschutzscheibe' }
    ]
  };

  var state = {
    open: false,
    showSettings: false,
    meta: loadCachedMeta() || FALLBACK_META,
    stockId: '',
    modell: '',
    vin: '',
    erstzulassung: '',
    getriebe: '',
    km: '',
    kba: '',
    huStatus: 'unknown',
    nach: 'nein',
    nachDatum: '',
    carolItems: [],
    workSel: {},
    dSel: {},
    autoWork: {},
    autoD: {},
    oilRow: 0,
    liters: '',
    bodyWorks: [],
    freeMech: ['', '', '', '', ''],
    freeParts: ['', '', '', '', ''],
    autoUploadPdf: true,
    autoOrderParts: true,
    runUebertrag: false,
    busy: false,
    lastScannedSid: '',
    graphqlCapture: {}
  };

  function getStored(key, defVal) {
    try {
      if (typeof GM_getValue === 'function') {
        var v = GM_getValue(key, '');
        if (v) return v;
      }
    } catch (e) {}
    try {
      return localStorage.getItem(key) || defVal || '';
    } catch (e2) {
      return defVal || '';
    }
  }

  function setStored(key, val) {
    try {
      if (typeof GM_setValue === 'function') GM_setValue(key, val);
    } catch (e) {}
    try {
      localStorage.setItem(key, val);
    } catch (e2) {}
  }

  function getWebAppUrl() {
    return String(getStored(STORE_URL_KEY, DEFAULT_WEB_APP_URL) || '').trim();
  }

  function loadCachedMeta() {
    try {
      var raw = getStored(STORE_META_KEY, '');
      if (!raw) return null;
      var parsed = JSON.parse(raw);
      if (parsed && parsed.workItems && parsed.dItems && parsed.parts && parsed.parts.length >= 40) return parsed;
    } catch (e) {}
    return null;
  }

  function injectGraphqlHook() {
    var hookSrc = '(function(){'
      + 'if(window.__kvaGraphqlHooked)return;window.__kvaGraphqlHooked=true;'
      + 'var of=window.fetch;'
      + 'window.fetch=function(){var a=arguments;'
      + 'var u=String(a[0]&&a[0].url?a[0].url:a[0]||"");'
      + 'var p=of.apply(this,a);'
      + 'if(u.indexOf("graphql")!==-1||u.indexOf("api")!==-1){p.then(function(r){'
      + 'try{r.clone().json().then(function(d){'
      + 'window.postMessage({' + MSG_MARKER + ':1,url:u,data:d},"*");'
      + '}).catch(function(){});}catch(e){}return r;});}'
      + 'return p;};'
      + 'var ox=XMLHttpRequest.prototype.open;var os=XMLHttpRequest.prototype.send;'
      + 'XMLHttpRequest.prototype.open=function(m,u){this.__kvaUrl=String(u||"");return ox.apply(this,arguments);};'
      + 'XMLHttpRequest.prototype.send=function(){var x=this;'
      + 'if(String(x.__kvaUrl).indexOf("graphql")!==-1||String(x.__kvaUrl).indexOf("api")!==-1){x.addEventListener("load",function(){'
      + 'try{var d=JSON.parse(x.responseText);'
      + 'window.postMessage({' + MSG_MARKER + ':1,url:x.__kvaUrl,data:d},"*");}catch(e){}});}'
      + 'return os.apply(this,arguments);};'
      + '})();';
    try {
      var s = document.createElement('script');
      s.textContent = hookSrc;
      (document.head || document.documentElement).appendChild(s);
      s.remove();
    } catch (e) {}
    try {
      if (typeof unsafeWindow !== 'undefined' && !unsafeWindow.__kvaGraphqlHooked && unsafeWindow.eval) {
        unsafeWindow.eval(hookSrc);
      }
    } catch (e2) {}
  }

  window.addEventListener('message', function (ev) {
    var d = ev && ev.data;
    if (!d || !d[MSG_MARKER] || !d.data) return;
    try { scanGraphqlNode(d.data, 0); } catch (e) {}
  });

  function strVal(v) {
    if (typeof v === 'string') return v.trim();
    if (v && typeof v === 'object' && typeof v.name === 'string') return v.name.trim();
    return '';
  }

  function cleanModelString(raw) {
    var s = String(raw || '').replace(/\s+/g, ' ').trim();
    s = s.replace(/Refurbishment\s*n[°ºo.]?.*$/i, '').trim();
    s = s.replace(/\bVIN\s*:.*$/i, '').trim();
    s = s.replace(/\b[A-HJ-NPR-Z0-9]{17}\b.*$/i, '').trim();
    return s;
  }

  function scanGraphqlNode(node, depth) {
    if (!node || typeof node !== 'object' || depth > 14) return;
    if (Array.isArray(node)) {
      for (var i = 0; i < node.length; i++) scanGraphqlNode(node[i], depth + 1);
      return;
    }
    if (typeof node.vin === 'string' && /^[A-HJ-NPR-Z0-9]{11,17}$/i.test(node.vin.trim())) {
      var vin = node.vin.trim().toUpperCase();
      var sid = '';
      var cands = ['stockNumber', 'stockId', 'stock_number', 'adId'];
      for (var c = 0; c < cands.length; c++) {
        if (typeof node[cands[c]] === 'string' && /^[A-Z]{2}\d{4,}$/i.test(node[cands[c]].trim())) {
          sid = node[cands[c]].trim().toUpperCase();
          break;
        }
      }
      var make = strVal(node.make) || strVal(node.brand) || strVal(node.manufacturer);
      var model = strVal(node.model) || strVal(node.modelName);
      var sub = strVal(node.subModel) || strVal(node.variant) || strVal(node.trim);
      var full = cleanModelString(strVal(node.fullName) || strVal(node.title) || [make, model, sub].filter(Boolean).join(' '));
      if (sid || vin) {
        state.graphqlCapture = {
          stockId: sid || state.graphqlCapture.stockId || '',
          vin: vin,
          modell: full || state.graphqlCapture.modell || ''
        };
      }
    }
    var keys = Object.keys(node);
    for (var k = 0; k < keys.length; k++) scanGraphqlNode(node[keys[k]], depth + 1);
  }

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function normTxt(s) {
    return String(s || '').toLowerCase()
      .replace(/ä/g, 'ae').replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ß/g, 'ss')
      .replace(/stossstange/g, 'stossstange')
      .replace(/stosstange/g, 'stossstange')
      .replace(/frensterrahmen/g, 'fensterrahmen')
      .replace(/kuhlergrill/g, 'kuehlergrill')
      .replace(/differnzial/g, 'differenzial')
      .replace(/scheibewischer/g, 'scheibenwischer')
      .replace(/moterhauben/g, 'motorhauben')
      .replace(/passendenes/g, 'passendes')
      .replace(/[^a-z0-9]+/g, ' ').trim();
  }

  function collapseDoubleConsonants(s) {
    return normTxt(s).replace(/sss/g, 'ss');
  }

  function hasWord(hay, word) {
    return new RegExp('\\b' + String(word).replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b').test(hay || '');
  }

  function hasAllWords(hay, words) {
    for (var i = 0; i < words.length; i++) {
      if (!hasWord(hay, words[i])) return false;
    }
    return true;
  }

  function hasPhrase(hay, phrase) {
    return (' ' + collapseDoubleConsonants(hay) + ' ').indexOf(' ' + collapseDoubleConsonants(phrase) + ' ') !== -1;
  }

  function isRefurbDetailPage() {
    return /\/refurbishment\/[0-9a-f-]{8,}/i.test(location.pathname);
  }

  function getPageTextWithoutDrawer() {
    if (!document.body) return '';
    var dr = document.getElementById('kva-drawer');
    var fab = document.getElementById('kva-fab');
    var prevDr = dr ? dr.style.display : '';
    var prevFab = fab ? fab.style.display : '';
    if (dr) dr.style.display = 'none';
    if (fab) fab.style.display = 'none';
    var txt = String(document.body.innerText || '');
    if (dr) dr.style.display = prevDr;
    if (fab) fab.style.display = prevFab;
    return txt;
  }

  function expandCollapsedCarolServices() {
    var dr = document.getElementById('kva-drawer');
    var btns = Array.from(document.querySelectorAll('button, a, span, div'));
    for (var i = 0; i < btns.length; i++) {
      var el = btns[i];
      if (dr && dr.contains(el)) continue;
      if (el.children && el.children.length > 2) continue;
      var t = String(el.textContent || '').trim();
      if (/^(zeige\s+alle\s+services|show\s+all\s+services)$/i.test(t) && !el.__kvaExpanded) {
        el.__kvaExpanded = true;
        try { el.click(); } catch (e) {}
      }
    }
  }

  function normalizeGermanDate(raw) {
    var s = String(raw || '').trim();
    if (!s || s === '-' || s === '–') return '';
    if (/^\d{2}\.\d{2}\.\d{4}$/.test(s)) return s;
    var months = {
      jan: '01', januar: '01', january: '01',
      feb: '02', februar: '02', february: '02',
      maer: '03', mar: '03', maerz: '03', mrz: '03', march: '03',
      apr: '04', april: '04',
      mai: '05', may: '05',
      jun: '06', juni: '06', june: '06',
      jul: '07', juli: '07', july: '07',
      aug: '08', august: '08',
      sep: '09', sept: '09', september: '09',
      okt: '10', oktober: '10', oct: '10', october: '10',
      nov: '11', november: '11',
      dez: '12', dezember: '12', dec: '12', december: '12'
    };
    var m = s.match(/(\d{1,2})\.?\s*([A-Za-zÄÖÜäöü]+)\.?\s*(\d{4})/);
    if (m) {
      var day = ('0' + m[1]).slice(-2);
      var monKey = m[2].toLowerCase().replace(/ä/g, 'ae').replace(/ö/g, 'oe').replace(/ü/g, 'ue');
      var mon = months[monKey] || months[monKey.slice(0, 3)];
      if (mon) return day + '.' + mon + '.' + m[3];
    }
    var mNum = s.match(/(\d{1,2})[./-](\d{1,2})[./-](\d{4})/);
    if (mNum) {
      return ('0' + mNum[1]).slice(-2) + '.' + ('0' + mNum[2]).slice(-2) + '.' + mNum[3];
    }
    return s;
  }

  function scrapeCarolDom() {
    var dr = document.getElementById('kva-drawer');
    var bodyText = getPageTextWithoutDrawer();
    var out = {
      stockId: '',
      modell: '',
      vin: '',
      erstzulassung: '',
      getriebe: '',
      km: '',
      huStatus: 'unknown',
      huDueDate: '',
      huDoneDate: '',
      detectedLiters: '',
      items: []
    };

    var headings = document.querySelectorAll('h1, h2, h3, [class*="title"], [class*="Title"], [class*="Header"]');
    for (var i = 0; i < headings.length; i++) {
      if (dr && dr.contains(headings[i])) continue;
      var hTxt = String(headings[i].innerText || headings[i].textContent || '').split(/\n/)[0].trim();
      var m = hTxt.match(/\b([A-Z]{2}\d{4,8})\s*[-–]\s*(.+)$/);
      if (m) {
        out.stockId = m[1].toUpperCase();
        out.modell = cleanModelString(m[2]);
        break;
      }
    }
    if (!out.stockId) {
      var mSid = bodyText.match(/\b([A-Z]{2}\d{5,8})\s*[-–]\s*([^\n\r]+)/);
      if (mSid) {
        out.stockId = mSid[1].toUpperCase();
        out.modell = cleanModelString(mSid[2]);
      }
    }

    var mVin = bodyText.match(/VIN[:\s]*([A-HJ-NPR-Z0-9]{17})/i);
    if (mVin) out.vin = mVin[1].toUpperCase();

    var mKm = bodyText.match(/Laufleistung[\s\n:]*([\d.,]+)\s*KM/i) || bodyText.match(/Mileage[\s\n:]*([\d.,]+)\s*KM/i);
    if (mKm) out.km = mKm[1].replace(/[.,\s]/g, '');

    var mGetr = bodyText.match(/Getriebeart[\s\n:]*([^\n\r]+)/i) || bodyText.match(/Transmission[\s\n:]*([^\n\r]+)/i);
    if (mGetr) out.getriebe = mGetr[1].trim();

    var mEz = bodyText.match(/Datum\s+Erstzulassung[\s\n:]*([^\n\r]+)/i) || bodyText.match(/First\s+registration[\s\n:]*([^\n\r]+)/i);
    if (mEz) out.erstzulassung = normalizeGermanDate(mEz[1].trim());

    var mHuDue = bodyText.match(/Hauptuntersuchung\s*\(HU\)\s*Fällig\s*Am[\s\n:]*([^\n\r]+)/i)
      || bodyText.match(/HU\s+Fällig\s+Am[\s\n:]*([^\n\r]+)/i)
      || bodyText.match(/Main\s+inspection\s*\(HU\)\s*due\s*on[\s\n:]*([^\n\r]+)/i);
    if (mHuDue) out.huDueDate = normalizeGermanDate(mHuDue[1].trim());

    var mHuDone = bodyText.match(/Hauptuntersuchung\s*\(HU\)\s*Durchgeführt\s*Am[\s\n:]*([^\n\r]+)/i)
      || bodyText.match(/Main\s+inspection\s*\(HU\)\s*performed\s*on[\s\n:]*([^\n\r]+)/i);
    if (mHuDone) out.huDoneDate = normalizeGermanDate(mHuDone[1].trim());

    if (state.graphqlCapture.vin && (!out.vin || out.vin === state.graphqlCapture.vin)) {
      if (!out.vin) out.vin = state.graphqlCapture.vin;
      if (!out.stockId && state.graphqlCapture.stockId) out.stockId = state.graphqlCapture.stockId;
      if (!out.modell && state.graphqlCapture.modell) out.modell = cleanModelString(state.graphqlCapture.modell);
    }

    if (/HU\s+nicht\s+bestanden|HU\s+failed|Nachuntersuchung\s+erforderlich/i.test(bodyText)) {
      out.huStatus = 'failed';
    } else if (/HU\s+bestanden|HU\s+passed|HU\s+Bericht\s+Verfügbar[\s\n:]*Ja/i.test(bodyText)) {
      out.huStatus = 'passed';
    }

    var seenItems = {};
    function isDocumentOrFilename(str) {
      var s = String(str || '').trim();
      if (!s) return true;
      if (/\.(pdf|jpg|jpeg|png|webp|gif|bmp|heic|doc|docx|xls|xlsx)\b/i.test(s)) return true;
      if (/^(image|obd-bericht|hu\s*bericht|service\s*nachweis|service\s*buch|zulassungsbescheinigung|n4parts|marketing-bilder|dokument)$/i.test(s)) return true;
      if (/^(v|vr|vl|h|hr|hl|cock|cockpit|headrest|navigation|rear\s*brake|oil\s*cap|break\s*fluid|brake\s*fluid|u)$/i.test(s)) return true;
      return false;
    }

    function extractOilLitersFromText(txt) {
      var s = String(txt || '');
      if (!/motoröl|motoroel|ölwechsel|oelwechsel|viskosität|sae\s+\d+w/i.test(s)) return '';
      var m = s.match(/\b(\d{1,2}[.,]\d{1,2}|\d{1,2})\s*(?:liter|l\b)/i);
      if (!m) return '';
      var val = m[1].replace('.', ',').replace(/,(\d)0$/, ',$1').replace(/,00$/, '');
      return val;
    }

    function addStructuredItem(entry) {
      var comp = String(entry.component || '').replace(/\s+/g, ' ').trim();
      var srv = String(entry.service || '').replace(/\s+/g, ' ').trim();
      var dmg = String(entry.damage || '').replace(/\s+/g, ' ').trim();
      var comm = String(entry.comment || '').replace(/\s+/g, ' ').trim();
      if (comm === '-' || comm === '–' || /^keine\s+kommentare\s+verfügbar\.?$/i.test(comm) || /^no\s+comments\s+available\.?$/i.test(comm)) {
        comm = '';
      }
      if (dmg === '-' || dmg === '–') dmg = '';

      var litVal = extractOilLitersFromText(comm);
      if (litVal && !out.detectedLiters) {
        out.detectedLiters = litVal;
      }

      if (isDocumentOrFilename(comp) && isDocumentOrFilename(srv)) return;
      if (/\.(pdf|jpg|jpeg|png)\b/i.test(comp) || /\.(pdf|jpg|jpeg|png)\b/i.test(srv)) return;

      if (/^(außenreinigung|aussenreinigung|innenreinigung|fotos|foto|fahrzeugreinigung|aufbereitung|marketing\s*bilder)/i.test(srv)) {
        return;
      }

      var includedInOther = false;
      if (/^in\s+anderer\s+dienstleistung\s+enthalten|^included\s+in\s+another\s+service/i.test(srv)) {
        srv = '';
        includedInOther = true;
        if (!dmg && !comm) return;
      }

      if (!comp && !srv && !comm) return;
      if (/^(leistung|service|dienstleistung|fahrzeugkomponente|kommentar|status|kosten|ersatzteilkosten|letztes\s+update\s+am)$/i.test(comp || srv)) return;

      var displayTitle = '';
      if (comp && srv && !/^(andere|other)$/i.test(srv)) {
        displayTitle = srv + ' > ' + comp;
      } else if (comp) {
        displayTitle = comp + (dmg && !/^(keine\s+leistung)$/i.test(dmg) ? (' (' + dmg + ')') : '');
      } else if (srv && !/^(andere|other)$/i.test(srv)) {
        displayTitle = srv;
      } else {
        displayTitle = comm;
      }

      if (!displayTitle || isDocumentOrFilename(displayTitle)) return;

      var key = collapseDoubleConsonants([comp, srv, dmg, comm].join('|'));
      if (!key || seenItems[key]) return;
      seenItems[key] = 1;

      out.items.push({
        component: comp,
        service: srv,
        damage: dmg,
        comment: comm,
        title: displayTitle,
        includedInOther: includedInOther,
        source: entry.source || 'Carol'
      });
    }

    var tables = document.querySelectorAll('table');
    for (var tb = 0; tb < tables.length; tb++) {
      if (dr && dr.contains(tables[tb])) continue;
      var tblText = String(tables[tb].innerText || '');
      if (/image\.jpg|obd-bericht|hu\s+bericht|dokument\s+hochladen/i.test(tblText) && !/ersatzteilkosten/i.test(tblText)) {
        continue;
      }
      var headers = Array.from(tables[tb].querySelectorAll('th')).map(function (th) {
        return String(th.innerText || '').trim().toLowerCase();
      });
      var isPartsTable = headers.some(function (h) {
        return h.indexOf('leistung') !== -1 || h.indexOf('ersatzteil') !== -1 || h.indexOf('service') !== -1;
      });
      if (!isPartsTable) continue;

      var rows = tables[tb].querySelectorAll('tbody tr');
      for (var r = 0; r < rows.length; r++) {
        var cells = rows[r].querySelectorAll('td');
        if (cells.length >= 2) {
          var c0 = String(cells[0].innerText || '').trim();
          var c1 = String(cells[1].innerText || '').trim();
          var c2 = cells[2] ? String(cells[2].innerText || '').trim() : '';
          var srvCell = (!c0 && c1) ? c1 : c0;
          var commCell = (!c0 && c1) ? c2 : c1;
          if (srvCell && !isDocumentOrFilename(srvCell)) {
            addStructuredItem({ service: srvCell, comment: commCell, source: 'Ersatzteile' });
          }
        }
      }
    }

    var tokenLines = bodyText
      .split(/\r?\n|\t+/)
      .map(function (l) { return l.replace(/\s+/g, ' ').trim(); })
      .filter(Boolean);

    var labelMap = {
      'schadenstyp': 'damage',
      'damage type': 'damage',
      'fahrzeugkomponente': 'component',
      'vehicle component': 'component',
      'dienstleistung': 'service',
      'service': 'service',
      'teildienstleistung': 'subservice',
      'teildienstleistungen': 'subservice',
      'sub-service': 'subservice',
      'sub-services': 'subservice',
      'kosten': 'cost',
      'cost': 'cost',
      'reparaturstandort und -partner': 'partner',
      'repair location and partner': 'partner',
      'kommentar': 'comment',
      'comment': 'comment'
    };

    function getLabelType(str) {
      var low = String(str || '').toLowerCase().replace(/\s+/g, ' ').trim();
      return labelMap[low] || '';
    }

    var currentCard = null;
    for (var idx = 0; idx < tokenLines.length; idx++) {
      var tok = tokenLines[idx];
      var lType = getLabelType(tok);

      if (/^beauftragt$/i.test(tok)) {
        if (currentCard && (currentCard.component || currentCard.service || currentCard.comment)) {
          addStructuredItem(currentCard);
        }
        currentCard = { damage: '', component: '', service: '', comment: '', partner: '', source: 'Carol' };
        continue;
      }

      if (lType) {
        if (!currentCard) {
          currentCard = { damage: '', component: '', service: '', comment: '', partner: '', source: 'Carol' };
        }
        if (lType === 'damage' && (currentCard.component || currentCard.service)) {
          addStructuredItem(currentCard);
          currentCard = { damage: '', component: '', service: '', comment: '', partner: '', source: 'Carol' };
        }
        if (lType === 'service' && !currentCard.damage && !currentCard.component && currentCard.service) {
          addStructuredItem(currentCard);
          currentCard = { damage: '', component: '', service: '', comment: '', partner: '', source: 'Carol' };
        }
        if (lType === 'subservice') {
          while (idx + 1 < tokenLines.length) {
            var subTok = tokenLines[idx + 1];
            if (getLabelType(subTok) || /^(beauftragt|vom\s+system\s+erstellt|total\s+supporting|räder|dokumente)$/i.test(subTok) || /^\d+\s+service\s*\|/i.test(subTok)) {
              break;
            }
            addStructuredItem({ service: subTok, source: 'Wartung' });
            currentCard.hasSubServices = true;
            idx++;
          }
          continue;
        }
        var val = (idx + 1 < tokenLines.length) ? tokenLines[idx + 1] : '';
        if (val && !getLabelType(val) && !/^beauftragt$/i.test(val) && !/^\d+\s+service\s*\|/i.test(val)) {
          if (lType === 'damage') currentCard.damage = val;
          else if (lType === 'component') currentCard.component = val;
          else if (lType === 'service') currentCard.service = val;
          else if (lType === 'partner') {
            currentCard.partner = val;
            if (/lack/i.test(val)) currentCard.source = 'Lack';
            else if (/mechanik/i.test(val)) currentCard.source = 'Mechanik';
          } else if (lType === 'comment') {
            currentCard.comment = val;
            addStructuredItem(currentCard);
            currentCard = null;
          }
          idx++;
        }
      }
    }
    if (currentCard && (currentCard.component || currentCard.service || currentCard.comment)) {
      addStructuredItem(currentCard);
    }

    return out;
  }

  function findItemByTokens(items, tokens) {
    if (!items || !tokens || !tokens.length) return null;
    for (var i = 0; i < items.length; i++) {
      var n = collapseDoubleConsonants(items[i].label);
      var ok = true;
      for (var t = 0; t < tokens.length; t++) {
        var tok = collapseDoubleConsonants(tokens[t]);
        if (!hasWord(n, tok) && n.indexOf(tok) === -1) { ok = false; break; }
      }
      if (ok) return items[i];
    }
    return null;
  }

  function findExactSheetPart(componentText, allText) {
    var parts = (state.meta && state.meta.parts && state.meta.parts.length) ? state.meta.parts : FALLBACK_META.parts;
    var compNorm = collapseDoubleConsonants(componentText);
    var allNorm = collapseDoubleConsonants(allText);

    if (compNorm) {
      for (var i = 0; i < parts.length; i++) {
        if (collapseDoubleConsonants(parts[i].label) === compNorm) return parts[i];
      }
      var bestPart = null;
      var bestLen = 0;
      for (var j = 0; j < parts.length; j++) {
        var pNorm = collapseDoubleConsonants(parts[j].label);
        if (pNorm.length >= 4 && (hasPhrase(compNorm, pNorm) || hasPhrase(pNorm, compNorm))) {
          if (pNorm.length > bestLen) {
            bestLen = pNorm.length;
            bestPart = parts[j];
          }
        }
      }
      if (bestPart) return bestPart;
    }

    var bestFromAll = null;
    var bestAllLen = 0;
    for (var k = 0; k < parts.length; k++) {
      var pkNorm = collapseDoubleConsonants(parts[k].label);
      if (pkNorm.length >= 5 && hasPhrase(allNorm, pkNorm)) {
        if (pkNorm.length > bestAllLen) {
          bestAllLen = pkNorm.length;
          bestFromAll = parts[k];
        }
      }
    }
    return bestFromAll;
  }

  function detectBodyActionNrs(serviceText, damageText, commentText) {
    var srv = collapseDoubleConsonants(serviceText);
    var combined = collapseDoubleConsonants([serviceText, damageText, commentText].join(' '));
    var acts = [];

    if (hasPhrase(srv, 'ganzes bauteil lackieren') || hasPhrase(combined, 'ganzes bauteil lackieren')) {
      acts.push(1);
    } else if (hasWord(srv, 'beilackieren') || hasWord(combined, 'beilackieren') || hasPhrase(combined, 'smart repair')) {
      acts.push(2);
    } else if (hasWord(srv, 'lackieren') || hasWord(srv, 'lackierung')) {
      acts.push(1);
    }

    if (hasPhrase(srv, 'delle druecken') || hasWord(srv, 'delle') || hasPhrase(combined, 'delle druecken')) {
      if (acts.indexOf(3) === -1) acts.push(3);
    }
    if (hasWord(srv, 'spaltmass') || hasWord(combined, 'spaltmass')) {
      if (acts.indexOf(4) === -1) acts.push(4);
    }
    if (hasPhrase(srv, 'polieren gesamtes fahrzeug') || hasWord(srv, 'polieren') || hasWord(srv, 'polierung')) {
      if (acts.indexOf(6) === -1) acts.push(6);
    }
    if (!acts.length && (hasWord(srv, 'reparieren') || hasWord(srv, 'reparatur') || hasWord(srv, 'instandsetzen'))) {
      acts.push(5);
    }
    return acts;
  }

  function addOrMergeAutoBodyWork(part, actionNrs) {
    if (!part || !actionNrs || !actionNrs.length) return;
    var metaActions = (state.meta && state.meta.actions) || FALLBACK_META.actions;
    var existingIdx = -1;
    for (var i = 0; i < state.bodyWorks.length; i++) {
      if (state.bodyWorks[i].side === part.side && state.bodyWorks[i].row === part.row) {
        existingIdx = i;
        break;
      }
    }
    var mergedActs = existingIdx >= 0 ? state.bodyWorks[existingIdx].actions.slice() : [];
    actionNrs.forEach(function (nr) {
      if (mergedActs.indexOf(nr) === -1) mergedActs.push(nr);
    });
    mergedActs.sort();
    var actionLabels = [];
    metaActions.forEach(function (a) {
      if (mergedActs.indexOf(a.nr) !== -1) actionLabels.push(a.label);
    });
    var item = {
      side: part.side,
      row: part.row,
      label: part.label,
      actions: mergedActs,
      actionLabels: actionLabels,
      auto: true
    };
    if (existingIdx >= 0) state.bodyWorks[existingIdx] = item;
    else state.bodyWorks.push(item);
  }

  function expandText(s) {
    var n = ' ' + collapseDoubleConsonants(s) + ' ';
    n = n.replace(/\bwss\b/g, ' windschutzscheibe ');
    n = n.replace(/\bern\b/g, ' ersetzen ');
    n = n.replace(/\bva\b/g, ' vorne ');
    n = n.replace(/\bha\b/g, ' hinten ');
    n = n.replace(/\bvl\b/g, ' vorne links ');
    n = n.replace(/\bvr\b/g, ' vorne rechts ');
    n = n.replace(/\bhl\b/g, ' hinten links ');
    n = n.replace(/\bhr\b/g, ' hinten rechts ');
    n = n.replace(/\breplace\s+front\s+right\s+tire\b/g, ' reifen vorne rechts erneuern ');
    n = n.replace(/\breplace\s+front\s+left\s+tire\b/g, ' reifen vorne links erneuern ');
    n = n.replace(/\breplace\s+rear\s+right\s+tire\b/g, ' reifen hinten rechts erneuern ');
    n = n.replace(/\breplace\s+rear\s+left\s+tire\b/g, ' reifen hinten links erneuern ');
    n = n.replace(/\breplace\s+rear\s+brake\s+axle\s+disks\s+and\s+pads\b/g, ' bremsscheiben und belaege hinten ersetzen ');
    n = n.replace(/\breplace\s+front\s+brake\s+axle\s+disks\s+and\s+pads\b/g, ' bremsscheiben und belaege vorne ersetzen ');
    n = n.replace(/\bengine\s+oil\s+including\s+filter\s+change\b/g, ' motoroelwechsel inkl filterersatz ');
    n = n.replace(/\binterior\s+filter\s+change\b/g, ' innenraum filter ersetzen ');
    return n.replace(/\s+/g, ' ').trim();
  }

  function autoMatchItems(items) {
    state.workSel = {};
    state.dSel = {};
    state.autoWork = {};
    state.autoD = {};
    state.bodyWorks = [];
    state.freeMech = ['', '', '', '', ''];

    var meta = state.meta || FALLBACK_META;
    var unmatched = [];

    var rules = [
      { kind: 'd', tokens: ['batterie', 'laden'], all: ['batterie', 'laden'] },
      { kind: 'd', tokens: ['batterie', 'ersetzen'], all: ['batterie', 'ersetzen'], blockIf: ['laden'] },
      { kind: 'd', tokens: ['bremsscheiben', 'vorne'], all: ['bremsscheiben', 'vorne'] },
      { kind: 'd', tokens: ['bremsscheiben', 'hinten'], all: ['bremsscheiben', 'hinten'] },
      { kind: 'd', tokens: ['reifen', 'vorne', 'rechts'], all: ['reifen', 'vorne', 'rechts'] },
      { kind: 'd', tokens: ['reifen', 'vorne', 'links'], all: ['reifen', 'vorne', 'links'] },
      { kind: 'd', tokens: ['reifen', 'hinten', 'rechts'], all: ['reifen', 'hinten', 'rechts'] },
      { kind: 'd', tokens: ['reifen', 'hinten', 'links'], all: ['reifen', 'hinten', 'links'] },
      { kind: 'd', tokens: ['radsatz'], words: ['radsatz', 'raederwechsel'] },
      { kind: 'd', tokens: ['scheibenwischer', 'links'], all: ['scheibenwischer', 'links'] },
      { kind: 'd', tokens: ['scheibenwischer', 'rechts'], all: ['scheibenwischer', 'rechts'] },
      { kind: 'd', tokens: ['heckwischer'], words: ['heckwischer'] },
      { kind: 'd', tokens: ['windschutzscheibe', 'smart'], all: ['smart', 'windschutzscheibe'] },
      { kind: 'd', tokens: ['windschutzscheibe', 'ersetzen'], all: ['windschutzscheibe', 'ersetzen'], blockIf: ['smart'] },
      { kind: 'd', tokens: ['antenne', 'ersetzen'], all: ['antenne', 'ersetzen'] },
      { kind: 'd', tokens: ['ozon'], words: ['ozon', 'ozonbehandlung'] },
      { kind: 'd', tokens: ['diagnose'], all: ['diagnose', 'vasold'] },
      { kind: 'd', tokens: ['fehlerspeicher'], words: ['fehlerspeicher'] },
      { kind: 'd', tokens: ['oelverlust'], words: ['oelverlust'] },
      { kind: 'd', tokens: ['marderschutz'], words: ['marderschutz'] },
      { kind: 'd', tokens: ['motorhaubendaempfer'], words: ['motorhaubendaempfer', 'moterhaubendaempfer'] },
      { kind: 'work', tokens: ['inspektion'], words: ['inspektion'] },
      { kind: 'work', tokens: ['motoroelwechsel'], words: ['motoroelwechsel', 'oelwechsel', 'motorolwechsel'] },
      { kind: 'work', tokens: ['bremsfluessigkeit'], words: ['bremsfluessigkeit'] },
      { kind: 'work', tokens: ['innenraum'], words: ['innenraumfilter', 'innenraum filter', 'pollenfilter'] },
      { kind: 'work', tokens: ['luftfilter', 'motor'], words: ['luftfilter'], blockIf: ['hochvolt'] },
      { kind: 'work', tokens: ['luftfilter', 'hochvolt'], all: ['luftfilter', 'hochvolt'] },
      { kind: 'work', tokens: ['kraftstofffilter'], words: ['kraftstofffilter'] },
      { kind: 'work', tokens: ['zuendkerzen'], words: ['zuendkerzen', 'zuendkerzenwechsel'] },
      { kind: 'work', tokens: ['kuehlfluessigkeit'], words: ['kuehlfluessigkeit', 'kuehlfluessigkeitswechsel', 'kuehlmittel'] },
      { kind: 'work', tokens: ['getriebeoel'], words: ['getriebeoel'] },
      { kind: 'work', tokens: ['haldex'], words: ['haldex'] },
      { kind: 'work', tokens: ['differenzialoelwechsel'], words: ['differenzialoelwechsel', 'differnzialoelwechsel'] },
      { kind: 'work', tokens: ['keilriemen'], words: ['keilriemen'] },
      { kind: 'work', tokens: ['zahnriemen'], words: ['zahnriemen'] },
      { kind: 'work', tokens: ['schluesselbatterie'], words: ['schluesselbatterie'] },
      { kind: 'work', tokens: ['panoramadach'], words: ['schiebedach', 'panoramadach'] }
    ];

    items.forEach(function (entry) {
      var isGenericWrapper = /^(andere|other|diagnose)$/i.test(String(entry.service || '').trim());
      if (/^wartung$/i.test(String(entry.service || '').trim()) && /motoröl|motoroel|viskosität/i.test(String(entry.comment || ''))) {
        return;
      }

      var coreText = isGenericWrapper
        ? [entry.component, entry.damage].filter(Boolean).join(' ')
        : [entry.component, entry.service, entry.damage, entry.title].filter(Boolean).join(' ');
      var commText = String(entry.comment || '');
      var expCore = expandText(coreText);
      var expComm = expandText(commText);
      var expAll = expandText((coreText + ' ' + commText).trim());
      var hit = false;

      var bodyActs = detectBodyActionNrs(entry.service || entry.title, entry.damage, '');
      if (bodyActs.length) {
        var matchedPart = findExactSheetPart(entry.component, coreText);
        if (matchedPart) {
          addOrMergeAutoBodyWork(matchedPart, bodyActs);
          hit = true;
        }
      }

      if (!hit && entry.component && /^(ersetzen|reparieren|instandsetzen)$/i.test(String(entry.service || '').trim())) {
        var matchedRepPart = findExactSheetPart(entry.component, '');
        if (matchedRepPart) {
          addOrMergeAutoBodyWork(matchedRepPart, [5]);
          hit = true;
        }
      }

      if (!hit && collapseDoubleConsonants(entry.component) === 'batterie') {
        if (hasWord(expAll, 'laden')) {
          state.dSel[2] = true;
          state.autoD[2] = true;
          hit = true;
        } else if (hasWord(expAll, 'ersetzen') || hasWord(expAll, 'erneuern')) {
          state.dSel[3] = true;
          state.autoD[3] = true;
          hit = true;
        }
      }

      if (!hit) {
        meta.workItems.forEach(function (wIt) {
          if (wIt.row === 10) return;
          if (hasPhrase(expCore, wIt.label) || (!entry.component && hasPhrase(expComm, wIt.label))) {
            state.workSel[wIt.row] = true;
            state.autoWork[wIt.row] = true;
            hit = true;
          }
        });

        meta.dItems.forEach(function (dIt) {
          if (dIt.row === 13 && !hasWord(expAll, 'vasold')) return;
          if (hasPhrase(expCore, dIt.label) || (!entry.component && hasPhrase(expComm, dIt.label))) {
            if (dIt.row === 3 && hasWord(expAll, 'laden')) return;
            state.dSel[dIt.row] = true;
            state.autoD[dIt.row] = true;
            hit = true;
          }
        });
      }

      if (!hit) {
        var checkTarget = entry.component ? expCore : expAll;
        rules.forEach(function (rule) {
          if (rule.blockIf) {
            for (var b = 0; b < rule.blockIf.length; b++) {
              if (hasWord(expAll, rule.blockIf[b])) return;
            }
          }
          var rHit = false;
          if (rule.all && hasAllWords(checkTarget, rule.all)) rHit = true;
          if (!rHit && rule.words) {
            for (var w = 0; w < rule.words.length; w++) {
              if (hasPhrase(checkTarget, rule.words[w]) || hasWord(checkTarget, rule.words[w])) { rHit = true; break; }
            }
          }
          if (!rHit) return;
          var list = rule.kind === 'work' ? meta.workItems : meta.dItems;
          var found = findItemByTokens(list, rule.tokens);
          if (found) {
            if (rule.kind === 'work') {
              state.workSel[found.row] = true;
              state.autoWork[found.row] = true;
            } else {
              state.dSel[found.row] = true;
              state.autoD[found.row] = true;
            }
            hit = true;
          }
        });
      }

      if (!hit && !entry.includedInOther) {
        var freeLine = '';
        if (entry.component && entry.service && !isGenericWrapper) {
          freeLine = entry.component + ' ' + entry.service.toLowerCase();
        } else if (!entry.component && isGenericWrapper && entry.comment) {
          freeLine = entry.comment;
        } else if (entry.component) {
          freeLine = entry.component + (entry.damage ? (' (' + entry.damage + ')') : '');
        } else if (entry.service && !isGenericWrapper) {
          freeLine = entry.service;
        }
        freeLine = String(freeLine || '').replace(/\s*\/\/\s*EQA:.*$/i, '').trim();
        if (freeLine && unmatched.indexOf(freeLine) === -1) {
          unmatched.push(freeLine);
        }
      }
    });

    if (state.workSel[12]) {
      state.workSel[14] = true;
      state.autoWork[14] = true;
    }

    for (var f = 0; f < Math.min(5, unmatched.length); f++) {
      state.freeMech[f] = unmatched[f];
    }
  }

  function scanAndPopulateFromCarol(force) {
    expandCollapsedCarolServices();
    var scraped = scrapeCarolDom();
    if (!scraped.stockId && !scraped.vin) return false;

    var isNewCar = force || (scraped.stockId && scraped.stockId !== state.lastScannedSid);
    state.stockId = scraped.stockId || state.stockId;
    state.modell = scraped.modell || state.modell;
    state.vin = scraped.vin || state.vin;
    state.erstzulassung = scraped.erstzulassung || state.erstzulassung;
    state.getriebe = scraped.getriebe || state.getriebe;
    state.km = scraped.km || state.km;
    state.huStatus = scraped.huStatus || 'unknown';
    state.huDueDate = scraped.huDueDate || state.huDueDate || '';
    state.huDoneDate = scraped.huDoneDate || state.huDoneDate || '';
    state.carolItems = scraped.items || [];

    if (scraped.detectedLiters && (isNewCar || !state.liters)) {
      state.liters = scraped.detectedLiters;
    }

    if (isNewCar) {
      state.lastScannedSid = state.stockId;
      state.nach = state.huStatus === 'failed' ? 'ja' : 'nein';
      state.nachDatum = state.huDueDate || '';
      autoMatchItems(state.carolItems);
    } else if (state.huDueDate && !state.nachDatum) {
      state.nachDatum = state.huDueDate;
    }
    renderHud();
    return true;
  }

  function callBackend(payload, cb) {
    var url = getWebAppUrl();
    if (!url) {
      cb({ success: false, message: 'Bitte zuerst die Apps Script Web-App URL in den Einstellungen (⚙) hinterlegen.' });
      return;
    }
    payload.secret = BRIDGE_SECRET;
    GM_xmlhttpRequest({
      method: 'POST',
      url: url,
      data: JSON.stringify(payload),
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      onload: function (res) {
        try {
          var parsed = JSON.parse(res.responseText);
          cb(parsed);
        } catch (e) {
          cb({ success: false, message: 'Ungültige Server-Antwort (HTTP ' + res.status + ')' });
        }
      },
      onerror: function () {
        cb({ success: false, message: 'Netzwerkfehler beim Aufruf des Apps Script Backends' });
      }
    });
  }

  function fetchBackendMeta() {
    var url = getWebAppUrl();
    if (!url) return;
    callBackend({ action: 'meta' }, function (res) {
      if (res && res.success && res.workItems) {
        state.meta = res;
        setStored(STORE_META_KEY, JSON.stringify(res));
        if (state.carolItems && state.carolItems.length) {
          autoMatchItems(state.carolItems);
        }
        renderHud();
      }
    });
  }

  function b64ToUint8Array(b64) {
    var bin = atob(b64);
    var len = bin.length;
    var bytes = new Uint8Array(len);
    for (var i = 0; i < len; i++) {
      bytes[i] = bin.charCodeAt(i);
    }
    return bytes;
  }

  function setNativeValue(el, value) {
    if (!el) return;
    var proto = Object.getPrototypeOf(el);
    var desc = Object.getOwnPropertyDescriptor(proto, 'value');
    if (desc && desc.set) {
      desc.set.call(el, value);
    } else {
      el.value = value;
    }
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function findButtonByText(regex) {
    var dr = document.getElementById('kva-drawer');
    var els = Array.from(document.querySelectorAll('button, a, [role="button"]'));
    for (var i = 0; i < els.length; i++) {
      if (dr && dr.contains(els[i])) continue;
      var t = String(els[i].textContent || '').replace(/\s+/g, ' ').trim();
      if (regex.test(t)) return els[i];
    }
    return null;
  }

  function waitForElement(finderFn, timeoutMs) {
    return new Promise(function (resolve) {
      var start = Date.now();
      var timer = setInterval(function () {
        var el = finderFn();
        if (el) {
          clearInterval(timer);
          resolve(el);
          return;
        }
        if (Date.now() - start >= timeoutMs) {
          clearInterval(timer);
          resolve(null);
        }
      }, 220);
    });
  }

  function sleep(ms) {
    return new Promise(function (r) { setTimeout(r, ms); });
  }

  async function selectDropdownOption(modalRoot, targetRegex) {
    var selects = modalRoot.querySelectorAll('select');
    for (var s = 0; s < selects.length; s++) {
      var opts = selects[s].options;
      for (var o = 0; o < opts.length; o++) {
        if (targetRegex.test(String(opts[o].text || '').trim())) {
          setNativeValue(selects[s], opts[o].value);
          return true;
        }
      }
    }

    var triggers = modalRoot.querySelectorAll('[role="combobox"], [aria-haspopup="listbox"], [class*="select"], [class*="Select"], [class*="dropdown"], [class*="Dropdown"]');
    for (var t = 0; t < triggers.length; t++) {
      triggers[t].click();
      await sleep(250);
      var allOptions = document.querySelectorAll('[role="option"], li, [class*="option"], [class*="Option"]');
      for (var k = 0; k < allOptions.length; k++) {
        var txt = String(allOptions[k].textContent || '').trim();
        if (targetRegex.test(txt)) {
          allOptions[k].click();
          await sleep(180);
          return true;
        }
      }
    }
    return false;
  }

  async function autoUploadPdfInCarol(pdfB64, fileName) {
    setHudStatus('wait', 'Öffne Carol Dokumenten-Upload für ' + fileName + '…');
    var openUploadBtn = findButtonByText(/dokument\s+hochladen|upload\s+document/i);
    if (!openUploadBtn) {
      throw new Error('Button "Dokument hochladen" auf der Carol-Seite nicht gefunden.');
    }
    openUploadBtn.scrollIntoView({ block: 'center' });
    openUploadBtn.click();

    var fileInput = await waitForElement(function () {
      return document.querySelector('input[type="file"]');
    }, 6000);

    if (!fileInput) {
      throw new Error('Upload-Modal (Datei-Input) hat sich nicht geöffnet.');
    }

    var modal = fileInput.closest('[role="dialog"], [class*="modal"], [class*="Modal"], form') || document.body;
    await sleep(300);

    await selectDropdownOption(modal, /^(andere|other)$/i);

    var textInputs = modal.querySelectorAll('input[type="text"], input:not([type]), textarea');
    for (var i = 0; i < textInputs.length; i++) {
      var inp = textInputs[i];
      if (inp.type === 'file' || inp.type === 'hidden') continue;
      setNativeValue(inp, 'Werkstattauftrag');
      break;
    }

    var pdfBytes = b64ToUint8Array(pdfB64);
    var fileObj = new File([pdfBytes], fileName, { type: 'application/pdf', lastModified: Date.now() });
    var dt = new DataTransfer();
    dt.items.add(fileObj);
    fileInput.files = dt.files;
    fileInput.dispatchEvent(new Event('input', { bubbles: true }));
    fileInput.dispatchEvent(new Event('change', { bubbles: true }));

    var dropZone = modal.querySelector('[class*="drop"], [class*="Drop"], [class*="upload"], [class*="Upload"]');
    if (dropZone) {
      try {
        var dropEv = new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: dt });
        dropZone.dispatchEvent(dropEv);
      } catch (eDrop) {}
    }

    await sleep(650);

    var submitBtn = null;
    var modalBtns = modal.querySelectorAll('button, [role="button"]');
    for (var b = 0; b < modalBtns.length; b++) {
      var bTxt = String(modalBtns[b].textContent || '').trim();
      if (/^(hochladen|upload)$/i.test(bTxt)) {
        submitBtn = modalBtns[b];
        break;
      }
    }
    if (!submitBtn) {
      submitBtn = findButtonByText(/^(hochladen|upload)$/i);
    }
    if (!submitBtn) {
      throw new Error('Button "Hochladen" im Modal nicht gefunden.');
    }
    submitBtn.click();
    await sleep(1200);
    return true;
  }

  async function autoSetPartsOrderedInCarol(stockId) {
    setHudStatus('wait', 'Setze Ersatzteile in Carol auf "Bestellt"…');
    var editPartsBtn = findButtonByText(/ersatzteile\s+bearbeiten|bearbeite\s+schäden\s+und\s+dienstleistungen|edit\s+damages\s+and\s+services/i);
    if (editPartsBtn && !/nachautorisierung/i.test(getPageTextWithoutDrawer())) {
      try { sessionStorage.setItem(STORE_AUTO_PARTS_KEY, stockId || '1'); } catch (e) {}
      editPartsBtn.click();
      await sleep(1400);
    }

    var sidebarParts = await waitForElement(function () {
      var els = document.querySelectorAll('a, button, li, div, span');
      for (var i = 0; i < els.length; i++) {
        var t = String(els[i].textContent || '').trim();
        if (/^Ersatzteile\s*\(\d+\)$/i.test(t) || /^Spare\s+parts\s*\(\d+\)$/i.test(t)) return els[i];
      }
      return null;
    }, 6000);

    if (sidebarParts) {
      sidebarParts.click();
      await sleep(900);
    }

    var changedCount = 0;
    var selects = document.querySelectorAll('select');
    for (var s = 0; s < selects.length; s++) {
      var opts = selects[s].options;
      for (var o = 0; o < opts.length; o++) {
        if (/^(bestellt|ordered)$/i.test(String(opts[o].text || '').trim())) {
          if (selects[s].value !== opts[o].value) {
            setNativeValue(selects[s], opts[o].value);
            changedCount++;
            await sleep(500);
          }
          break;
        }
      }
    }

    var statusCells = Array.from(document.querySelectorAll('div, span, button')).filter(function (el) {
      var t = String(el.textContent || '').trim();
      return /^(nicht\s+bestellt|not\s+ordered)$/i.test(t) && el.children.length <= 1;
    });
    for (var c = 0; c < statusCells.length; c++) {
      statusCells[c].click();
      await sleep(280);
      var optOrdered = Array.from(document.querySelectorAll('[role="option"], li, div, span')).find(function (el) {
        return /^(bestellt|ordered)$/i.test(String(el.textContent || '').trim()) && el.children.length === 0;
      });
      if (optOrdered) {
        optOrdered.click();
        changedCount++;
        await sleep(600);
      }
    }

    var closeSavedBtn = await waitForElement(function () {
      return findButtonByText(/^(schließen|close)$/i);
    }, 2500);
    if (closeSavedBtn) {
      closeSavedBtn.click();
      await sleep(500);
    }

    var finishBtn = findButtonByText(/^(abschließen|complete|finish)$/i);
    if (finishBtn && !finishBtn.disabled) {
      finishBtn.click();
      await sleep(800);
    }

    try { sessionStorage.removeItem(STORE_AUTO_PARTS_KEY); } catch (e2) {}
    return changedCount;
  }

  function isOilMitgeliefert(oilRow) {
    return Number(oilRow) === 41;
  }

  function isOilLitersRequired() {
    var oilSelected = Number(state.oilRow) >= 29 && Number(state.oilRow) <= 40;
    var oilChangeChecked = !!(state.workSel[12] || state.workSel[14]);
    return oilSelected || (oilChangeChecked && !isOilMitgeliefert(state.oilRow));
  }

  function hasValidLiters() {
    var n = parseFloat(String(state.liters || '').replace(',', '.').trim());
    return !isNaN(n) && n > 0;
  }

  async function runFullAutomation() {
    if (state.busy) return;
    syncInputsFromDom();

    if (!state.stockId || !state.modell || !state.vin) {
      setHudStatus('err', 'Stock ID, Modell und VIN sind Pflichtfelder.');
      return;
    }

    var oilChangeChecked = !!(state.workSel[12] || state.workSel[14]);
    if (oilChangeChecked && !(Number(state.oilRow) >= 29 && Number(state.oilRow) <= 41)) {
      var oilEl = document.getElementById('kva-in-oil');
      if (oilEl) {
        oilEl.style.borderColor = '#f85149';
        oilEl.focus();
      }
      setHudStatus('err', '⚠️ Motorölwechsel ist ausgewählt – bitte zuerst ein Öl auswählen!');
      return;
    }

    if (isOilLitersRequired() && !hasValidLiters()) {
      var litEl = document.getElementById('kva-in-lit');
      if (litEl) {
        litEl.style.borderColor = '#f85149';
        litEl.focus();
      }
      setHudStatus('err', '⚠️ Bitte Füllmenge (Liter) eintragen! Ohne Füllmenge erscheint #N/A auf dem Werkstattauftrag.');
      return;
    }

    if (state.nach === 'ja' && !state.nachDatum) {
      var ndEl = document.getElementById('kva-in-nachdatum');
      if (ndEl) {
        ndEl.style.borderColor = '#f85149';
        ndEl.focus();
      }
      setHudStatus('err', '⚠️ Nachuntersuchung Ja ist aktiv – bitte das Nachuntersuchungs-Datum eintragen!');
      return;
    }

    state.busy = true;
    renderHud();
    setHudStatus('wait', '1/3 · Erstelle Reparaturauftrag & speichere ' + state.stockId + '.pdf im Drive…');

    var payload = {
      action: 'create',
      stockId: state.stockId,
      modell: state.modell,
      vin: state.vin,
      erstzulassung: state.erstzulassung,
      getriebe: state.getriebe,
      km: state.km,
      kba: state.kba,
      workRows: Object.keys(state.workSel).map(Number),
      dRows: Object.keys(state.dSel).map(Number),
      oilRow: Number(state.oilRow) || 0,
      liters: state.liters,
      freeMech: state.freeMech.filter(Boolean),
      freeParts: state.freeParts.filter(Boolean),
      bodyWorks: state.bodyWorks.map(function (b) {
        return { side: b.side, row: b.row, actions: b.actions, label: b.label || '' };
      }),
      nach: state.nach,
      nachDatum: state.nachDatum,
      runUebertrag: state.runUebertrag
    };

    callBackend(payload, async function (res) {
      if (!res || !res.success) {
        state.busy = false;
        renderHud();
        setHudStatus('err', 'Backend-Fehler: ' + (res && res.message || 'Unbekannt'));
        return;
      }

      try {
        if (state.autoUploadPdf && res.pdfB64) {
          setHudStatus('wait', '2/3 · Lade ' + res.fileName + ' automatisch in Carol hoch…');
          await autoUploadPdfInCarol(res.pdfB64, res.fileName);
        }

        if (state.autoOrderParts) {
          setHudStatus('wait', '3/3 · Setze Ersatzteile auf "Bestellt" & schließe Nachautorisierung ab…');
          await autoSetPartsOrderedInCarol(state.stockId);
        }

        state.busy = false;
        renderHud();
        setHudStatus(
          'ok',
          'Fertig! ' + res.fileName + ' im Drive gespeichert'
            + (state.autoUploadPdf ? ', in Carol hochgeladen' : '')
            + (state.autoOrderParts ? ' & Ersatzteile auf Bestellt gesetzt.' : '.')
        );
      } catch (stepErr) {
        state.busy = false;
        renderHud();
        setHudStatus('err', 'PDF im Drive gespeichert, aber Carol-Schritt meldet: ' + (stepErr.message || stepErr));
      }
    });
  }

  function injectStyles() {
    if (document.getElementById('kva-carol-hud-styles')) return;
    var st = document.createElement('style');
    st.id = 'kva-carol-hud-styles';
    st.textContent = ''
      + '#kva-fab{position:fixed;bottom:20px;right:20px;z-index:2147483640;display:flex;align-items:center;gap:8px;'
      + 'padding:11px 18px;border-radius:999px;background:linear-gradient(180deg,#f2cc60 0%,#e3b341 100%);'
      + 'color:#0d1117;font:700 13.5px/1 "Segoe UI",system-ui,sans-serif;border:1px solid #f2cc60;'
      + 'box-shadow:0 8px 26px rgba(0,0,0,.45);cursor:pointer;user-select:none;}'
      + '#kva-fab:hover{filter:brightness(1.06);transform:translateY(-1px);}'
      + '#kva-drawer{position:fixed;top:0;right:0;bottom:0;width:560px;max-width:96vw;z-index:2147483645;'
      + 'background:#0c1017;color:#eef2f7;border-left:1px solid #283246;box-shadow:-12px 0 38px rgba(0,0,0,.55);'
      + 'font:13px/1.4 "Segoe UI",system-ui,sans-serif;display:none;flex-direction:column;}'
      + '#kva-drawer.open{display:flex;}'
      + '.kva-head{display:flex;align-items:center;gap:10px;padding:12px 16px;background:#121824;border-bottom:1px solid #283246;}'
      + '.kva-title{font-size:15px;font-weight:800;color:#f2cc60;letter-spacing:.3px;flex:1;}'
      + '.kva-ibtn{background:#1b2333;border:1px solid #2d3a52;color:#eef2f7;border-radius:7px;padding:5px 10px;font-size:12px;font-weight:600;cursor:pointer;}'
      + '.kva-ibtn:hover{border-color:#f2cc60;color:#f2cc60;}'
      + '.kva-body{flex:1;overflow-y:auto;padding:14px 16px 24px;}'
      + '.kva-sec{background:#121824;border:1px solid #232d40;border-radius:10px;padding:12px;margin-bottom:12px;}'
      + '.kva-sec h4{margin:0 0 8px;font-size:11.5px;text-transform:uppercase;letter-spacing:.8px;color:#f2cc60;display:flex;justify-content:space-between;align-items:center;}'
      + '.kva-grid2{display:grid;grid-template-columns:1fr 1fr;gap:8px;}'
      + '.kva-grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;}'
      + '.kva-fld label{display:block;font-size:11px;color:#93a1b5;margin-bottom:3px;}'
      + '.kva-fld input,.kva-fld select{width:100%;background:#090c11;border:1px solid #283246;color:#eef2f7;border-radius:6px;padding:6px 8px;font-size:12.5px;outline:none;}'
      + '.kva-fld input:focus,.kva-fld select:focus{border-color:#f2cc60;}'
      + '.kva-chks{display:grid;grid-template-columns:1fr 1fr;gap:4px 8px;max-height:210px;overflow-y:auto;padding-right:4px;}'
      + '.kva-chk{display:flex;align-items:center;gap:6px;padding:4px 6px;border-radius:6px;font-size:12px;cursor:pointer;border:1px solid transparent;user-select:none;}'
      + '.kva-chk:hover{background:#1a2232;}'
      + '.kva-chk.on{background:rgba(242,204,96,.12);border-color:rgba(242,204,96,.35);}'
      + '.kva-chk.auto{background:rgba(63,185,80,.14);border-color:rgba(63,185,80,.45);}'
      + '.kva-chk input{accent-color:#f2cc60;width:14px;height:14px;margin:0;cursor:pointer;}'
      + '.kva-pill{font-size:10px;padding:1px 6px;border-radius:99px;background:rgba(63,185,80,.2);color:#56d364;margin-left:auto;font-weight:700;}'
      + '.kva-pill.err{background:rgba(248,81,73,.2);color:#f85149;}'
      + '.kva-items{list-style:none;margin:0;padding:0;max-height:130px;overflow-y:auto;}'
      + '.kva-items li{padding:5px 7px;border-bottom:1px dashed #232d40;font-size:12px;}'
      + '.kva-items li:last-child{border-bottom:0;}'
      + '.kva-bw-list{list-style:none;margin:8px 0 0;padding:0;}'
      + '.kva-bw-list li{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:6px 8px;border-radius:6px;background:#171f2e;border:1px solid #283246;margin-bottom:5px;font-size:12px;}'
      + '.kva-bw-list li.auto{border-color:rgba(63,185,80,.55);background:rgba(63,185,80,.1);}'
      + '.kva-bw-del{background:transparent;border:0;color:#f85149;font-size:16px;cursor:pointer;line-height:1;}'
      + '.kva-foot{padding:12px 16px;background:#121824;border-top:1px solid #283246;display:flex;flex-direction:column;gap:8px;}'
      + '.kva-cta{width:100%;padding:11px 14px;border-radius:9px;border:none;background:linear-gradient(180deg,#f2cc60 0%,#e3b341 100%);color:#0d1117;font-weight:800;font-size:13.5px;cursor:pointer;}'
      + '.kva-cta:disabled{opacity:.5;cursor:not-allowed;}'
      + '.kva-status{padding:8px 10px;border-radius:7px;font-size:12px;display:none;}'
      + '.kva-status.show{display:block;}'
      + '.kva-status.wait{background:rgba(88,166,255,.14);border:1px solid rgba(88,166,255,.4);color:#58a6ff;}'
      + '.kva-status.ok{background:rgba(63,185,80,.14);border:1px solid rgba(63,185,80,.4);color:#56d364;}'
      + '.kva-status.err{background:rgba(248,81,73,.14);border:1px solid rgba(248,81,73,.4);color:#f85149;}';
    (document.head || document.documentElement).appendChild(st);
  }

  function ensureDom() {
    if (!document.body) return;
    injectStyles();
    if (!document.getElementById('kva-fab')) {
      var fab = document.createElement('button');
      fab.id = 'kva-fab';
      fab.type = 'button';
      fab.innerHTML = '<span>⚡ KVA Auftrag</span>';
      fab.addEventListener('click', function () {
        state.open = !state.open;
        if (state.open) scanAndPopulateFromCarol(false);
        renderHud();
      });
      document.body.appendChild(fab);
    }
    if (!document.getElementById('kva-drawer')) {
      var dr = document.createElement('aside');
      dr.id = 'kva-drawer';
      document.body.appendChild(dr);
      renderHud();
    }
  }

  function setHudStatus(mode, msg) {
    var el = document.getElementById('kva-status-box');
    if (!el) return;
    if (!mode) {
      el.className = 'kva-status';
      el.textContent = '';
      return;
    }
    el.className = 'kva-status show ' + mode;
    el.textContent = msg;
  }

  function syncInputsFromDom() {
    var getVal = function (id, fallback) {
      var el = document.getElementById(id);
      return el ? el.value.trim() : fallback;
    };
    state.stockId = getVal('kva-in-sid', state.stockId).toUpperCase();
    state.modell = cleanModelString(getVal('kva-in-mod', state.modell));
    state.vin = getVal('kva-in-vin', state.vin).toUpperCase();
    state.erstzulassung = getVal('kva-in-ez', state.erstzulassung);
    state.getriebe = getVal('kva-in-getr', state.getriebe);
    state.km = getVal('kva-in-km', state.km);
    state.oilRow = parseInt(getVal('kva-in-oil', state.oilRow), 10) || 0;
    state.liters = getVal('kva-in-lit', state.liters);
    state.nachDatum = getVal('kva-in-nachdatum', state.nachDatum);
    for (var i = 0; i < 5; i++) {
      var fmEl = document.getElementById('kva-fm-' + i);
      if (fmEl) state.freeMech[i] = fmEl.value.trim();
    }
  }

  function renderHud() {
    var dr = document.getElementById('kva-drawer');
    if (!dr) return;
    dr.classList.toggle('open', !!state.open);
    if (!state.open) return;

    var meta = state.meta || FALLBACK_META;
    var url = getWebAppUrl();

    var huBadgeHtml = '';
    if (state.huStatus === 'failed') {
      huBadgeHtml = '<span class="kva-pill err">HU NICHT BESTANDEN ✕</span>';
    } else if (state.huStatus === 'passed') {
      huBadgeHtml = '<span class="kva-pill">HU BESTANDEN ✓</span>';
    }

    var carolListHtml = '';
    if (state.carolItems.length) {
      carolListHtml = '<ul class="kva-items">' + state.carolItems.map(function (it) {
        return '<li><b>' + esc(it.title) + '</b>' + (it.comment ? (' <span style="color:#93a1b5;">— ' + esc(it.comment) + '</span>') : '') + '</li>';
      }).join('') + '</ul>';
    } else {
      carolListHtml = '<div style="color:#93a1b5;font-size:12px;">Noch keine Positionen erkannt — klicke oben auf "Neu einlesen".</div>';
    }

    var workHtml = (meta.workItems || []).filter(function (it) { return it.row !== 10; }).map(function (it) {
      var on = !!state.workSel[it.row];
      var isAuto = !!state.autoWork[it.row];
      return '<label class="kva-chk' + (on ? ' on' : '') + (isAuto ? ' auto' : '') + '">'
        + '<input type="checkbox" data-kind="work" data-row="' + it.row + '"' + (on ? ' checked' : '') + '>'
        + '<span>' + esc(it.label) + '</span>'
        + (isAuto ? '<span class="kva-pill">AUTO</span>' : '')
        + '</label>';
    }).join('');

    var dHtml = (meta.dItems || []).map(function (it) {
      var on = !!state.dSel[it.row];
      var isAuto = !!state.autoD[it.row];
      return '<label class="kva-chk' + (on ? ' on' : '') + (isAuto ? ' auto' : '') + '">'
        + '<input type="checkbox" data-kind="d" data-row="' + it.row + '"' + (on ? ' checked' : '') + '>'
        + '<span>' + esc(it.label) + '</span>'
        + (isAuto ? '<span class="kva-pill">AUTO</span>' : '')
        + '</label>';
    }).join('');

    var partOpts = '<option value="">— Bauteil wählen —</option>' + (meta.parts || []).map(function (p) {
      return '<option value="' + p.side + ':' + p.row + '">' + esc(p.label) + '</option>';
    }).join('');

    var actChecks = (meta.actions || FALLBACK_META.actions).map(function (a) {
      return '<label class="kva-chk"><input type="checkbox" class="kva-bw-act-cb" value="' + a.nr + '"><span>' + esc(a.nr + ' · ' + a.label) + '</span></label>';
    }).join('');

    var bwListHtml = '';
    if (state.bodyWorks.length) {
      bwListHtml = '<ul class="kva-bw-list">' + state.bodyWorks.map(function (bw, idx) {
        return '<li class="' + (bw.auto ? 'auto' : '') + '">'
          + '<span><b>' + esc((bw.actionLabels || []).join(', ')) + '</b> <span style="color:#f2cc60;">→ ' + esc(bw.label) + '</span></span>'
          + '<span style="display:flex;align-items:center;gap:6px;">'
          + (bw.auto ? '<span class="kva-pill">AUTO</span>' : '')
          + '<button type="button" class="kva-bw-del" data-bw-del="' + idx + '" title="Entfernen">×</button>'
          + '</span></li>';
      }).join('') + '</ul>';
    } else {
      bwListHtml = '<div style="color:#93a1b5;font-size:12px;">Keine Karosserie-/Lack-/Dellen-Arbeiten in Carol beauftragt.</div>';
    }

    var oilOpts = '<option value="0">— Kein Öl ausgewählt —</option>' + (meta.oils || []).map(function (o) {
      var sel = Number(state.oilRow) === Number(o.row) ? ' selected' : '';
      var prefix = o.nr ? ('Öl Nr. ' + String(o.nr).replace(/^Öl\s*Nr\.?\s*/i, '') + ' · ') : '';
      return '<option value="' + o.row + '"' + sel + '>' + esc(prefix + o.label) + '</option>';
    }).join('');

    var settingsHtml = '';
    if (state.showSettings || !url) {
      settingsHtml = '<div class="kva-sec" style="border-color:#f2cc60;">'
        + '<h4>Backend-Verbindung (Apps Script Web-App URL)</h4>'
        + '<div class="kva-fld">'
        + '<label>Bereitgestellte Web-App URL aus "Werkstattauftrag Tobi"</label>'
        + '<div style="display:flex;gap:6px;">'
        + '<input type="text" id="kva-cfg-url" value="' + esc(url) + '" placeholder="https://script.google.com/macros/s/.../exec">'
        + '<button type="button" class="kva-ibtn" id="kva-cfg-save">Speichern</button>'
        + '</div></div></div>';
    }

    dr.innerHTML = ''
      + '<div class="kva-head">'
      + '<div class="kva-title">⚡ KVA Werkstattauftrag <span style="font-size:11px;color:#93a1b5;font-weight:500;">v' + VER + '</span></div>'
      + '<button type="button" class="kva-ibtn" id="kva-btn-rescan">↻ Neu einlesen</button>'
      + '<button type="button" class="kva-ibtn" id="kva-btn-cfg" title="Einstellungen">⚙</button>'
      + '<button type="button" class="kva-ibtn" id="kva-btn-close">✕</button>'
      + '</div>'
      + '<div class="kva-body">'
      + settingsHtml
      + '<div id="kva-status-box" class="kva-status"></div>'
      + '<div class="kva-sec">'
      + '<h4><span>Fahrzeugdaten (aus Carol)</span>' + huBadgeHtml + '</h4>'
      + '<div class="kva-grid3" style="margin-bottom:8px;">'
      + '<div class="kva-fld"><label>Stock ID</label><input id="kva-in-sid" value="' + esc(state.stockId) + '"></div>'
      + '<div class="kva-fld" style="grid-column:span 2;"><label>Modell</label><input id="kva-in-mod" value="' + esc(state.modell) + '"></div>'
      + '</div>'
      + '<div class="kva-grid3">'
      + '<div class="kva-fld"><label>VIN</label><input id="kva-in-vin" value="' + esc(state.vin) + '"></div>'
      + '<div class="kva-fld"><label>KM-Stand</label><input id="kva-in-km" value="' + esc(state.km) + '"></div>'
      + '<div class="kva-fld"><label>Getriebe</label><input id="kva-in-getr" value="' + esc(state.getriebe) + '"></div>'
      + '</div>'
      + '<input type="hidden" id="kva-in-ez" value="' + esc(state.erstzulassung) + '">'
      + '</div>'
      + '<div class="kva-sec">'
      + '<h4><span>Erkannte Carol-Positionen (' + state.carolItems.length + ')</span></h4>'
      + carolListHtml
      + '</div>'
      + '<div class="kva-sec">'
      + '<h4><span>Karosserie, Lack &amp; Delle (Automatisch aus Carol · ' + state.bodyWorks.length + ')</span></h4>'
      + bwListHtml
      + '<details style="margin-top:8px;">'
      + '<summary style="cursor:pointer;color:#93a1b5;font-size:11.5px;user-select:none;">+ Optional manuell Bauteil hinzufügen</summary>'
      + '<div style="margin-top:8px;">'
      + '<div class="kva-grid2" style="margin-bottom:6px;">'
      + '<div class="kva-fld"><select id="kva-bw-part">' + partOpts + '</select></div>'
      + '<button type="button" class="kva-ibtn" id="kva-bw-add">+ Hinzufügen</button>'
      + '</div>'
      + '<div class="kva-chks" style="max-height:95px;">' + actChecks + '</div>'
      + '</div>'
      + '</details>'
      + '</div>'
      + '<div class="kva-sec">'
      + '<h4><span>Mechanik &amp; Zusatzarbeiten</span></h4>'
      + '<div class="kva-chks">' + dHtml + '</div>'
      + '</div>'
      + '<div class="kva-sec">'
      + '<h4><span>Wartung &amp; Service</span><button type="button" class="kva-ibtn" id="kva-btn-srvpkg" style="padding:2px 8px;font-size:11px;">+ Service Paket</button></h4>'
      + '<div class="kva-chks">' + workHtml + '</div>'
      + '</div>'
      + '<div class="kva-sec">'
      + '<h4><span>Öl Nr. &amp; Nachuntersuchung</span>'
      + (state.huDueDate ? ('<span class="kva-pill' + (state.huStatus === 'failed' ? ' err' : '') + '">HU Fällig: ' + esc(state.huDueDate) + '</span>') : '')
      + '</h4>'
      + '<div class="kva-grid2" style="margin-bottom:8px;">'
      + '<div class="kva-fld"><label>Öl (für G9:G12)' + ((state.workSel[12] || state.workSel[14]) ? ' <b style="color:#f85149;">* Pflicht</b>' : '') + '</label><select id="kva-in-oil">' + oilOpts + '</select></div>'
      + '<div class="kva-fld"><label id="kva-lbl-lit">Füllmenge (Liter)' + (isOilLitersRequired() ? ' <b style="color:#f85149;">* PFLICHT</b>' : '') + '</label>'
      + '<input id="kva-in-lit" value="' + esc(state.liters) + '" placeholder="' + (isOilLitersRequired() ? 'PFLICHT (z.B. 4,5)' : 'z.B. 4,5') + '"'
      + (isOilLitersRequired() && !hasValidLiters() ? ' style="border-color:#f85149;background:rgba(248,81,73,.08);"' : '') + '></div>'
      + '</div>'
      + (isOilLitersRequired() && !hasValidLiters()
        ? '<div style="margin-bottom:8px;padding:6px 9px;border-radius:6px;background:rgba(248,81,73,.14);border:1px solid rgba(248,81,73,.45);color:#f85149;font-size:11.5px;font-weight:600;">⚠️ Öl ausgewählt: Bitte Füllmenge (Liter) eintragen, sonst erscheint #N/A auf dem Werkstattauftrag!</div>'
        : '')
      + '<div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;font-size:12px;">'
      + '<span>GTÜ Nachuntersuchung:</span>'
      + '<label class="kva-chk"><input type="radio" name="kva-nach" value="nein"' + (state.nach === 'nein' ? ' checked' : '') + '><span>Nein</span></label>'
      + '<label class="kva-chk"><input type="radio" name="kva-nach" value="ja"' + (state.nach === 'ja' ? ' checked' : '') + '><span>Ja, bis:</span></label>'
      + '<input type="text" id="kva-in-nachdatum" value="' + esc(state.nachDatum) + '" placeholder="TT.MM.JJJJ" style="width:110px;background:#090c11;border:1px solid ' + (state.nach === 'ja' && !state.nachDatum ? '#f85149' : '#283246') + ';color:#eef2f7;border-radius:6px;padding:4px 7px;font-size:12px;">'
      + '<label class="kva-chk"><input type="radio" name="kva-nach" value=""' + (state.nach === '' ? ' checked' : '') + '><span>Aus</span></label>'
      + '</div>'
      + '</div>'
      + '<div class="kva-sec">'
      + '<h4><span>Sonstige Mechanik / Freitext (D24:D28)</span></h4>'
      + '<div class="kva-fld" style="display:flex;flex-direction:column;gap:5px;">'
      + '<input id="kva-fm-0" placeholder="Freitext Zeile 1 (D24)" value="' + esc(state.freeMech[0] || '') + '">'
      + '<input id="kva-fm-1" placeholder="Freitext Zeile 2 (D25)" value="' + esc(state.freeMech[1] || '') + '">'
      + '<input id="kva-fm-2" placeholder="Freitext Zeile 3 (D26)" value="' + esc(state.freeMech[2] || '') + '">'
      + '<input id="kva-fm-3" placeholder="Freitext Zeile 4 (D27)" value="' + esc(state.freeMech[3] || '') + '">'
      + '<input id="kva-fm-4" placeholder="Freitext Zeile 5 (D28)" value="' + esc(state.freeMech[4] || '') + '">'
      + '</div>'
      + '</div>'
      + '<div class="kva-sec">'
      + '<h4><span>Automatisierungs-Schritte</span></h4>'
      + '<label class="kva-chk"><input type="checkbox" id="kva-opt-upload"' + (state.autoUploadPdf ? ' checked' : '') + '><span>PDF (<b id="kva-fn-lbl">' + esc((state.stockId || 'STOCK') + '.pdf') + '</b>) automatisch in Carol unter Dokumente hochladen</span></label>'
      + '<label class="kva-chk"><input type="checkbox" id="kva-opt-parts"' + (state.autoOrderParts ? ' checked' : '') + '><span>Ersatzteile in Nachautorisierung automatisch auf "Bestellt" setzen &amp; abschließen</span></label>'
      + '<label class="kva-chk"><input type="checkbox" id="kva-opt-ueb"' + (state.runUebertrag ? ' checked' : '') + '><span>Übertrag (Nachtrag-Sheet) direkt mit ausführen</span></label>'
      + '</div>'
      + '</div>'
      + '<div class="kva-foot">'
      + '<button type="button" class="kva-cta" id="kva-btn-run"' + (state.busy ? ' disabled' : '') + '>'
      + (state.busy ? '⏳ Verarbeitung läuft…' : '🚀 Auftrag erstellen, PDF in Carol hochladen & Abschließen')
      + '</button>'
      + '<div style="display:flex;gap:6px;">'
      + '<button type="button" class="kva-ibtn" id="kva-btn-only-parts" style="flex:1;">📦 Nur Ersatzteile -> Bestellt</button>'
      + '<button type="button" class="kva-ibtn" id="kva-btn-reset-sheet">Sheet Reset</button>'
      + '</div>'
      + '</div>';

    bindHudEvents();
  }

  function bindHudEvents() {
    var closeBtn = document.getElementById('kva-btn-close');
    if (closeBtn) closeBtn.addEventListener('click', function () { state.open = false; renderHud(); });

    var rescanBtn = document.getElementById('kva-btn-rescan');
    if (rescanBtn) rescanBtn.addEventListener('click', function () {
      scanAndPopulateFromCarol(true);
      setHudStatus('ok', 'Fahrzeug-, Lack- und Service-Daten aus Carol neu eingelesen.');
    });

    var cfgBtn = document.getElementById('kva-btn-cfg');
    if (cfgBtn) cfgBtn.addEventListener('click', function () {
      syncInputsFromDom();
      state.showSettings = !state.showSettings;
      renderHud();
    });

    var cfgSave = document.getElementById('kva-cfg-save');
    if (cfgSave) cfgSave.addEventListener('click', function () {
      var v = String(document.getElementById('kva-cfg-url').value || '').trim();
      setStored(STORE_URL_KEY, v);
      state.showSettings = false;
      fetchBackendMeta();
      renderHud();
      setHudStatus('ok', 'Web-App URL gespeichert.');
    });

    var srvPkg = document.getElementById('kva-btn-srvpkg');
    if (srvPkg) srvPkg.addEventListener('click', function () {
      syncInputsFromDom();
      (state.meta.workItems || []).forEach(function (it) {
        var l = it.label.toLowerCase();
        if (l.indexOf('motorölwechsel') !== -1 || l.indexOf('innenraum') !== -1 || l.indexOf('luftfilter (motor)') !== -1) {
          state.workSel[it.row] = true;
        }
      });
      renderHud();
    });

    var bwAdd = document.getElementById('kva-bw-add');
    if (bwAdd) bwAdd.addEventListener('click', function () {
      syncInputsFromDom();
      var partVal = document.getElementById('kva-bw-part').value;
      if (!partVal) return;
      var sp = partVal.split(':');
      var side = parseInt(sp[0], 10);
      var row = parseInt(sp[1], 10);
      var partObj = null;
      ((state.meta && state.meta.parts) || FALLBACK_META.parts).forEach(function (p) {
        if (p.side === side && p.row === row) partObj = p;
      });
      if (!partObj) return;
      var acts = [];
      document.querySelectorAll('.kva-bw-act-cb:checked').forEach(function (cb) {
        acts.push(parseInt(cb.value, 10));
      });
      if (!acts.length) return;
      addOrMergeAutoBodyWork(partObj, acts);
      state.bodyWorks[state.bodyWorks.length - 1].auto = false;
      renderHud();
    });

    var dr = document.getElementById('kva-drawer');
    dr.querySelectorAll('[data-bw-del]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        syncInputsFromDom();
        var idx = parseInt(btn.getAttribute('data-bw-del'), 10);
        state.bodyWorks.splice(idx, 1);
        renderHud();
      });
    });

    dr.querySelectorAll('input[data-kind]').forEach(function (cb) {
      cb.addEventListener('change', function () {
        syncInputsFromDom();
        var kind = cb.getAttribute('data-kind');
        var row = Number(cb.getAttribute('data-row'));
        var map = kind === 'work' ? state.workSel : state.dSel;
        if (cb.checked) map[row] = true;
        else delete map[row];
        cb.parentElement.classList.toggle('on', cb.checked);
        if (kind === 'work' && (row === 12 || row === 14)) {
          renderHud();
        }
      });
    });

    var oilSelectEl = document.getElementById('kva-in-oil');
    if (oilSelectEl) {
      oilSelectEl.addEventListener('change', function () {
        syncInputsFromDom();
        if (Number(state.oilRow) >= 29 && Number(state.oilRow) <= 41) {
          state.workSel[14] = true;
        }
        renderHud();
        if (isOilLitersRequired() && !hasValidLiters()) {
          var litInput = document.getElementById('kva-in-lit');
          if (litInput) litInput.focus();
        }
      });
    }

    var litInputEl = document.getElementById('kva-in-lit');
    if (litInputEl) {
      litInputEl.addEventListener('input', function () {
        state.liters = litInputEl.value.trim();
        if (hasValidLiters()) {
          litInputEl.style.borderColor = '#283246';
          litInputEl.style.background = '#090c11';
        }
      });
    }

    dr.querySelectorAll('input[name="kva-nach"]').forEach(function (rb) {
      rb.addEventListener('change', function () {
        syncInputsFromDom();
        if (rb.checked) {
          state.nach = rb.value;
          if (state.nach === 'ja' && !state.nachDatum && state.huDueDate) {
            state.nachDatum = state.huDueDate;
          }
          renderHud();
        }
      });
    });

    var optUpload = document.getElementById('kva-opt-upload');
    if (optUpload) optUpload.addEventListener('change', function () { state.autoUploadPdf = optUpload.checked; });

    var optParts = document.getElementById('kva-opt-parts');
    if (optParts) optParts.addEventListener('change', function () { state.autoOrderParts = optParts.checked; });

    var optUeb = document.getElementById('kva-opt-ueb');
    if (optUeb) optUeb.addEventListener('change', function () { state.runUebertrag = optUeb.checked; });

    var runBtn = document.getElementById('kva-btn-run');
    if (runBtn) runBtn.addEventListener('click', runFullAutomation);

    var onlyPartsBtn = document.getElementById('kva-btn-only-parts');
    if (onlyPartsBtn) onlyPartsBtn.addEventListener('click', function () {
      autoSetPartsOrderedInCarol(state.stockId);
    });

    var resetSheetBtn = document.getElementById('kva-btn-reset-sheet');
    if (resetSheetBtn) resetSheetBtn.addEventListener('click', function () {
      setHudStatus('wait', 'Setze Google Sheet zurück…');
      callBackend({ action: 'reset' }, function (res) {
        if (res && res.success) setHudStatus('ok', 'Sheet wurde zurückgesetzt.');
        else setHudStatus('err', (res && res.message) || 'Fehler beim Reset.');
      });
    });
  }

  function checkPendingAutoPartsSession() {
    try {
      var pendingSid = sessionStorage.getItem(STORE_AUTO_PARTS_KEY);
      if (pendingSid && /nachautorisierung/i.test(getPageTextWithoutDrawer())) {
        sessionStorage.removeItem(STORE_AUTO_PARTS_KEY);
        setTimeout(function () {
          autoSetPartsOrderedInCarol(pendingSid);
        }, 700);
      }
    } catch (e) {}
  }

  injectGraphqlHook();

  function boot() {
    if (!document.body) {
      setTimeout(boot, 250);
      return;
    }
    ensureDom();
    if (getWebAppUrl()) fetchBackendMeta();

    var lastHref = '';
    setInterval(function () {
      ensureDom();
      checkPendingAutoPartsSession();
      if (location.href !== lastHref) {
        lastHref = location.href;
        if (isRefurbDetailPage()) {
          setTimeout(function () { scanAndPopulateFromCarol(false); }, 900);
        }
      }
    }, 1200);
  }

  boot();
})();
