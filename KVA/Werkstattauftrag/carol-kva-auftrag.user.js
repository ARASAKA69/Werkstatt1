// ==UserScript==
// @name         Carol KVA Werkstattauftrag Automation
// @namespace    arasaka-kva
// @version      1.0.0
// @description  In-Carol KVA HUD: Liest Fahrzeug & Services aus Carol, erstellt Reparaturauftrag-PDF im Sheet/Drive, lädt PDF in Carol hoch und setzt Ersatzteile auf Bestellt
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

  var VER = '1.0.0';
  if (window.__kvaCarolHudVer) return;
  window.__kvaCarolHudVer = VER;

  var DEFAULT_WEB_APP_URL = '';
  var BRIDGE_SECRET = 'ARASAKA69';
  var MSG_MARKER = '__kvaBridgeGraphql';
  var STORE_URL_KEY = 'kva_wa_webapp_url_v1';
  var STORE_META_KEY = 'kva_wa_meta_cache_v1';
  var STORE_AUTO_PARTS_KEY = 'kva_wa_auto_parts_sid';

  var FALLBACK_META = {
    workItems: [
      { row: 11, label: 'Inspektion' },
      { row: 12, label: 'Motorölwechsel inkl. Filterersatz' },
      { row: 14, label: 'Bremsflüssigkeit wechseln' },
      { row: 15, label: 'Innenraum Filter ersetzen' },
      { row: 16, label: 'Luftfilter (Motor) ersetzen' },
      { row: 17, label: 'Luftfilter (Hochvolt) ersetzen' },
      { row: 18, label: 'Kraftstofffilter ersetzen' },
      { row: 19, label: 'Zündkerzen ersetzen' },
      { row: 20, label: 'Kühlflüssigkeit wechseln' },
      { row: 21, label: 'Getriebeöl wechseln inkl. Filterersatz' },
      { row: 22, label: 'Haldex Öl wechseln' },
      { row: 23, label: 'Vorderachsdifferenzial Öl wechseln' },
      { row: 24, label: 'Hinterachsdifferenzial Öl wechseln' },
      { row: 25, label: 'Keilriemen ersetzen' },
      { row: 26, label: 'Zahnriemen inkl. Wasserpumpe ersetzen' },
      { row: 27, label: 'Schlüsselbatterie ersetzen' },
      { row: 28, label: 'Schiebedach / Panoramadach schmieren' }
    ],
    oils: [
      { row: 29, label: '7000 AV 0W-20', nr: 'Öl Nr. 1' },
      { row: 30, label: '7000 E 0W-30', nr: 'Öl Nr. 2' },
      { row: 31, label: '7000 DM 0W-30', nr: 'Öl Nr. 3' },
      { row: 32, label: '7000 0W-40', nr: 'Öl Nr. 4' },
      { row: 33, label: '5000 AV 5W-30', nr: 'Öl Nr. 5' },
      { row: 34, label: '5000 XS 5W-30', nr: 'Öl Nr. 6' },
      { row: 35, label: '5000 CP 5W-30', nr: 'Öl Nr. 7' },
      { row: 36, label: '5000 RN 5W-30', nr: 'Öl Nr. 8' },
      { row: 37, label: '5000 FR 5W-30', nr: 'Öl Nr. 9' },
      { row: 38, label: '5000 FR 5W-20', nr: 'Öl Nr. 10' },
      { row: 39, label: '3000 E 5W-40', nr: 'Öl Nr. 11' },
      { row: 40, label: '3000 AV 5W-40', nr: 'Öl Nr. 12' },
      { row: 41, label: 'Öl wird mitgeliefert', nr: 'Öl mitgeliefert' }
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
      { row: 10, label: 'Bremsbeläge vorne ersetzen' },
      { row: 11, label: 'Bremsscheiben und -beläge hinten ersetzen' },
      { row: 12, label: 'Bremsbeläge hinten ersetzen' },
      { row: 13, label: 'Klimaservice' },
      { row: 14, label: 'Antenne ersetzen' },
      { row: 15, label: 'Ozon-Behandlung' },
      { row: 16, label: 'Diagnose durchführen' },
      { row: 17, label: 'Ölverlust prüfen' },
      { row: 18, label: 'Fehlerspeicher löschen' },
      { row: 19, label: 'Scheibenwischer vorne links erneuern' },
      { row: 20, label: 'Scheibenwischer vorne rechts erneuern' },
      { row: 21, label: 'Heckwischer erneuern' },
      { row: 22, label: 'Marderschutz entfernen' },
      { row: 23, label: 'Windschutzscheibe ersetzen' }
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
      { side: 1, row: 7, label: 'Außenspiegel links' },
      { side: 1, row: 8, label: 'Außenspiegel rechts' },
      { side: 1, row: 15, label: 'Dach' },
      { side: 1, row: 22, label: 'Frontscheibe' },
      { side: 1, row: 24, label: 'Heckklappe' },
      { side: 1, row: 29, label: 'Kotflügel hinten links' },
      { side: 1, row: 30, label: 'Kotflügel hinten rechts' },
      { side: 1, row: 31, label: 'Kotflügel vorne links' },
      { side: 1, row: 32, label: 'Kotflügel vorne rechts' },
      { side: 2, row: 5, label: 'Motorhaube' },
      { side: 2, row: 16, label: 'Schweller links' },
      { side: 2, row: 17, label: 'Schweller rechts' },
      { side: 2, row: 25, label: 'Stoßfänger hinten' },
      { side: 2, row: 26, label: 'Stoßfänger vorne' },
      { side: 2, row: 31, label: 'Tür hinten links' },
      { side: 2, row: 32, label: 'Tür hinten rechts' },
      { side: 2, row: 33, label: 'Tür vorne links' },
      { side: 2, row: 34, label: 'Tür vorne rechts' }
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
    huPassed: false,
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
      if (parsed && parsed.workItems && parsed.dItems) return parsed;
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
      var full = strVal(node.fullName) || strVal(node.title) || [make, model, sub].filter(Boolean).join(' ');
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
      .replace(/[^a-z0-9]+/g, ' ').trim();
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
    return (' ' + hay + ' ').indexOf(' ' + normTxt(phrase) + ' ') !== -1;
  }

  function isRefurbDetailPage() {
    return /\/refurbishment\/[0-9a-f-]{8,}/i.test(location.pathname);
  }

  function expandCollapsedCarolServices() {
    var btns = Array.from(document.querySelectorAll('button, a, span, div'));
    for (var i = 0; i < btns.length; i++) {
      var el = btns[i];
      if (el.children && el.children.length > 2) continue;
      var t = String(el.textContent || '').trim();
      if (/^(zeige\s+alle\s+services|show\s+all\s+services)$/i.test(t) && !el.__kvaExpanded) {
        el.__kvaExpanded = true;
        try { el.click(); } catch (e) {}
      }
    }
  }

  function scrapeCarolDom() {
    var bodyText = document.body ? document.body.innerText : '';
    var out = {
      stockId: '',
      modell: '',
      vin: '',
      erstzulassung: '',
      getriebe: '',
      km: '',
      huPassed: false,
      items: []
    };

    var headings = document.querySelectorAll('h1, h2, h3, [class*="title"], [class*="Title"], [class*="Header"]');
    for (var i = 0; i < headings.length; i++) {
      var hTxt = String(headings[i].textContent || '').replace(/\s+/g, ' ').trim();
      var m = hTxt.match(/\b([A-Z]{2}\d{4,8})\s*[-–]\s*([^\n|]+)/);
      if (m) {
        out.stockId = m[1].toUpperCase();
        out.modell = m[2].trim();
        break;
      }
    }
    if (!out.stockId) {
      var mSid = bodyText.match(/\b([A-Z]{2}\d{5,8})\s*[-–]\s*([^\n]+)/);
      if (mSid) {
        out.stockId = mSid[1].toUpperCase();
        out.modell = mSid[2].trim();
      }
    }

    var mVin = bodyText.match(/VIN[:\s]*([A-HJ-NPR-Z0-9]{17})/i);
    if (mVin) out.vin = mVin[1].toUpperCase();

    var mKm = bodyText.match(/Laufleistung[\s\n:]*([\d.,]+)\s*KM/i) || bodyText.match(/Mileage[\s\n:]*([\d.,]+)\s*KM/i);
    if (mKm) out.km = mKm[1].replace(/[.,\s]/g, '');

    var mGetr = bodyText.match(/Getriebeart[\s\n:]*([^\n\r]+)/i) || bodyText.match(/Transmission[\s\n:]*([^\n\r]+)/i);
    if (mGetr) out.getriebe = mGetr[1].trim();

    var mEz = bodyText.match(/Datum\s+Erstzulassung[\s\n:]*([^\n\r]+)/i) || bodyText.match(/First\s+registration[\s\n:]*([^\n\r]+)/i);
    if (mEz) out.erstzulassung = mEz[1].trim();

    if (state.graphqlCapture.vin && (!out.vin || out.vin === state.graphqlCapture.vin)) {
      if (!out.vin) out.vin = state.graphqlCapture.vin;
      if (!out.stockId && state.graphqlCapture.stockId) out.stockId = state.graphqlCapture.stockId;
      if (!out.modell && state.graphqlCapture.modell) out.modell = state.graphqlCapture.modell;
    }

    if (/HU\s+Bestanden|HU\s+Bericht/i.test(bodyText)) {
      out.huPassed = true;
    }

    var seenItems = {};
    function addItem(title, comment, source) {
      var t = String(title || '').replace(/\s+/g, ' ').trim();
      var c = String(comment || '').replace(/\s+/g, ' ').trim();
      if (!t && !c) return;
      var combined = [t, c].filter(Boolean).join(' — ');
      if (/remove\s+license\s+plate\s+holder|kennzeichenhalter\s+entfernen/i.test(combined)) return;
      if (/^(leistung|service|dienstleistung|fahrzeugkomponente|kommentar|status|kosten)$/i.test(t)) return;
      var key = normTxt(t + '|' + c);
      if (!key || seenItems[key]) return;
      seenItems[key] = 1;
      out.items.push({ title: t, comment: c, source: source || 'Carol' });
    }

    var tables = document.querySelectorAll('table');
    for (var tb = 0; tb < tables.length; tb++) {
      var rows = tables[tb].querySelectorAll('tbody tr, tr');
      for (var r = 0; r < rows.length; r++) {
        var cells = rows[r].querySelectorAll('td');
        if (cells.length >= 2) {
          var c0 = String(cells[0].innerText || '').trim();
          var c1 = String(cells[1].innerText || '').trim();
          var c2 = cells[2] ? String(cells[2].innerText || '').trim() : '';
          if (c0 && !/^(image\.jpg|.*\.pdf)$/i.test(c0) && c0.length > 2) {
            if (/^\s*$/.test(c0) && c1) addItem(c1, c2, 'Ersatzteile');
            else addItem(c0, c1, 'Ersatzteile');
          }
        }
      }
    }

    var blocks = bodyText.split(/\n+/);
    for (var b = 0; b < blocks.length; b++) {
      var line = blocks[b].trim();
      if (/^(Fahrzeugkomponente|Vehicle\s+Component)$/i.test(line)) {
        var comp = (blocks[b + 1] || '').trim();
        var srv = '';
        var comm = '';
        for (var k = b + 1; k < Math.min(blocks.length, b + 14); k++) {
          if (/^(Dienstleistung|Service)$/i.test(blocks[k].trim())) srv = (blocks[k + 1] || '').trim();
          if (/^(Kommentar|Comment)$/i.test(blocks[k].trim())) comm = (blocks[k + 1] || '').trim();
        }
        var label = comp;
        if (srv && !/^andere|other$/i.test(srv)) label = comp + ' ' + srv;
        addItem(label, comm, 'Schaden / Komponente');
      } else if (/^(Service|Dienstleistung)$/i.test(line)) {
        var sVal = (blocks[b + 1] || '').trim();
        var sComm = '';
        for (var k2 = b + 1; k2 < Math.min(blocks.length, b + 10); k2++) {
          if (/^(Kommentar|Comment)$/i.test(blocks[k2].trim())) {
            sComm = (blocks[k2 + 1] || '').trim();
            break;
          }
        }
        if (sVal && !/^(andere|other)$/i.test(sVal)) {
          addItem(sVal, sComm, 'Service');
        } else if (/^(andere|other)$/i.test(sVal) && sComm) {
          addItem(sComm, '', 'Service');
        }
      }
    }

    return out;
  }

  function findItemByTokens(items, tokens) {
    if (!items || !tokens || !tokens.length) return null;
    for (var i = 0; i < items.length; i++) {
      var n = normTxt(items[i].label);
      var ok = true;
      for (var t = 0; t < tokens.length; t++) {
        if (!hasWord(n, tokens[t]) && n.indexOf(tokens[t]) === -1) { ok = false; break; }
      }
      if (ok) return items[i];
    }
    return null;
  }

  function expandText(s) {
    var n = ' ' + normTxt(s) + ' ';
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
    n = n.replace(/\bengine\s+oil\s+including\s+filter\s+change\b/g, ' motoroelwechsel ');
    n = n.replace(/\binterior\s+filter\s+change\b/g, ' innenraumfilter ');
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
      { kind: 'd', tokens: ['batterie', 'ersetzen'], words: ['batterie ersetzen', 'batterie erneuern', 'starterbatterie'] },
      { kind: 'd', tokens: ['bremsscheiben', 'vorne'], all: ['bremsscheiben', 'vorne'] },
      { kind: 'd', tokens: ['bremsscheiben', 'hinten'], all: ['bremsscheiben', 'hinten'] },
      { kind: 'd', tokens: ['bremsbelaege', 'vorne'], all: ['bremsbelaege', 'vorne'] },
      { kind: 'd', tokens: ['bremsbelaege', 'hinten'], all: ['bremsbelaege', 'hinten'] },
      { kind: 'd', tokens: ['reifen', 'vorne', 'rechts'], all: ['reifen', 'vorne', 'rechts'] },
      { kind: 'd', tokens: ['reifen', 'vorne', 'links'], all: ['reifen', 'vorne', 'links'] },
      { kind: 'd', tokens: ['reifen', 'hinten', 'rechts'], all: ['reifen', 'hinten', 'rechts'] },
      { kind: 'd', tokens: ['reifen', 'hinten', 'links'], all: ['reifen', 'hinten', 'links'] },
      { kind: 'd', tokens: ['radsatz'], words: ['radsatz', 'raederwechsel'] },
      { kind: 'd', tokens: ['scheibenwischer', 'links'], all: ['scheibenwischer', 'links'] },
      { kind: 'd', tokens: ['scheibenwischer', 'rechts'], all: ['scheibenwischer', 'rechts'] },
      { kind: 'd', tokens: ['heckwischer'], words: ['heckwischer'] },
      { kind: 'd', tokens: ['windschutzscheibe'], words: ['windschutzscheibe', 'frontscheibe'] },
      { kind: 'd', tokens: ['klima'], words: ['klimaservice', 'klimaanlage'] },
      { kind: 'd', tokens: ['ozon'], words: ['ozon'] },
      { kind: 'd', tokens: ['diagnose'], words: ['diagnose'] },
      { kind: 'd', tokens: ['fehlerspeicher'], words: ['fehlerspeicher'] },
      { kind: 'd', tokens: ['oelverlust'], words: ['oelverlust'] },
      { kind: 'd', tokens: ['marderschutz'], words: ['marderschutz'] },
      { kind: 'work', tokens: ['inspektion'], words: ['inspektion'] },
      { kind: 'work', tokens: ['motoroelwechsel'], words: ['motoroelwechsel', 'oelwechsel', 'motorolwechsel'] },
      { kind: 'work', tokens: ['bremsfluessigkeit'], words: ['bremsfluessigkeit'] },
      { kind: 'work', tokens: ['innenraum'], words: ['innenraumfilter', 'innenraum filter', 'pollenfilter'] },
      { kind: 'work', tokens: ['luftfilter', 'motor'], words: ['luftfilter'] },
      { kind: 'work', tokens: ['kraftstofffilter'], words: ['kraftstofffilter'] },
      { kind: 'work', tokens: ['zuendkerzen'], words: ['zuendkerzen'] },
      { kind: 'work', tokens: ['kuehlfluessigkeit'], words: ['kuehlfluessigkeit', 'kuehlmittel'] },
      { kind: 'work', tokens: ['getriebeoel'], words: ['getriebeoel'] },
      { kind: 'work', tokens: ['haldex'], words: ['haldex'] },
      { kind: 'work', tokens: ['keilriemen'], words: ['keilriemen'] },
      { kind: 'work', tokens: ['zahnriemen'], words: ['zahnriemen'] },
      { kind: 'work', tokens: ['schluesselbatterie'], words: ['schluesselbatterie'] }
    ];

    items.forEach(function (entry) {
      var rawText = [entry.title, entry.comment].filter(Boolean).join(' ');
      var exp = expandText(rawText);
      var hit = false;

      rules.forEach(function (rule) {
        var rHit = false;
        if (rule.all && hasAllWords(exp, rule.all)) rHit = true;
        if (!rHit && rule.words) {
          for (var w = 0; w < rule.words.length; w++) {
            if (hasPhrase(exp, rule.words[w]) || hasWord(exp, rule.words[w])) { rHit = true; break; }
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

      if (!hit && entry.title) {
        unmatched.push(entry.title + (entry.comment ? (' (' + entry.comment + ')') : ''));
      }
    });

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
    state.huPassed = scraped.huPassed;
    state.carolItems = scraped.items || [];

    if (isNewCar) {
      state.lastScannedSid = state.stockId;
      state.nach = 'nein';
      autoMatchItems(state.carolItems);
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
    var els = Array.from(document.querySelectorAll('button, a, [role="button"]'));
    for (var i = 0; i < els.length; i++) {
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
    if (editPartsBtn && !/nachautorisierung/i.test(document.body.innerText)) {
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

  async function runFullAutomation() {
    if (state.busy) return;
    syncInputsFromDom();

    if (!state.stockId || !state.modell || !state.vin) {
      setHudStatus('err', 'Stock ID, Modell und VIN sind Pflichtfelder.');
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
      bodyWorks: state.bodyWorks,
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
      + '#kva-drawer{position:fixed;top:0;right:0;bottom:0;width:540px;max-width:96vw;z-index:2147483645;'
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
      + '.kva-items{list-style:none;margin:0;padding:0;max-height:125px;overflow-y:auto;}'
      + '.kva-items li{padding:5px 7px;border-bottom:1px dashed #232d40;font-size:12px;}'
      + '.kva-items li:last-child{border-bottom:0;}'
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
    state.modell = getVal('kva-in-mod', state.modell);
    state.vin = getVal('kva-in-vin', state.vin).toUpperCase();
    state.erstzulassung = getVal('kva-in-ez', state.erstzulassung);
    state.getriebe = getVal('kva-in-getr', state.getriebe);
    state.km = getVal('kva-in-km', state.km);
    state.oilRow = parseInt(getVal('kva-in-oil', state.oilRow), 10) || 0;
    state.liters = getVal('kva-in-lit', state.liters);
    for (var i = 0; i < 3; i++) {
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

    var carolListHtml = '';
    if (state.carolItems.length) {
      carolListHtml = '<ul class="kva-items">' + state.carolItems.map(function (it) {
        return '<li><b>' + esc(it.title) + '</b>' + (it.comment ? (' <span style="color:#93a1b5;">— ' + esc(it.comment) + '</span>') : '') + '</li>';
      }).join('') + '</ul>';
    } else {
      carolListHtml = '<div style="color:#93a1b5;font-size:12px;">Noch keine Positionen erkannt — klicke oben auf "Neu einlesen".</div>';
    }

    var workHtml = (meta.workItems || []).map(function (it) {
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

    var oilOpts = '<option value="0">— Kein Öl ausgewählt —</option>' + (meta.oils || []).map(function (o) {
      var sel = Number(state.oilRow) === Number(o.row) ? ' selected' : '';
      return '<option value="' + o.row + '"' + sel + '>' + esc((o.nr ? o.nr + ' · ' : '') + o.label) + '</option>';
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
      + '<h4><span>Fahrzeugdaten (aus Carol)</span>'
      + (state.huPassed ? '<span class="kva-pill">HU Bestanden ✓</span>' : '')
      + '</h4>'
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
      + '<h4><span>Mechanik & Zusatzarbeiten</span></h4>'
      + '<div class="kva-chks">' + dHtml + '</div>'
      + '</div>'
      + '<div class="kva-sec">'
      + '<h4><span>Wartung & Service</span><button type="button" class="kva-ibtn" id="kva-btn-srvpkg" style="padding:2px 8px;font-size:11px;">+ Service Paket</button></h4>'
      + '<div class="kva-chks">' + workHtml + '</div>'
      + '</div>'
      + '<div class="kva-sec">'
      + '<h4><span>Öl Nr. & Nachuntersuchung</span></h4>'
      + '<div class="kva-grid2" style="margin-bottom:8px;">'
      + '<div class="kva-fld"><label>Öl (für G9:G12)</label><select id="kva-in-oil">' + oilOpts + '</select></div>'
      + '<div class="kva-fld"><label>Füllmenge (Liter)</label><input id="kva-in-lit" value="' + esc(state.liters) + '" placeholder="optional"></div>'
      + '</div>'
      + '<div style="display:flex;gap:12px;align-items:center;font-size:12px;">'
      + '<span>GTÜ Nachuntersuchung:</span>'
      + '<label class="kva-chk"><input type="radio" name="kva-nach" value="nein"' + (state.nach === 'nein' ? ' checked' : '') + '><span>Nein</span></label>'
      + '<label class="kva-chk"><input type="radio" name="kva-nach" value="ja"' + (state.nach === 'ja' ? ' checked' : '') + '><span>Ja</span></label>'
      + '<label class="kva-chk"><input type="radio" name="kva-nach" value=""' + (state.nach === '' ? ' checked' : '') + '><span>Aus</span></label>'
      + '</div>'
      + '</div>'
      + '<div class="kva-sec">'
      + '<h4><span>Sonstige Mechanik / Freitext (D24:D26)</span></h4>'
      + '<div class="kva-fld" style="display:flex;flex-direction:column;gap:5px;">'
      + '<input id="kva-fm-0" placeholder="Freitext Zeile 1" value="' + esc(state.freeMech[0] || '') + '">'
      + '<input id="kva-fm-1" placeholder="Freitext Zeile 2" value="' + esc(state.freeMech[1] || '') + '">'
      + '<input id="kva-fm-2" placeholder="Freitext Zeile 3" value="' + esc(state.freeMech[2] || '') + '">'
      + '</div>'
      + '</div>'
      + '<div class="kva-sec">'
      + '<h4><span>Automatisierungs-Schritte</span></h4>'
      + '<label class="kva-chk"><input type="checkbox" id="kva-opt-upload"' + (state.autoUploadPdf ? ' checked' : '') + '><span>PDF (<b id="kva-fn-lbl">' + esc((state.stockId || 'STOCK') + '.pdf') + '</b>) automatisch in Carol unter Dokumente hochladen</span></label>'
      + '<label class="kva-chk"><input type="checkbox" id="kva-opt-parts"' + (state.autoOrderParts ? ' checked' : '') + '><span>Ersatzteile in Nachautorisierung automatisch auf "Bestellt" setzen & abschließen</span></label>'
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
      setHudStatus('ok', 'Fahrzeug- und Service-Daten aus Carol neu eingelesen.');
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

    var dr = document.getElementById('kva-drawer');
    dr.querySelectorAll('input[data-kind]').forEach(function (cb) {
      cb.addEventListener('change', function () {
        var kind = cb.getAttribute('data-kind');
        var row = Number(cb.getAttribute('data-row'));
        var map = kind === 'work' ? state.workSel : state.dSel;
        if (cb.checked) map[row] = true;
        else delete map[row];
        cb.parentElement.classList.toggle('on', cb.checked);
      });
    });

    dr.querySelectorAll('input[name="kva-nach"]').forEach(function (rb) {
      rb.addEventListener('change', function () {
        if (rb.checked) state.nach = rb.value;
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
      if (pendingSid && /nachautorisierung/i.test(document.body ? document.body.innerText : '')) {
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
