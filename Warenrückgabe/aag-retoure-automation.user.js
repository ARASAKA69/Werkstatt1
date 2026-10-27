// ==UserScript==
// @name         AAG Retoure Bot
// @namespace    http://tampermonkey.net/
// @version      1.11
// @description  Sucht AAG Retouren, nimmt nur die neueste Lieferung und schreibt sie ins Blatt KNOLL / AAG
// @author       ARASAKA
// @match        https://portal.allianceautomotive.de/*
// @run-at       document-start
// @updateURL    https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/Warenr%C3%BCckgabe/aag-retoure-automation.user.js
// @downloadURL  https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/Warenr%C3%BCckgabe/aag-retoure-automation.user.js
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      myproduct.guru
// @connect      ean-suche.net
// @connect      api.barcodelookup.com
// @connect      script.google.com
// @connect      script.googleusercontent.com
// ==/UserScript==

(function () {
    'use strict';

    var BOT_VERSION = '1.11';
    var WORKER_KEY = 'aag_retoure_worker';
    var SHEET_URL_KEY = 'aag_retoure_sheet_url';
    var GRUND = 'Falscher Artikel';
    var HUD_POS_KEY = 'aag_retoure_hud_position';
    var QUEUE_KEY = 'aag_retoure_queue';
    var PENDING_KEY = 'aag_retoure_pending';
    var FINISH_KEY = 'aag_retoure_finish';
    var ADDED_KEY = 'aag_retoure_added';
    var EAN_MAP_KEY = 'aag_retoure_ean_map';
    var BARCODE_KEY = 'aag_retoure_barcode_key';
    var TRANSLATE_KEY = 'aag_retoure_translate';
    var PATHS_KEY = 'aag_retoure_api_paths';
    var MIN_LEN = 3;

    var bearer = '';
    var busy = false;
    var leaving = false;
    var abortMission = false;
    var scanTimer = null;
    var scanBurst = 0;
    var lastKeyAt = 0;
    var eanMap = loadEanMap();
    var eanMiss = {};
    var lookupEanNow = '';
    var translateBusy = false;
    var apiPaths = null;
    var pageWin = (typeof unsafeWindow !== 'undefined' && unsafeWindow) ? unsafeWindow : window;
    var pendingPairEan = '';
    var askWait = null;
    var lastEan = '';
    var sheetBusyToken = '';
    var lastPath = location.pathname;
    var grundAt = 0;
    var grundTries = 0;
    var cartHold = false;

    patchTraffic();
    hookHistory();

    function loadEanMap() {
        var map = {};
        try {
            var saved = JSON.parse(gmGet(EAN_MAP_KEY, '{}') || '{}');
            if (saved && typeof saved === 'object') map = saved;
        } catch (e) {}
        try {
            var data = JSON.parse(sessionStorage.getItem(EAN_MAP_KEY) || '{}');
            if (data && typeof data === 'object') {
                Object.keys(data).forEach(function (key) { map[key] = data[key]; });
            }
        } catch (e2) {}
        return map;
    }

    function rememberEan(ean, article) {
        var key = onlyDigits(ean);
        var art = usableArticle(article, key);
        if (!key || !art) return;
        eanMap[key] = art;
        var keys = Object.keys(eanMap);
        if (keys.length > 300) delete eanMap[keys[0]];
        try { sessionStorage.setItem(EAN_MAP_KEY, JSON.stringify(eanMap)); } catch (e) {}
        gmSet(EAN_MAP_KEY, JSON.stringify(eanMap));
    }

    function gmGet(key, fallback) {
        try {
            if (typeof GM_getValue === 'function') {
                var stored = GM_getValue(key, fallback);
                if (stored != null && stored !== '') return stored;
            }
        } catch (e) {}
        try {
            var local = localStorage.getItem(key);
            return local == null || local === '' ? fallback : local;
        } catch (e2) {
            return fallback;
        }
    }

    function gmSet(key, value) {
        try { if (typeof GM_setValue === 'function') GM_setValue(key, value); } catch (e) {}
        try { localStorage.setItem(key, String(value)); } catch (e2) {}
    }

    function workerName() {
        return normSpace(gmGet(WORKER_KEY, ''));
    }

    function saveWorkerName(value) {
        var name = normSpace(value).slice(0, 40);
        gmSet(WORKER_KEY, name);
        return name;
    }

    function sheetUrl() {
        var saved = String(gmGet(SHEET_URL_KEY, '') || '').trim();
        return /^https:\/\/script\.google\.com\/.+\/exec\/?(\?.*)?$/i.test(saved) ? saved : '';
    }

    function saveSheetUrl(value) {
        var url = String(value || '').trim();
        if (!url) {
            gmSet(SHEET_URL_KEY, '');
            return true;
        }
        if (!/^https:\/\/script\.google\.com\/.+\/exec\/?(\?.*)?$/i.test(url)) return false;
        gmSet(SHEET_URL_KEY, url);
        return true;
    }

    function paintSheet() {
        var box = document.getElementById('aag-retoure-sheet-box');
        var state = document.getElementById('aag-retoure-sheet-state');
        var saved = sheetUrl();
        if (box) box.hidden = !!saved;
        if (state) {
            state.textContent = saved ? 'verbunden' : 'Link fehlt';
            state.className = 'aag-sheet-state' + (saved ? ' is-ok' : ' is-bad');
        }
    }

    function cleanStockId(value) {
        var raw = String(value || '').toUpperCase().replace(/\s+/g, '');
        if (!raw) return '';
        var match = raw.match(/[A-Z]{2,3}\d{4,8}/);
        if (!match) return null;
        return match[0];
    }

    function loadQueue() {
        try {
            var data = JSON.parse(sessionStorage.getItem(QUEUE_KEY) || '[]');
            return Array.isArray(data) ? data : [];
        } catch (e) {
            return [];
        }
    }

    function saveQueue(queue) {
        try { sessionStorage.setItem(QUEUE_KEY, JSON.stringify(queue.slice(0, 30))); } catch (e) {}
    }

    function loadAdded() {
        try {
            var data = JSON.parse(sessionStorage.getItem(ADDED_KEY) || '[]');
            return Array.isArray(data) ? data : [];
        } catch (e) {
            return [];
        }
    }

    function rememberAdded(item) {
        var list = loadAdded();
        list.push(item);
        try { sessionStorage.setItem(ADDED_KEY, JSON.stringify(list.slice(-30))); } catch (e) {}
        renderAdded();
    }

    function sleep(ms) {
        return new Promise(function (resolve) { setTimeout(resolve, ms); });
    }

    async function waitFor(fn, ms) {
        var start = Date.now();
        while (Date.now() - start < ms) {
            if (abortMission) return null;
            try {
                var value = fn();
                if (value) return value;
            } catch (e) {}
            await sleep(180);
        }
        return null;
    }

    function normSpace(value) {
        return String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
    }

    function fold(value) {
        return normSpace(value).toLowerCase()
            .replace(/ä/g, 'ae')
            .replace(/ö/g, 'oe')
            .replace(/ü/g, 'ue')
            .replace(/ß/g, 'ss');
    }

    function normArt(value) {
        return normSpace(value).toUpperCase().replace(/[^A-Z0-9]/g, '');
    }

    function onlyDigits(value) {
        return String(value || '').replace(/\D/g, '');
    }

    function cleanScan(raw) {
        return String(raw || '')
            .replace(/[\u0000-\u001F]/g, '')
            .replace(/^\][A-Za-z0-9]{2}/, '')
            .replace(/\s+/g, '')
            .trim();
    }

    function looksLikeEan(value) {
        return /^\d{12,14}$/.test(cleanScan(value));
    }

    function usableArticle(article, ean) {
        var value = normArt(article);
        if (!value || value.length > 40) return '';
        if (value === normArt(ean)) return '';
        if (/^\d{12,14}$/.test(value)) return '';
        return normSpace(article);
    }

    function esc(value) {
        return String(value == null ? '' : value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function onRetours() {
        return /\/ordering\/retours(?:\/|$)/.test(location.pathname || '');
    }

    function pageKind() {
        var path = location.pathname || '';
        if (/\/cart\/?$/.test(path)) return 'cart';
        if (onRetours()) return 'search';
        return 'other';
    }

    function isVisible(el) {
        if (!el || !el.isConnected) return false;
        var style = window.getComputedStyle(el);
        if (style.display === 'none' || style.visibility === 'hidden') return false;
        return el.getClientRects().length > 0;
    }

    function classText(el) {
        if (!el) return '';
        if (typeof el.className === 'string') return el.className;
        return el.getAttribute('class') || '';
    }

    function isDisabled(el) {
        if (!el) return true;
        if (el.disabled) return true;
        if (el.getAttribute('aria-disabled') === 'true') return true;
        return /disabled/i.test(classText(el));
    }

    function findButton(re, reject) {
        var nodes = document.querySelectorAll('button, a, [role="button"], [role="tab"]');
        var best = null;
        var bestLen = 999;
        for (var i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (el.closest('#aag-retoure-hud')) continue;
            if (!isVisible(el)) continue;
            var text = fold(el.innerText || el.getAttribute('aria-label') || '');
            if (!text || text.length > 48) continue;
            if (reject && reject.test(text)) continue;
            if (!re.test(text)) continue;
            if (text.length < bestLen) {
                best = el;
                bestLen = text.length;
            }
        }
        return best;
    }

    function findByText(root, re, maxLen) {
        var scope = root && root.querySelectorAll ? root : document;
        var nodes = scope.querySelectorAll('button, a, span, div, mat-option, mat-select, select, [role="option"], [role="combobox"]');
        var best = null;
        var bestLen = 999;
        for (var i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (el.closest('#aag-retoure-hud')) continue;
            if (!isVisible(el)) continue;
            var text = fold(el.innerText || el.getAttribute('aria-label') || '');
            if (!text || text.length > maxLen) continue;
            if (!re.test(text)) continue;
            if (text.length < bestLen) {
                best = el;
                bestLen = text.length;
            }
        }
        return best;
    }

    function setNativeValue(input, value) {
        var proto = input instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
        var desc = Object.getOwnPropertyDescriptor(proto, 'value');
        if (desc && desc.set) desc.set.call(input, value);
        else input.value = value;
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
    }

    function formInputs() {
        var suchen = findButton(/suchen$/, /rueckgaben/);
        if (!suchen) return [];
        var root = suchen.parentElement;
        for (var depth = 0; depth < 8 && root; depth++) {
            var inputs = Array.prototype.filter.call(root.querySelectorAll('input'), function (input) {
                if (input.closest('#aag-retoure-hud')) return false;
                if (!isVisible(input) || input.readOnly) return false;
                var type = (input.getAttribute('type') || 'text').toLowerCase();
                return type !== 'hidden' && type !== 'checkbox' && type !== 'radio' && type !== 'submit';
            });
            if (inputs.length) return inputs;
            root = root.parentElement;
        }
        return [];
    }

    function resultRoot() {
        var suchen = findButton(/suchen$/, /rueckgaben/);
        var el = suchen ? suchen.parentElement : document.body;
        for (var i = 0; i < 8 && el && el.parentElement; i++) {
            if ((el.innerText || '').length > 250) return el.parentElement;
            el = el.parentElement;
        }
        return document.body;
    }

    function collectDeliveryCards(root) {
        var nodes = (root || document.body).querySelectorAll('tr, [role="row"], mat-expansion-panel, li, div');
        var hits = [];
        for (var i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (el.closest('#aag-retoure-hud')) continue;
            var text = normSpace(el.innerText || '');
            if (text.length > 180) continue;
            if ((text.match(/LS\d{6,}/gi) || []).length !== 1) continue;
            if (!/\d{2}\.\d{2}\.\d{2}/.test(text)) continue;
            hits.push(el);
        }
        return hits.filter(function (el) {
            return !hits.some(function (other) { return other !== el && el.contains(other); });
        }).sort(function (a, b) { return cardTime(b) - cardTime(a); });
    }

    function deliveryCards() {
        var root = resultRoot();
        var cards = collectDeliveryCards(root);
        if (cards.length || root === document.body) return cards;
        return collectDeliveryCards(document.body);
    }

    function cardTime(el) {
        var match = normSpace(el.innerText || '').match(/(\d{2})\.(\d{2})\.(\d{2})/);
        if (!match) return 0;
        var year = Number(match[3]);
        year = year < 70 ? 2000 + year : 1900 + year;
        return new Date(year, Number(match[2]) - 1, Number(match[1])).getTime();
    }

    function lsOf(el) {
        var match = normSpace(el.innerText || '').match(/LS\d{6,}/i);
        return match ? match[0].toUpperCase() : '';
    }

    function dateOf(el) {
        var match = normSpace(el.innerText || '').match(/\d{2}\.\d{2}\.\d{2}/);
        return match ? match[0] : '';
    }

    function expander(card) {
        var buttons = card.querySelectorAll('button, [role="button"]');
        var candidate = null;
        for (var i = 0; i < buttons.length; i++) {
            var button = buttons[i];
            var text = fold(button.innerText || '');
            if (/suchen|zuruecksetzen|zurueckgeben|warenkorb|drucken|absenden/.test(text)) continue;
            if (button.getAttribute('aria-expanded') === 'false') return button;
            candidate = button;
        }
        return candidate || card;
    }

    function lineOf(button, article) {
        var want = normArt(article);
        var el = button.parentElement;
        for (var i = 0; i < 14 && el; i++) {
            var raw = el.innerText || '';
            if (raw.length < 1200 && normArt(raw).indexOf(want) !== -1 && lineHasArticle(el, article)) return el;
            el = el.parentElement;
        }
        return null;
    }

    function lineHasArticle(line, article) {
        var want = normArt(article);
        if (!want || !line) return false;
        var cells = line.querySelectorAll('td, [role="cell"], mat-cell');
        if (cells.length) {
            for (var i = 0; i < Math.min(cells.length, 3); i++) {
                if (normArt(cells[i].innerText) === want) return true;
            }
        }
        return normArt(line.innerText).indexOf(want) !== -1;
    }

    function qtyIndex(line) {
        var table = line.closest('table, [role="table"], mat-table');
        if (!table) return 3;
        var headers = table.querySelectorAll('th, [role="columnheader"], mat-header-cell');
        for (var i = 0; i < headers.length; i++) {
            if (fold(headers[i].innerText).indexOf('rueckgabefaehige') !== -1) return i;
        }
        return 3;
    }

    function returnableQty(line) {
        if (!line) return null;
        var cells = line.querySelectorAll('td, [role="cell"], mat-cell');
        var index = qtyIndex(line);
        if (cells.length <= index) return null;
        var raw = normSpace(cells[index].innerText).replace(/\./g, '').replace(',', '.');
        var qty = parseFloat(raw);
        return isNaN(qty) ? null : qty;
    }

    function lineName(line, article) {
        if (!line) return '';
        var cells = line.querySelectorAll('td, [role="cell"], mat-cell');
        if (cells.length >= 2) return normSpace(cells[1].innerText).slice(0, 80);
        return normSpace(line.innerText).replace(new RegExp(article, 'ig'), '').replace(/zurückgeben/ig, '').slice(0, 80);
    }

    function looksFilled(button) {
        var style = window.getComputedStyle(button);
        var match = String(style.backgroundColor || '').match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?/);
        if (!match) return false;
        if (match[4] != null && Number(match[4]) < 0.25) return false;
        return (Number(match[1]) + Number(match[2]) + Number(match[3])) / 3 < 90;
    }

    function isAddButton(button) {
        if (!button || !isVisible(button) || isDisabled(button)) return false;
        var raw = (button.innerText || '') + ' ' + (button.getAttribute('aria-label') || '');
        if (buttonText(button).indexOf('zurueckgeben') === -1) return false;
        if (/[−–]|remove/i.test(raw) && !/[+]/.test(raw)) return false;
        if (looksFilled(button) && !/[+]/.test(raw)) return false;
        return true;
    }

    function buttonText(el) {
        return fold((el.innerText || '') + ' ' + (el.getAttribute('aria-label') || '') + ' ' + (el.getAttribute('title') || ''));
    }

    function isHeaderNode(el) {
        if (!el || (el.closest && el.closest('#aag-retoure-hud'))) return false;
        var text = normSpace(el.innerText || '');
        if (!text || text.length > 180) return false;
        if ((text.match(/LS\d{6,}/gi) || []).length !== 1) return false;
        return /\d{2}\.\d{2}\.\d{2}/.test(text);
    }

    function branchOf(header) {
        var headers = deliveryCards();
        var el = header;
        var best = header;
        while (el && el.parentElement && el.parentElement !== document.body) {
            var parent = el.parentElement;
            if (parent.closest && parent.closest('#aag-retoure-hud')) break;
            var count = 0;
            for (var i = 0; i < headers.length; i++) {
                if (parent.contains(headers[i])) count++;
                if (count > 1) break;
            }
            if (count !== 1) break;
            best = parent;
            el = parent;
        }
        return best;
    }

    function headerOwns(header, node) {
        if (!header || !node) return false;
        var branch = branchOf(header);
        if (branch.contains(node)) return true;
        var headers = deliveryCards();
        var el = branch;
        var steps = 0;
        while (el && steps < 8) {
            var next = el.nextElementSibling;
            if (!next) break;
            if (isHeaderNode(next)) break;
            var foreign = false;
            for (var i = 0; i < headers.length; i++) {
                if (headers[i] !== header && next.contains(headers[i])) {
                    foreign = true;
                    break;
                }
            }
            if (foreign) break;
            if (next.contains(node)) return true;
            el = next;
            steps++;
        }
        return false;
    }

    function headerByLs(ls) {
        var headers = deliveryCards();
        for (var i = 0; i < headers.length; i++) {
            if (lsOf(headers[i]) === ls) return headers[i];
        }
        return null;
    }

    function textHasArticle(text, article) {
        var want = normArt(article);
        if (!want) return false;
        var parts = String(text || '').split(/[\s,;|/]+/);
        var i;
        for (i = 0; i < parts.length; i++) {
            if (normArt(parts[i]) === want) return true;
        }
        return false;
    }

    function articleRows(article) {
        var want = normArt(article);
        var nodes = document.querySelectorAll('td, [role="cell"], mat-cell, div, span, a');
        var rows = [];
        var i;
        for (i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (el.closest && el.closest('#aag-retoure-hud')) continue;
            var own = normSpace(el.innerText || '');
            if (!own || own.length > 40) continue;
            if (normArt(own) !== want && !textHasArticle(own, article)) continue;
            var row = el;
            var p = el;
            var up;
            for (up = 0; up < 8 && p; up++) {
                var tag = (p.tagName || '').toLowerCase();
                var role = p.getAttribute ? (p.getAttribute('role') || '') : '';
                var cls = String(p.className || '');
                if (tag === 'tr' || role === 'row' || /mat-row|mdc-row/.test(cls)) {
                    row = p;
                    break;
                }
                var size = (p.innerText || '').length;
                if (size > 40 && size < 500 && textHasArticle(p.innerText, article)) row = p;
                p = p.parentElement;
            }
            if (rows.indexOf(row) === -1) rows.push(row);
        }
        return rows;
    }

    function returnButtonOnRow(row) {
        if (!row) return null;
        var inside = row.querySelectorAll('button, a, [role="button"]');
        var i;
        for (i = 0; i < inside.length; i++) {
            if (inside[i].closest && inside[i].closest('#aag-retoure-hud')) continue;
            if (buttonText(inside[i]).indexOf('zurueckgeben') !== -1) return inside[i];
        }
        var rect = row.getBoundingClientRect();
        if (!rect.height) return null;
        var limit = Math.max(28, rect.height * 0.65);
        var mid = (rect.top + rect.bottom) / 2;
        var all = document.querySelectorAll('button, a, [role="button"]');
        var best = null;
        var bestDy = limit;
        for (i = 0; i < all.length; i++) {
            var button = all[i];
            if (button.closest && button.closest('#aag-retoure-hud')) continue;
            if (buttonText(button).indexOf('zurueckgeben') === -1) continue;
            var br = button.getBoundingClientRect();
            if (!br.height) continue;
            var dy = Math.abs((br.top + br.bottom) / 2 - mid);
            if (dy < bestDy) {
                bestDy = dy;
                best = button;
            }
        }
        return best;
    }

    function returnButtonsNear(card, article) {
        var header = headerByLs(lsOf(card)) || card;
        var rows = articleRows(article);
        var owned = [];
        var i;
        for (i = 0; i < rows.length; i++) {
            if (headerOwns(header, rows[i])) owned.push(rows[i]);
        }
        var use = owned.length ? owned : rows;
        var seen = [];
        for (i = 0; i < use.length; i++) {
            var button = returnButtonOnRow(use[i]);
            if (button && seen.indexOf(button) === -1) seen.push(button);
        }
        return seen;
    }

    function expandHeader(header) {
        var root = branchOf(header);
        var buttons = root.querySelectorAll('button, [role="button"]');
        for (var i = 0; i < buttons.length; i++) {
            var button = buttons[i];
            if (buttonText(button).indexOf('zurueckgeben') !== -1) continue;
            if (button.getAttribute('aria-expanded') === 'true') return;
            if (button.getAttribute('aria-expanded') === 'false') {
                forceClick(button);
                return;
            }
        }
        forceClick(header);
    }

    function forceClick(el) {
        if (!el) return;
        try { el.scrollIntoView({ block: 'center' }); } catch (e) {}
        el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: pageWin }));
        el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: pageWin }));
        try { el.click(); } catch (e2) {}
    }

    function cartBadge() {
        var tab = findButton(/warenkorb/);
        if (!tab) return 0;
        var root = tab.parentElement || tab;
        var nodes = root.querySelectorAll('span, sup, div, em');
        for (var i = 0; i < nodes.length; i++) {
            var text = normSpace(nodes[i].innerText || '');
            if (/^\d{1,3}$/.test(text)) return Number(text);
        }
        return 0;
    }

    function outsideHudText() {
        var text = document.body ? (document.body.innerText || '') : '';
        var hud = document.getElementById('aag-retoure-hud');
        var hudText = hud ? (hud.innerText || '') : '';
        if (hudText && text.indexOf(hudText) !== -1) text = text.replace(hudText, ' ');
        return text;
    }

    function cartVisible() {
        if (!onRetours() || !document.body) return false;
        if (pageKind() === 'cart') return true;
        if (findByText(document, /bitte grund/, 40)) return true;
        var send = findButton(/^absenden$/);
        return !!(send && /LS\d{6,}/i.test(outsideHudText()));
    }

    function cartHasItems() {
        if (!cartVisible()) return false;
        if (/keine artikel/i.test(outsideHudText())) return false;
        if (/LS\d{6,}/i.test(outsideHudText())) return true;
        return !!findByText(document, /bitte grund/, 40);
    }

    function goSearch() {
        var tab = findButton(/rueckgaben suchen/);
        if (tab) tab.click();
        else location.assign('/ordering/retours/search');
    }

    function goCart() {
        var tab = findButton(/warenkorb/);
        if (tab) tab.click();
        else location.assign('/ordering/retours/cart');
    }

    function captureBearer(value) {
        var match = String(value || '').match(/^Bearer\s+(.+)$/i);
        if (match && match[1].length > 20) bearer = match[1];
    }

    function articleField(obj) {
        var keys = ['articleNumber', 'articleNo', 'articleNr', 'manufacturerArticleNumber', 'manufacturerNumber', 'partNumber', 'itemNumber', 'materialNumber', 'tecdocArticleNumber', 'wholesalerArticleNumber', 'supplierArticleNumber'];
        for (var i = 0; i < keys.length; i++) {
            if (obj[keys[i]]) return String(obj[keys[i]]);
        }
        return '';
    }

    function eanField(obj) {
        var keys = ['ean', 'EAN', 'gtin', 'GTIN', 'barcode', 'barCode', 'eanCode', 'gtinCode', 'europeanArticleNumber'];
        for (var i = 0; i < keys.length; i++) {
            if (obj[keys[i]]) return String(obj[keys[i]]);
        }
        return '';
    }

    function harvestText(text) {
        if (!text || text.length > 400000) return;
        var trimmed = String(text).trim();
        if (trimmed.charAt(0) !== '{' && trimmed.charAt(0) !== '[') return;
        var data;
        try { data = JSON.parse(trimmed); } catch (e) { return; }
        harvestNode(data, 0);
        if (lookupEanNow) harvestLookup(data, lookupEanNow);
    }

    function preferredArticle(obj, ean) {
        var keys = ['articleNumber', 'articleNo', 'articleNr', 'wholesalerArticleNumber', 'supplierArticleNumber', 'itemNumber', 'partNumber', 'manufacturerArticleNumber', 'tecdocArticleNumber'];
        for (var i = 0; i < keys.length; i++) {
            if (!obj[keys[i]]) continue;
            var art = usableArticle(obj[keys[i]], ean);
            if (art) return art;
        }
        return '';
    }

    function objectMentionsEan(obj, ean) {
        var keys = Object.keys(obj);
        for (var i = 0; i < keys.length; i++) {
            var value = obj[keys[i]];
            if ((typeof value === 'string' || typeof value === 'number') && String(value).indexOf(ean) !== -1) return true;
        }
        return false;
    }

    function harvestLookup(node, ean, depth, found) {
        if (depth == null) depth = 0;
        if (!found) found = [];
        if (!node || depth > 7) {
            if (depth === 0 && found.length === 1) rememberEan(ean, found[0]);
            return;
        }
        if (!found) found = [];
        if (Array.isArray(node)) {
            for (var i = 0; i < node.length && i < 40; i++) harvestLookup(node[i], ean, depth + 1, found);
        } else if (typeof node === 'object' && objectMentionsEan(node, ean)) {
            var art = preferredArticle(node, ean);
            if (art && found.indexOf(art) === -1) found.push(art);
            var keys = Object.keys(node);
            for (var k = 0; k < keys.length && k < 30; k++) harvestLookup(node[keys[k]], ean, depth + 1, found);
        } else if (typeof node === 'object') {
            var childKeys = Object.keys(node);
            for (var c = 0; c < childKeys.length && c < 30; c++) harvestLookup(node[childKeys[c]], ean, depth + 1, found);
        }
        if (depth === 0 && found.length === 1) rememberEan(ean, found[0]);
    }

    function harvestNode(node, depth) {
        if (!node || depth > 7) return;
        if (Array.isArray(node)) {
            for (var i = 0; i < node.length && i < 80; i++) harvestNode(node[i], depth + 1);
            return;
        }
        if (typeof node !== 'object') return;
        var art = articleField(node);
        var ean = onlyDigits(eanField(node));
        if (art && ean.length >= 8) rememberEan(ean, art);
        var keys = Object.keys(node);
        for (var k = 0; k < keys.length && k < 40; k++) harvestNode(node[keys[k]], depth + 1);
    }

    function patchTraffic() {
        var xhrProto = pageWin.XMLHttpRequest.prototype;
        if (!xhrProto.__aagRetourePatched) {
            var originalHeader = xhrProto.setRequestHeader;
            xhrProto.setRequestHeader = function (name, value) {
                if (String(name || '').toLowerCase() === 'authorization') captureBearer(value);
                return originalHeader.apply(this, arguments);
            };
            var originalSend = xhrProto.send;
            xhrProto.send = function () {
                this.addEventListener('load', function () {
                    try {
                        if (!this.responseType || this.responseType === 'text') harvestText(this.responseText);
                    } catch (e) {}
                });
                return originalSend.apply(this, arguments);
            };
            xhrProto.__aagRetourePatched = true;
        }
        if (pageWin.fetch && !pageWin.fetch.__aagRetoureFetch) {
            var originalFetch = pageWin.fetch;
            var wrapped = function () {
                try {
                    var init = arguments[1];
                    var headers = new Headers((init && init.headers) || (arguments[0] && arguments[0].headers) || {});
                    captureBearer(headers.get('authorization'));
                } catch (e) {}
                return originalFetch.apply(this, arguments).then(function (response) {
                    try { response.clone().text().then(harvestText).catch(function () {}); } catch (e) {}
                    return response;
                });
            };
            wrapped.__aagRetoureFetch = true;
            pageWin.fetch = wrapped;
        }
    }

    async function learnApiPaths() {
        if (apiPaths) return apiPaths;
        try {
            var cached = JSON.parse(sessionStorage.getItem(PATHS_KEY) || 'null');
            if (Array.isArray(cached) && cached.length) {
                apiPaths = cached;
                return apiPaths;
            }
        } catch (e) {}
        var srcs = [];
        document.querySelectorAll('script[src]').forEach(function (script) {
            var src = script.src || '';
            if (!/\.js(\?|$)/.test(src)) return;
            if (/polyfill|runtime|zone|vendor/i.test(src)) return;
            srcs.push(src);
        });
        var found = {};
        var budget = Date.now() + 7000;
        for (var i = 0; i < srcs.length && i < 8; i++) {
            if (Date.now() > budget) break;
            try {
                var response = await fetch(srcs[i], { credentials: 'omit' });
                var text = await response.text();
                var matches = text.match(/\/api\/v1\/[A-Za-z0-9_./${}-]{3,90}/g) || [];
                matches.forEach(function (path) { found[path] = true; });
            } catch (e) {}
        }
        apiPaths = Object.keys(found);
        try { sessionStorage.setItem(PATHS_KEY, JSON.stringify(apiPaths)); } catch (e) {}
        return apiPaths;
    }

    function lookupUrls(ean) {
        var paths = (apiPaths || []).filter(function (path) {
            return /article|catalog|product|part|ean|gtin|barcode/i.test(path) && !/receipt|invoice|user|auth|login/i.test(path);
        }).sort(function (a, b) {
            function score(path) {
                var value = 0;
                if (/ean|gtin|barcode/i.test(path)) value += 5;
                if (/search/i.test(path)) value += 3;
                if (/article|product|catalog/i.test(path)) value += 2;
                return value;
            }
            return score(b) - score(a);
        }).slice(0, 3);
        var urls = [];
        paths.forEach(function (path) {
            var base = path.replace(/\$\{[^}]+\}/g, encodeURIComponent(ean)).split('?')[0];
            if (base.indexOf(ean) !== -1) urls.push(base);
            urls.push(base + '?ean=' + encodeURIComponent(ean));
            urls.push(base + '?query=' + encodeURIComponent(ean));
        });
        [
            '/api/v1/customer/articles?ean=',
            '/api/v1/customer/articles?query=',
            '/api/v1/customer/article/search?ean=',
            '/api/v1/customer/article/search?query=',
            '/api/v1/customer/articles/search?search=',
            '/api/v1/catalog/articles?ean=',
            '/api/v1/customer/product/search?query='
        ].forEach(function (path) { urls.push(path + encodeURIComponent(ean)); });
        var unique = [];
        urls.forEach(function (url) {
            if (unique.indexOf(url) === -1) unique.push(url);
        });
        return unique.slice(0, 8);
    }

    function pickArticle(node, ean, depth, found) {
        if (!node || depth > 7) return;
        if (Array.isArray(node)) {
            for (var i = 0; i < node.length; i++) pickArticle(node[i], ean, depth + 1, found);
            return;
        }
        if (typeof node !== 'object') return;
        var art = articleField(node);
        var nodeEan = onlyDigits(eanField(node));
        if (art && nodeEan && nodeEan === ean) {
            var matched = usableArticle(art, ean);
            if (matched) found.matched = matched;
        }
        if (art && found.length < 3) {
            var sole = usableArticle(art, ean);
            if (sole && found.indexOf(sole) === -1) found.push(sole);
        }
        var keys = Object.keys(node);
        for (var k = 0; k < keys.length; k++) pickArticle(node[keys[k]], ean, depth + 1, found);
    }

    async function fetchJson(url) {
        var headers = { Accept: 'application/json, text/plain, */*' };
        if (bearer) headers.Authorization = 'Bearer ' + bearer;
        var controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
        var timer = controller ? setTimeout(function () { controller.abort(); }, 2500) : null;
        try {
            var response = await fetch(url, {
                credentials: 'include',
                headers: headers,
                signal: controller ? controller.signal : undefined
            });
            if (!response.ok) return null;
            var text = await response.text();
            harvestText(text);
            if (!text || text.length > 400000) return null;
            var trimmed = text.trim();
            if (trimmed.charAt(0) !== '{' && trimmed.charAt(0) !== '[') return null;
            return JSON.parse(trimmed);
        } catch (e) {
            return null;
        } finally {
            if (timer) clearTimeout(timer);
        }
    }

    async function lookupEan(ean) {
        if (eanMap[ean]) return eanMap[ean];
        if (eanMiss[ean]) return '';
        await learnApiPaths();
        var urls = lookupUrls(ean);
        for (var i = 0; i < urls.length; i++) {
            if (abortMission) return '';
            var data = await fetchJson(urls[i]);
            if (!data) continue;
            var found = [];
            pickArticle(data, ean, 0, found);
            var art = found.matched || (found.length === 1 ? found[0] : '');
            if (art) {
                rememberEan(ean, art);
                return art;
            }
        }
        eanMiss[ean] = true;
        return '';
    }

    function loadTranslate() {
        try {
            var data = JSON.parse(sessionStorage.getItem(TRANSLATE_KEY) || 'null');
            return data && data.ean ? data : null;
        } catch (e) {
            return null;
        }
    }

    function saveTranslate(job) {
        try { sessionStorage.setItem(TRANSLATE_KEY, JSON.stringify(job)); } catch (e) {}
    }

    function findArtikelsucheLink() {
        var nodes = document.querySelectorAll('a, button, [role="link"], [role="menuitem"]');
        var best = null;
        for (var i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (el.closest('#aag-retoure-hud') || !isVisible(el)) continue;
            var text = fold(el.innerText || el.getAttribute('aria-label') || '');
            if (!text || text.length > 24 || text.indexOf('artikelsuche') === -1) continue;
            if (text === 'artikelsuche') return el;
            best = el;
        }
        return best;
    }

    function openArtikelsuche(ean) {
        var job = loadTranslate() || {};
        if (job.ean === ean && job.clicked) return;
        lookupEanNow = ean;
        job.ean = ean;
        job.phase = 'open';
        job.clicked = true;
        job.at = Date.now();
        saveTranslate(job);
        setStatus('EAN ' + ean + ' geht in die Artikelsuche...', 'active');
        var link = findArtikelsucheLink();
        if (!link) {
            sessionStorage.removeItem(TRANSLATE_KEY);
            lookupEanNow = '';
            setStatus('Artikelsuche in der oberen Leiste nicht gefunden.', 'error');
            return;
        }
        var href = link.href || link.getAttribute('href') || '';
        link.click();
        setTimeout(function () {
            var current = loadTranslate();
            if (!current || current.phase !== 'open' || !onRetours()) return;
            if (href && href.charAt(0) === '/') location.assign(href);
        }, 1200);
    }

    function findCatalogInput() {
        var inputs = document.querySelectorAll('input');
        var fallback = null;
        for (var i = 0; i < inputs.length; i++) {
            var input = inputs[i];
            if (input.closest('#aag-retoure-hud') || !isVisible(input) || input.readOnly) continue;
            var type = (input.getAttribute('type') || 'text').toLowerCase();
            if (type === 'hidden' || type === 'checkbox' || type === 'radio' || type === 'password') continue;
            var hint = fold((input.getAttribute('placeholder') || '') + ' ' + (input.getAttribute('aria-label') || '') + ' ' + (input.id || '') + ' ' + (input.name || ''));
            if (/ean|artikel|suche|search|gtin|barcode|nummer/.test(hint)) return input;
            if (!fallback) fallback = input;
        }
        return fallback;
    }

    function looksLikePart(value, ean) {
        var raw = normSpace(value);
        var compact = raw.replace(/\s/g, '');
        if (compact.length < 4 || compact.length > 22) return false;
        if (onlyDigits(compact) === ean) return false;
        if (/^\d{12,14}$/.test(compact)) return false;
        if (/^\d{1,2}\.\d{1,2}\.\d{2,4}$/.test(raw)) return false;
        if (/€|eur/i.test(raw)) return false;
        if (!/[0-9]/.test(compact)) return false;
        if (!/[A-Za-z]/.test(compact) && !/^\d{5,11}$/.test(compact)) return false;
        return true;
    }

    function labeledArticle(ean) {
        var root = document.querySelector('main') || document.body;
        if (!root) return '';
        var match = (root.innerText || '').match(/Artikel(?:nummer|[\s\-]*Nr\.?)\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9\-\/\.]{2,24})/i);
        if (!match) return '';
        var art = normSpace(match[1]);
        if (onlyDigits(art) === ean || !looksLikePart(art, ean)) return '';
        return art;
    }

    function scrapeChoices(ean) {
        var root = document.querySelector('main') || document.body;
        var nodes = root.querySelectorAll('a, button, td, span, div, h1, h2, h3, h4');
        var out = [];
        for (var i = 0; i < nodes.length && out.length < 8; i++) {
            var el = nodes[i];
            if (el.closest('#aag-retoure-hud') || !isVisible(el) || el.children.length > 2) continue;
            var raw = normSpace(el.innerText || '');
            if (!looksLikePart(raw, ean) || out.indexOf(raw) !== -1) continue;
            out.push(raw);
        }
        return out;
    }

    function renderChoices(list) {
        var host = document.getElementById('aag-retoure-choices');
        if (!host) return;
        var sig = (list || []).join('|');
        if (host.getAttribute('data-sig') === sig) return;
        host.setAttribute('data-sig', sig);
        host.innerHTML = '';
        if (!list || !list.length) return;
        var label = document.createElement('div');
        label.className = 'aag-last-label';
        label.textContent = 'Artikelnummer antippen';
        host.appendChild(label);
        list.forEach(function (art) {
            var btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'aag-choice';
            btn.textContent = art;
            btn.addEventListener('click', function () { acceptTranslated(art); });
            host.appendChild(btn);
        });
    }

    function acceptTranslated(article) {
        var job = loadTranslate();
        var art = normSpace(article);
        if (!art) return;
        if (job && job.ean) rememberEan(job.ean, art);
        lookupEanNow = '';
        sessionStorage.removeItem(TRANSLATE_KEY);
        renderChoices([]);
        setStatus(art + ' aus der Artikelsuche. Zurück zur Retoure...', 'done');
        if (onRetours()) {
            addArticle(art);
            return;
        }
        sessionStorage.setItem(PENDING_KEY, art);
        location.assign('/ordering/retours/search');
    }

    function catalogMiss() {
        var nodes = document.querySelectorAll('[class*="toast"], [class*="snack"], [role="alert"]');
        for (var i = 0; i < nodes.length; i++) {
            if (!isVisible(nodes[i])) continue;
            if (fold(nodes[i].innerText || '').indexOf('kein artikel') !== -1) return true;
        }
        return false;
    }

    async function fillCatalog(job) {
        if (translateBusy) return;
        translateBusy = true;
        try {
            lookupEanNow = job.ean;
            var input = await waitFor(findCatalogInput, 8000);
            if (abortMission || !loadTranslate()) return;
            if (!input) {
                setStatus('Suchfeld in der Artikelsuche nicht gefunden. Nummer hier eintippen und Enter.', 'error');
                job.phase = 'pick';
                job.filledAt = Date.now();
                saveTranslate(job);
                return;
            }
            input.focus();
            setNativeValue(input, job.ean);
            await sleep(120);
            input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
            input.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
            var suchen = findButton(/suchen$/, /rueckgaben|retour/);
            if (suchen) suchen.click();
            job.phase = 'pick';
            job.filledAt = Date.now();
            saveTranslate(job);
            setStatus('Artikelsuche läuft für ' + job.ean + '...', 'active');
        } finally {
            translateBusy = false;
        }
    }

    function translateTick() {
        var job = loadTranslate();
        if (!job) return;
        ensureHud();
        if (eanMap[job.ean]) {
            acceptTranslated(eanMap[job.ean]);
            return;
        }
        if (job.phase === 'open' && !onRetours()) {
            job.phase = 'fill';
            saveTranslate(job);
        }
        if (job.phase === 'fill' && !translateBusy) fillCatalog(job);
        if (job.phase !== 'pick') return;
        var labeled = labeledArticle(job.ean);
        if (labeled) {
            acceptTranslated(labeled);
            return;
        }
        if (job.filledAt && Date.now() - job.filledAt > 900) renderChoices(scrapeChoices(job.ean));
        if (job.filledAt && Date.now() - job.filledAt > 2500 && catalogMiss()) setStatus('Artikelsuche kennt diese EAN nicht. Die Artikelnummer hier eintippen und Enter.', 'error');
        else setStatus('EAN ' + job.ean + ' liegt in der Artikelsuche. Artikelnummer antippen, oder hier eintippen.', 'active');
    }

    function partCandidates_(text, ean) {
        var plain = normSpace(String(text || '').replace(/<[^>]+>/g, ' '));
        var out = [];
        var seen = {};
        function add(raw) {
            var compact = normArt(raw);
            if (!compact || seen[compact] || compact.length > 18) return;
            if (compact.replace(/\D/g, '').length < 6) return;
            if (!usableArticle(compact, ean)) return;
            seen[compact] = true;
            out.push(compact);
        }
        var re = /\b((?=[A-Z0-9]{0,3}\d)[A-Z0-9]{2,4}(?:[\s.\-]+(?=[A-Z0-9]{0,3}\d)[A-Z0-9]{2,4}){1,5})\b/gi;
        var hit;
        while ((hit = re.exec(plain))) {
            var groups = hit[1].split(/[\s.\-]+/);
            var start = 0;
            if (groups.length >= 3 && /[A-Z]/i.test(groups[0])) {
                var rest = groups.slice(1).join('');
                if (/^\d{8,12}$/.test(rest)) start = 1;
            }
            add(groups.slice(start).join(''));
        }
        var word = /\b(?=[A-Z0-9]*\d)[A-Z0-9]{6,16}\b/gi;
        while ((hit = word.exec(plain))) add(hit[0]);
        return out;
    }

    function preferPart_(list) {
        var best = '';
        var bestScore = -1;
        var i;
        for (i = 0; i < list.length; i++) {
            var token = list[i];
            var digits = token.replace(/\D/g, '').length;
            var letters = token.replace(/[0-9]/g, '').length;
            var score = digits * 10 - letters * 3;
            if (/^\d{7,12}$/.test(token)) score += 40;
            if (score > bestScore) {
                bestScore = score;
                best = token;
            }
        }
        return best;
    }

    function partFromLookupText(text, ean) {
        return preferPart_(partCandidates_(text, ean));
    }

    function offerTexts_(html) {
        var block = String(html || '').match(/<table[^>]*class="[^"]*offers-table[^"]*"[\s\S]*?<\/table>/i);
        if (!block) return '';
        var texts = [];
        var re = /<td\b[^>]*>([\s\S]*?)<\/td>/gi;
        var hit;
        while ((hit = re.exec(block[0]))) texts.push(hit[1].replace(/<[^>]+>/g, ' '));
        return texts.join('\n');
    }

    function shortProduct_(name) {
        var text = normSpace(name);
        if (!text) return '';
        if (text.length > 90) text = text.slice(0, 87) + '...';
        return text + '. ';
    }

    function gmFetch(url, accept) {
        return new Promise(function (resolve) {
            if (typeof GM_xmlhttpRequest !== 'function') {
                resolve({ status: 0, text: '' });
                return;
            }
            GM_xmlhttpRequest({
                method: 'GET',
                url: url,
                timeout: 12000,
                headers: { 'Accept': accept || 'application/json, text/html;q=0.9', 'Accept-Language': 'de-DE,de;q=0.9' },
                onload: function (res) {
                    resolve({ status: res && res.status || 0, text: res && res.responseText || '' });
                },
                onerror: function () { resolve({ status: 0, text: '' }); },
                ontimeout: function () { resolve({ status: 0, text: '' }); }
            });
        });
    }

    function readJson_(text) {
        try { return JSON.parse(text); } catch (e) { return null; }
    }

    async function lookupBarcodeApi(ean) {
        var key = String(gmGet(BARCODE_KEY, '') || '').trim();
        if (!key) return null;
        var res = await gmFetch('https://api.barcodelookup.com/v3/products?barcode=' + encodeURIComponent(ean) + '&key=' + encodeURIComponent(key));
        var data = readJson_(res.text);
        var item = data && data.products && data.products[0];
        if (!item) return { article: '', name: '' };
        var name = normSpace([item.manufacturer || item.brand || '', item.title || item.product_name || ''].join(' '));
        var article = partFromLookupText(item.mpn || '', ean) || partFromLookupText(name, ean);
        return { article: article, name: name };
    }

    async function lookupProductGuru(ean) {
        var res = await gmFetch('https://myproduct.guru/lookup/' + encodeURIComponent(ean) + '.json');
        if (res.status === 429 || res.status === 504) return { article: '', name: '', limited: true };
        var data = readJson_(res.text);
        if (!data || !data.name) return { article: '', name: '' };
        var name = normSpace(data.name);
        var article = partFromLookupText(name, ean);
        return { article: article, name: name };
    }

    async function lookupEanSucheNet(ean) {
        var res = await gmFetch('https://ean-suche.net/api/produkt/' + encodeURIComponent(ean));
        if (res.status === 429) return { article: '', name: '', limited: true };
        var data = readJson_(res.text);
        if (!data || data.found === false) return { article: '', name: '' };
        var name = normSpace(data.product_name || '');
        var article = partFromLookupText(name, ean);
        if (article) return { article: article, name: name };
        var page = await gmFetch('https://ean-suche.net/produkt/' + encodeURIComponent(ean), 'text/html');
        var offers = offerTexts_(page.text);
        article = partFromLookupText(offers, ean);
        return { article: article, name: name };
    }

    async function translateEan(ean) {
        if (eanMap[ean]) return { article: eanMap[ean], name: '' };
        var best = { article: '', name: '', limited: false };
        var api = await lookupBarcodeApi(ean);
        if (api && api.article) return api;
        if (api && api.name) best.name = api.name;
        var guru = await lookupProductGuru(ean);
        if (guru.limited) best.limited = true;
        if (guru.article) return { article: guru.article, name: guru.name || best.name };
        if (guru.name) best.name = guru.name;
        var net = await lookupEanSucheNet(ean);
        if (net.limited) best.limited = true;
        if (net.article) return { article: net.article, name: net.name || best.name };
        if (net.name) best.name = net.name;
        return best;
    }

    function askStock(found, article) {
        return new Promise(function (resolve) {
            var box = document.getElementById('aag-retoure-ask');
            var stock = document.getElementById('aag-retoure-stock');
            var who = document.getElementById('aag-retoure-who');
            var take = document.getElementById('aag-retoure-take');
            if (!box || !stock || !who || !take) {
                resolve(null);
                return;
            }
            var done = false;
            box.hidden = false;
            take.hidden = false;
            if (!who.value) who.value = workerName();
            setStatus(article + ' in ' + (found.ls || 'der neuesten Lieferung') + '. Stock-ID leer lassen oder eintragen, Enter gibt zurück.', 'active');
            stock.focus();
            function close() {
                done = true;
                askWait = null;
                box.hidden = true;
                take.hidden = true;
                stock.removeEventListener('keydown', onKey);
                take.removeEventListener('click', onTake);
            }
            function finish(ok) {
                if (done) return;
                if (!ok) {
                    close();
                    resolve(null);
                    return;
                }
                if (!normSpace(who.value)) {
                    setStatus('Name fehlt. Zum Beispiel Semmel.', 'error');
                    who.focus();
                    return;
                }
                var stockId = cleanStockId(stock.value);
                if (stockId === null) {
                    setStatus('Stock-ID leer lassen oder so schreiben: JR97017.', 'error');
                    stock.focus();
                    return;
                }
                var name = saveWorkerName(who.value);
                close();
                resolve({ who: name, stockId: stockId });
            }
            function onKey(e) {
                if (e.key !== 'Enter') return;
                e.preventDefault();
                e.stopPropagation();
                finish(true);
            }
            function onTake() { finish(true); }
            askWait = function () { finish(false); };
            stock.addEventListener('keydown', onKey);
            take.addEventListener('click', onTake);
        });
    }

    function sheetEcho(txt) {
        var raw = String(txt || '').replace(/&amp;/g, '&');
        var match = raw.match(/https:\/\/script\.googleusercontent\.com\/macros\/echo[^"'\\\s<>]*/);
        if (match) return match[0];
        var exec = raw.match(/https:\/\/script\.google\.com\/macros\/s\/[^"'\\\s<>]+\/exec[^"'\\\s<>]*/);
        return exec ? exec[0] : '';
    }

    function sheetDecode(text) {
        return String(text || '')
            .replace(/^\uFEFF/, '')
            .replace(/^\)\]\}',?\s*/, '')
            .replace(/&quot;/g, '"')
            .replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>');
    }

    function sheetParse(txt) {
        var raw = sheetDecode(txt).trim();
        function read(text) {
            try {
                var data = JSON.parse(text);
                if (data && typeof data === 'object' && ('ok' in data)) return data;
            } catch (e) {}
            return null;
        }
        var direct = read(raw);
        if (direct) return direct;
        var start = raw.indexOf('{"ok"');
        if (start >= 0) {
            var end = raw.lastIndexOf('}');
            if (end > start) {
                var found = read(raw.slice(start, end + 1));
                if (found) return found;
            }
        }
        return null;
    }

    function sheetLocation(res) {
        var headers = String((res && res.responseHeaders) || '');
        var loc = headers.match(/^location:\s*(\S+)/im);
        return loc ? loc[1] : '';
    }

    function sheetReq(url, hop) {
        return new Promise(function (resolve, reject) {
            GM_xmlhttpRequest({
                method: 'GET',
                url: url,
                anonymous: false,
                timeout: 30000,
                onload: function (res) {
                    var txt = res && res.responseText != null ? res.responseText : '';
                    var data = sheetParse(txt);
                    var jump = hop < 4 ? (sheetEcho(txt) || sheetLocation(res)) : '';
                    if (jump === url) jump = '';
                    if (data && data.ok === false) {
                        reject(new Error(data.error || 'Sheet hat abgelehnt'));
                        return;
                    }
                    if (data && data.ok === true && (data.added != null || data.skipped != null)) {
                        resolve(data);
                        return;
                    }
                    if (jump) {
                        sheetReq(jump, hop + 1).then(resolve, reject);
                        return;
                    }
                    var low = String(txt || '').toLowerCase();
                    if (/accounts\.google|servicelogin|anmelden|zugriffsanfrage/.test(low)) reject(new Error('Google-Login fehlt fürs Sheet'));
                    else if (data && data.ok === true) reject(new Error('Web-App neu veröffentlichen'));
                    else reject(new Error('Sheet Antwort ' + (res && res.status)));
                },
                onerror: function () { reject(new Error('Sheet nicht erreichbar')); },
                ontimeout: function () { reject(new Error('Sheet Timeout')); }
            });
        });
    }

    function postAagSheet(row) {
        var url = sheetUrl();
        if (!url) return Promise.resolve({ ok: false, error: 'Sheet-Link fehlt' });
        if (!row || !row.article) return Promise.resolve({ ok: false, error: 'Teil fehlt' });
        if (row.token && sheetBusyToken === row.token) return Promise.resolve({ ok: false, error: 'läuft' });
        sheetBusyToken = row.token || 'x';
        var payload = JSON.stringify({ action: 'aagAppend', entries: [row] });
        var join = url.indexOf('?') === -1 ? '?' : '&';
        return sheetReq(url + join + 'payload=' + encodeURIComponent(payload), 0).then(function (data) {
            if (sheetBusyToken === row.token) sheetBusyToken = '';
            return data;
        }).catch(function (err) {
            if (sheetBusyToken === row.token) sheetBusyToken = '';
            var message = err && err.message ? err.message : 'Sheet fehlgeschlagen';
            if (message === 'unbekannt') message = 'Web-App neu veröffentlichen';
            return { ok: false, error: message };
        });
    }

    async function addArticle(raw, fromEan) {
        var code = cleanScan(raw);
        if (!code || code.length < MIN_LEN) return;
        if (busy) {
            var queued = loadQueue();
            queued.push(code);
            saveQueue(queued);
            setStatus('Wartet: ' + code, 'active');
            return;
        }
        busy = true;
        leaving = false;
        abortMission = false;
        var skipPump = false;
        try {
            var sourceEan = looksLikeEan(fromEan) ? fromEan : '';
            if (looksLikeEan(code)) sourceEan = code;
            if (looksLikeEan(code) && !eanMap[code]) {
                setStatus('EAN ' + code + ' wird extern nachgeschlagen...', 'active');
                var looked = await translateEan(code);
                if (abortMission) return;
                if (!looked.article) {
                    pendingPairEan = code;
                    var why = looked.limited ? 'EAN-Suche ist gerade voll. ' : shortProduct_(looked.name);
                    setStatus(why + 'Keine Hersteller-Nr. AAG-Artikelnummer eintippen, ich merke sie zu ' + code + '.', 'error');
                    focusScan();
                    return;
                }
                sourceEan = code;
                code = looked.article;
            } else if (looksLikeEan(code)) {
                sourceEan = code;
                code = eanMap[code];
            }
            if (!sourceEan && lastEan && eanMap[lastEan] === code) sourceEan = lastEan;
            lastEan = sourceEan;
            var article = code;
            if (pageKind() !== 'search') {
                sessionStorage.setItem(PENDING_KEY, article);
                leaving = true;
                skipPump = true;
                goSearch();
                return;
            }
            await runSearchAdd(article, false);
        } catch (e) {
            setStatus('Fehler beim Suchen.', 'error');
        } finally {
            busy = false;
            if (!skipPump && !abortMission && sessionStorage.getItem(FINISH_KEY) !== '1') pump();
        }
    }

    async function clearSearchForm() {
        var inputs = formInputs();
        if (inputs[1]) setNativeValue(inputs[1], '');
        await sleep(40);
    }

    async function fillArticle(code) {
        var inputs = formInputs();
        var input = inputs[0];
        if (!input) return false;
        if (inputs[1]) setNativeValue(inputs[1], '');
        input.focus();
        setNativeValue(input, code);
        await sleep(100);
        if (normArt(input.value) !== normArt(code)) setNativeValue(input, code);
        return normArt(input.value) === normArt(code);
    }

    function articleOptionScore(text, want) {
        var parts = String(text || '').split(/[\s:()\/,;.]+/);
        var exact = false;
        var i;
        for (i = 0; i < parts.length; i++) {
            if (normArt(parts[i]) === want) exact = true;
        }
        if (!exact) return -1;
        var score = 40;
        if (normArt(parts[parts.length - 1] || '') === want) score += 30;
        score -= Math.min(String(text || '').length, 80) / 10;
        return score;
    }

    function findArticleOption(code) {
        var want = normArt(code);
        var input = formInputs()[0];
        if (!want || !input) return null;
        var box = input.getBoundingClientRect();
        var nodes = document.querySelectorAll('mat-option, [role="option"], li, a, button, div, span');
        var best = null;
        var bestScore = -1;
        var i;
        for (i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (!isVisible(el) || (el.closest && el.closest('#aag-retoure-hud'))) continue;
            if (el === input || (el.contains && el.contains(input))) continue;
            var text = normSpace(el.innerText || el.textContent || '');
            if (!text || text.length > 160) continue;
            var score = articleOptionScore(text, want);
            if (score < 0) continue;
            var kids = el.querySelectorAll('mat-option, [role="option"], li, a, button, div, span');
            var nested = false;
            var k;
            for (k = 0; k < kids.length; k++) {
                if (kids[k] === el) continue;
                var childText = normSpace(kids[k].innerText || kids[k].textContent || '');
                if (childText && childText.length < text.length && articleOptionScore(childText, want) >= 0) {
                    var childRect = kids[k].getBoundingClientRect();
                    if (childRect.height >= 8 && childRect.width >= 30) {
                        nested = true;
                        break;
                    }
                }
            }
            if (nested) continue;
            var rect = el.getBoundingClientRect();
            if (rect.height < 8 || rect.height > 120 || rect.width < 30) continue;
            if (rect.bottom < box.top - 8 || rect.top > box.bottom + 460) continue;
            if (rect.right < box.left - 40 || rect.left > box.right + 80) continue;
            if (score > bestScore) {
                bestScore = score;
                best = el;
            }
        }
        return best;
    }

    async function chooseArticleOption(code, waitMs) {
        var picked = await waitFor(function () { return findArticleOption(code); }, waitMs || 5600);
        if (!picked) return false;
        var host = picked.closest('mat-option, [role="option"], li, a, button') || picked;
        forceClick(host);
        await sleep(280);
        return true;
    }

    async function addFromCard(card, article, hold) {
        var ls = lsOf(card);
        var date = dateOf(card);
        function current() {
            return headerByLs(ls) || card;
        }
        var buttons = returnButtonsNear(current(), article);
        if (!buttons.length) {
            var header = current();
            try { header.scrollIntoView({ block: 'center' }); } catch (e) {}
            expandHeader(header);
            var appeared = await waitFor(function () {
                return returnButtonsNear(current(), article).length ? true : null;
            }, 3500);
            if (!appeared) return 'miss';
            buttons = returnButtonsNear(current(), article);
        }
        var addButton = null;
        var line = null;
        for (var i = 0; i < buttons.length; i++) {
            var row = lineOf(buttons[i], article);
            if (!row) {
                var foundRows = articleRows(article);
                var r;
                for (r = 0; r < foundRows.length; r++) {
                    if (returnButtonOnRow(foundRows[r]) === buttons[i]) {
                        row = foundRows[r];
                        break;
                    }
                }
            }
            var qty = returnableQty(row);
            if (qty === 0) continue;
            if (!isAddButton(buttons[i])) continue;
            addButton = buttons[i];
            line = row;
            break;
        }
        if (!addButton) return buttons.length ? 'already' : 'miss';
        var fresh = current();
        if (hold) return { hold: true, button: addButton, name: lineName(line, article), ls: lsOf(fresh) || ls, date: dateOf(fresh) || date };
        return clickReturnButton(addButton, { name: lineName(line, article), ls: lsOf(fresh) || ls, date: dateOf(fresh) || date });
    }

    async function clickReturnButton(addButton, info) {
        var before = cartBadge();
        forceClick(addButton);
        var proof = await waitFor(function () {
            if (cartBadge() > before) return 'added';
            if (!isAddButton(addButton)) return 'flipped';
            return null;
        }, 4000);
        if (cartBadge() < before) {
            forceClick(addButton);
            return 'fail';
        }
        if (!proof) return 'fail';
        return { ok: true, name: info.name || '', ls: info.ls || '', date: info.date || '' };
    }

    async function runSearchAdd(code, eanFailed) {
        if (abortMission) return;
        setStatus('Suche ' + code + '...', 'active');
        var ready = await waitFor(function () { return formInputs()[0] || null; }, 8000);
        if (!ready) {
            setStatus('Suchfeld Artikelnummer nicht gefunden.', 'error');
            return;
        }
        await clearSearchForm();
        if (abortMission) return;
        var filled = await fillArticle(code);
        if (!filled) {
            setStatus('Artikelnummer steht nicht im Suchfeld.', 'error');
            return;
        }
        await chooseArticleOption(code, lastEan ? 5600 : 1800);
        if (abortMission) return;
        var suchen = findButton(/suchen$/, /rueckgaben/);
        if (!suchen) {
            setStatus('Button Suchen nicht gefunden.', 'error');
            return;
        }
        if (findArticleOption(code)) await chooseArticleOption(code, 1200);
        var beforeSig = deliveryCards().map(lsOf).join('|');
        var sawSpinner = false;
        suchen.click();
        var result = await waitFor(function () {
            var option = findArticleOption(code);
            if (option) return { option: option };
            var spinning = !!document.querySelector('mat-progress-spinner, mat-spinner, .mat-mdc-progress-spinner');
            if (spinning) sawSpinner = true;
            if (spinning) return null;
            var cards = deliveryCards();
            var sig = cards.map(lsOf).join('|');
            if (!sig) return null;
            if (sig !== beforeSig || sawSpinner) return cards;
            return null;
        }, 8000);
        if (result && result.option) {
            var optionHost = result.option.closest('mat-option, [role="option"], li, a, button') || result.option;
            forceClick(optionHost);
            await sleep(280);
            suchen = findButton(/suchen$/, /rueckgaben/);
            if (suchen) suchen.click();
            beforeSig = deliveryCards().map(lsOf).join('|');
            sawSpinner = false;
            result = await waitFor(function () {
                var still = findArticleOption(code);
                if (still) {
                    forceClick(still.closest('mat-option, [role="option"], li, a, button') || still);
                    return null;
                }
                var spinning = !!document.querySelector('mat-progress-spinner, mat-spinner, .mat-mdc-progress-spinner');
                if (spinning) sawSpinner = true;
                if (spinning) return null;
                var cards = deliveryCards();
                var sig = cards.map(lsOf).join('|');
                if (!sig) return null;
                if (sig !== beforeSig || sawSpinner) return cards;
                return null;
            }, 8000);
        }
        focusScan();
        var cards = (result && result.option) ? deliveryCards() : (result || deliveryCards());
        if (!cards.length) {
            if (lastEan && code !== lastEan) {
                pendingPairEan = lastEan;
                setStatus(code + ' ist die Hersteller-Nr zu ' + lastEan + '. AAG findet dazu keine Lieferung. AAG-Artikelnummer eintippen.', 'error');
            } else if (eanFailed) setStatus('EAN nicht aufgelöst. Artikelnummer eintippen, zum Beispiel J1340842, und Enter.', 'error');
            else setStatus(code + ' nicht gefunden. Letztes Jahr steht oben, sonst die Nummer prüfen.', 'error');
            focusScan();
            return;
        }
        var card = cards[0];
        if (abortMission) return;
        setStatus('Nur ' + (lsOf(card) || 'die neueste Lieferung') + ' vom ' + (dateOf(card) || '?') + '.', 'active');
        var found = await addFromCard(card, code, true);
        if (found === 'already') {
            if (lastEan) rememberEan(lastEan, code);
            setStatus(code + ' liegt in der neuesten Lieferung schon im Warenkorb.', 'warn');
            focusScan();
            return;
        }
        if (!found || !found.hold) {
            if (lastEan && code !== lastEan) {
                pendingPairEan = lastEan;
                setStatus(code + ' steht nicht als Artikel auf der neuesten Lieferung. AAG-Artikelnummer eintippen, ich merke sie zu ' + lastEan + '.', 'error');
            } else if (eanFailed) setStatus('EAN nicht aufgelöst. Artikelnummer eintippen und Enter.', 'error');
            else setStatus(code + ' steht in der neuesten Lieferung nicht zum Zurückgeben. Ältere Lieferungen bleiben zu.', 'error');
            focusScan();
            return;
        }
        if (lastEan) rememberEan(lastEan, code);
        var extra = await askStock(found, code);
        if (!extra || abortMission) {
            setStatus('Zurückgeben abgebrochen.', 'warn');
            focusScan();
            return;
        }
        var added = await clickReturnButton(found.button, found);
        if (!(added && added.ok)) {
            setStatus(code + ' ist nicht in den Warenkorb gegangen.', 'error');
            focusScan();
            return;
        }
        rememberAdded({ article: code, name: added.name || '', ls: added.ls || '' });
        var stockBox = document.getElementById('aag-retoure-stock');
        if (stockBox) stockBox.value = '';
        var sheet = await postAagSheet({
            token: [code, added.ls || '', extra.stockId || ''].join('|'),
            orderId: added.ls || '',
            stockId: extra.stockId || '',
            article: code,
            name: added.name || '',
            supplier: 'AAG',
            who: extra.who || '',
            ean: lastEan && lastEan !== code ? lastEan : ''
        });
        if (sheet && sheet.added) setStatus(code + ' im Warenkorb und im Blatt KNOLL / AAG.', 'done');
        else if (sheet && sheet.skipped) setStatus(code + ' im Warenkorb. Steht schon im Blatt.', 'done');
        else setStatus(code + ' im Warenkorb. Sheet: ' + ((sheet && sheet.error) || 'nicht geschrieben') + '.', 'warn');
        focusScan();
    }

    function pump() {
        if (busy || abortMission || sessionStorage.getItem(FINISH_KEY) === '1') return;
        if (pageKind() !== 'search') return;
        var queue = loadQueue();
        if (!queue.length) return;
        var next = queue.shift();
        saveQueue(queue);
        addArticle(next);
    }

    function submitScan(raw) {
        var code = cleanScan(raw);
        if (!code || code.length < MIN_LEN) return;
        if (askWait) return;
        var paired = '';
        if (pendingPairEan && !looksLikeEan(code)) {
            rememberEan(pendingPairEan, code);
            paired = pendingPairEan;
            pendingPairEan = '';
        }
        addArticle(code, paired);
    }

    function isGrundText(text) {
        if (!text || text.length > 48) return false;
        if (text.indexOf('falscher artikel') !== -1) return true;
        return text.indexOf('bitte grund') !== -1;
    }

    function inReasonMenu(el) {
        var parent = el && el.parentElement;
        if (!parent) return false;
        var text = fold(parent.innerText || '');
        var marks = ['zur auswahl bestellt', 'doppelt geliefert', 'falscher artikel', 'falscher inhalt', 'transportschaden'];
        var n = 0;
        for (var i = 0; i < marks.length; i++) {
            if (text.indexOf(marks[i]) !== -1) n++;
        }
        return n >= 3;
    }

    function selectInRow(trigger) {
        if (!trigger) return null;
        if (trigger.tagName === 'SELECT') return trigger;
        var node = trigger;
        var depth;
        for (depth = 0; depth < 6 && node; depth++) {
            var selects = node.querySelectorAll ? node.querySelectorAll('select') : [];
            var found = [];
            var i;
            for (i = 0; i < selects.length; i++) {
                if (selects[i].closest('#aag-retoure-hud')) continue;
                var s;
                for (s = 0; s < selects[i].options.length; s++) {
                    if (fold(selects[i].options[s].text).indexOf('falscher artikel') !== -1) found.push(selects[i]);
                }
            }
            if (found.length === 1) return found[0];
            if (found.length > 1) return null;
            node = node.parentElement;
        }
        return null;
    }

    function grundIsSet(trigger) {
        if (!trigger) return false;
        if (trigger.tagName === 'SELECT') {
            var selected = trigger.options[trigger.selectedIndex];
            return !!(selected && fold(selected.text).indexOf('falscher artikel') !== -1);
        }
        return fold(trigger.innerText || '').indexOf('falscher artikel') !== -1 && fold(trigger.innerText || '').indexOf('bitte grund') === -1;
    }

    function applyNativeGrund(select) {
        var i;
        for (i = 0; i < select.options.length; i++) {
            if (fold(select.options[i].text).indexOf('falscher artikel') === -1) continue;
            select.selectedIndex = i;
            select.options[i].selected = true;
            select.value = select.options[i].value;
            select.dispatchEvent(new Event('input', { bubbles: true }));
            select.dispatchEvent(new Event('change', { bubbles: true }));
            return grundIsSet(select);
        }
        return false;
    }

    async function chooseGrund(trigger) {
        if (grundIsSet(trigger)) return true;
        var select = selectInRow(trigger);
        if (select && applyNativeGrund(select)) {
            await sleep(120);
            return true;
        }
        forceClick(trigger);
        var option = await waitFor(findGrundOption, 2000);
        if (!option) return false;
        var host = option.closest('mat-option, [role="option"], li, button, a') || option;
        forceClick(host);
        var set = await waitFor(function () {
            return grundIsSet(trigger) || (select && grundIsSet(select)) ? true : null;
        }, 1200);
        return !!set;
    }

    function findGrundOption() {
        var nodes = document.querySelectorAll('mat-option, [role="option"], li, button, a, span, div');
        var best = null;
        var bestRank = 9;
        var i;
        for (i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (el.closest && el.closest('#aag-retoure-hud')) continue;
            if (!isVisible(el)) continue;
            if (fold(el.innerText || '') !== 'falscher artikel') continue;
            var rank = 2;
            if (el.tagName === 'MAT-OPTION' || el.getAttribute('role') === 'option') rank = 0;
            else if (el.tagName === 'LI' || el.tagName === 'BUTTON' || el.tagName === 'A') rank = 1;
            if (rank < bestRank) {
                best = el;
                bestRank = rank;
            }
        }
        return best;
    }

    function grundTriggers() {
        var hits = [];
        Array.prototype.forEach.call(document.querySelectorAll('select'), function (select) {
            if (select.closest('#aag-retoure-hud') || !isVisible(select)) return;
            for (var i = 0; i < select.options.length; i++) {
                if (fold(select.options[i].text).indexOf('falscher artikel') !== -1) {
                    hits.push(select);
                    return;
                }
            }
        });
        var nodes = document.querySelectorAll('mat-select, [role="combobox"], button, div, span');
        for (var n = 0; n < nodes.length; n++) {
            var el = nodes[n];
            if (el.closest('#aag-retoure-hud') || !isVisible(el)) continue;
            if (inReasonMenu(el)) continue;
            if (el.closest && el.closest('[role="listbox"], [role="menu"], mat-option, .cdk-overlay-pane')) continue;
            if (!isGrundText(fold(el.innerText || ''))) continue;
            hits.push(el);
        }
        return hits.filter(function (el) {
            return !hits.some(function (other) { return other !== el && el.contains(other); });
        });
    }

    function grundReady() {
        var triggers = grundTriggers();
        if (!triggers.length) return false;
        for (var i = 0; i < triggers.length; i++) {
            if (!grundIsSet(triggers[i])) return false;
        }
        return true;
    }

    async function applyGrundAll() {
        for (var n = 0; n < 20; n++) {
            if (abortMission) return false;
            if (grundReady()) return true;
            var pending = grundTriggers().filter(function (trigger) { return !grundIsSet(trigger); });
            if (!pending.length) return false;
            var total = grundTriggers().length;
            setStatus(GRUND + ' ' + (total - pending.length + 1) + ' von ' + total + '...', 'active');
            var ok = await chooseGrund(pending[0]);
            if (!ok) {
                setStatus(GRUND + ' ' + (total - pending.length + 1) + ' von ' + total + ' noch offen...', 'active');
                await sleep(200);
            }
        }
        return grundReady();
    }

    async function ensureCartReasons() {
        if (busy || !cartVisible() || grundReady()) return;
        if (grundTries >= 4) return;
        grundTries += 1;
        busy = true;
        try {
            setStatus('Setze ' + GRUND + ' auf alle Positionen...', 'active');
            var seen = await waitFor(function () {
                if (!cartHasItems()) return null;
                var count = grundTriggers().length;
                if (!count) return null;
                setStatus(GRUND + ' auf ' + count + (count === 1 ? ' Position...' : ' Positionen...'), 'active');
                return true;
            }, 5000);
            if (!cartHasItems()) {
                setStatus('Warenkorb ist leer.', 'error');
                return;
            }
            if (!seen) {
                setStatus('Warenkorb ohne Grund-Feld.', 'error');
                return;
            }
            var ok = await applyGrundAll();
            if (ok) {
                grundTries = 0;
                setStatus(GRUND + ' steht auf allen Positionen. Drucken und absenden.', 'done');
            } else if (grundTries >= 4) {
                setStatus('Grund ' + GRUND + ' nicht auf allen Positionen gesetzt.', 'error');
            }
        } catch (e) {
            setStatus('Grund ' + GRUND + ' nicht gesetzt.', 'error');
        } finally {
            busy = false;
        }
    }

    function maybeCartReasons() {
        if (cartHold) {
            if (!cartVisible()) cartHold = false;
            return;
        }
        if (!cartVisible()) {
            grundTries = 0;
            return;
        }
        if (!cartHasItems()) return;
        if (grundReady() || busy) return;
        if (Date.now() - grundAt < 1500) return;
        grundAt = Date.now();
        ensureCartReasons();
    }

    function hideForPrint() {
        var hud = document.getElementById('aag-retoure-hud');
        if (hud) hud.style.display = 'none';
        var sync = document.getElementById('aag-lieferschein-sync-status');
        if (sync) sync.setAttribute('data-aag-print-hidden', sync.style.display || '');
        if (sync) sync.style.display = 'none';
    }

    function showAfterPrint() {
        var hud = document.getElementById('aag-retoure-hud');
        if (hud) hud.style.display = '';
        var sync = document.getElementById('aag-lieferschein-sync-status');
        if (sync && sync.hasAttribute('data-aag-print-hidden')) {
            sync.style.display = sync.getAttribute('data-aag-print-hidden');
            sync.removeAttribute('data-aag-print-hidden');
        }
    }

    async function finishCart() {
        if (busy) return;
        busy = true;
        abortMission = false;
        var skipPump = false;
        try {
            if (!cartVisible()) {
                sessionStorage.setItem(FINISH_KEY, '1');
                leaving = true;
                skipPump = true;
                setStatus('Öffne den Warenkorb...', 'active');
                goCart();
                return;
            }
            sessionStorage.removeItem(FINISH_KEY);
            setStatus('Setze ' + GRUND + '...', 'active');
            await waitFor(function () { return cartHasItems() ? true : null; }, 8000);
            if (!cartHasItems()) {
                setStatus('Warenkorb ist leer.', 'error');
                return;
            }
            var grundOk = await applyGrundAll();
            if (!grundOk || abortMission) {
                setStatus('Grund ' + GRUND + ' nicht gesetzt.', 'error');
                return;
            }
            setStatus('Grund steht. Jetzt nochmal Drucken und absenden, dann geht der Druck raus und die Retoure wird abgeschickt.', 'done');
        } catch (e) {
            showAfterPrint();
            setStatus('Drucken oder Absenden hat nicht geklappt.', 'error');
        } finally {
            busy = false;
            leaving = false;
            if (!skipPump) pump();
        }
    }

    function printAndSendSync() {
        var printBtn = findButton(/drucken/);
        var sendBtn = findButton(/absenden/);
        if (!printBtn || !sendBtn) {
            setStatus('Drucken oder Absenden nicht gefunden.', 'error');
            return;
        }
        if (isDisabled(sendBtn)) {
            setStatus('Absenden ist noch gesperrt. Grund prüfen.', 'error');
            return;
        }
        setStatus('Drucke die Liste...', 'active');
        hideForPrint();
        printBtn.click();
        showAfterPrint();
        sendBtn = findButton(/absenden/);
        if (!sendBtn || isDisabled(sendBtn)) {
            setStatus('Gedruckt. Absenden ist noch gesperrt.', 'warn');
            return;
        }
        sendBtn.click();
        setStatus('Sende die Retoure...', 'active');
        setTimeout(function () {
            var dialog = document.querySelector('mat-dialog-container, [role="dialog"]');
            if (dialog) {
                var confirm = findByText(dialog, /^(bestaetigen|ja|ok|absenden)$/, 24);
                if (confirm) confirm.click();
            }
            clearAfterSend();
        }, 700);
    }

    function clearAfterSend() {
        cartHold = true;
        grundTries = 0;
        pendingPairEan = '';
        lastEan = '';
        saveQueue([]);
        sessionStorage.removeItem(PENDING_KEY);
        sessionStorage.removeItem(FINISH_KEY);
        sessionStorage.removeItem(ADDED_KEY);
        sessionStorage.removeItem(TRANSLATE_KEY);
        lookupEanNow = '';
        renderChoices([]);
        var input = document.getElementById('aag-retoure-scan');
        if (input) input.value = '';
        var stock = document.getElementById('aag-retoure-stock');
        if (stock) stock.value = '';
        renderAdded();
        setStatus('Abgeschickt. Bereit für die nächsten.', 'done');
        goSearch();
        focusScan();
    }

    function onFinishClick() {
        if (busy || loadTranslate()) return;
        if (cartVisible() && grundReady()) {
            sessionStorage.removeItem(FINISH_KEY);
            printAndSendSync();
            return;
        }
        finishCart();
    }

    function continueFlow() {
        ensureHud();
        if (!onRetours() || busy) return;
        if (sessionStorage.getItem(FINISH_KEY) === '1' && cartVisible()) {
            finishCart();
            return;
        }
        var pending = sessionStorage.getItem(PENDING_KEY) || '';
        if (pending && pageKind() === 'search') {
            sessionStorage.removeItem(PENDING_KEY);
            addArticle(pending);
            return;
        }
        pump();
    }

    function hookHistory() {
        var hist = pageWin.history;
        if (hist.pushState && !hist.pushState.__aagRetoure) {
            var original = hist.pushState;
            var wrapped = function () {
                var result = original.apply(this, arguments);
                setTimeout(continueFlow, 40);
                return result;
            };
            wrapped.__aagRetoure = true;
            hist.pushState = wrapped;
        }
        if (hist.replaceState && !hist.replaceState.__aagRetoure) {
            var originalReplace = hist.replaceState;
            var wrappedReplace = function () {
                var result = originalReplace.apply(this, arguments);
                setTimeout(continueFlow, 40);
                return result;
            };
            wrappedReplace.__aagRetoure = true;
            hist.replaceState = wrappedReplace;
        }
        pageWin.addEventListener('popstate', function () { setTimeout(continueFlow, 40); });
    }

    function setStatus(text, tone) {
        var badge = document.getElementById('aag-retoure-badge');
        var message = document.getElementById('aag-retoure-message');
        var hud = document.getElementById('aag-retoure-hud');
        if (hud) hud.setAttribute('data-tone', tone || 'neutral');
        if (badge) badge.textContent = tone === 'error' ? 'FEHLER' : tone === 'warn' ? 'HINWEIS' : tone === 'done' ? 'FERTIG' : tone === 'active' ? 'LÄUFT' : 'BEREIT';
        if (message) message.textContent = text;
    }

    function renderAdded() {
        var host = document.getElementById('aag-retoure-last');
        if (!host) return;
        var list = loadAdded();
        if (!list.length) {
            host.innerHTML = '';
            return;
        }
        var html = '<div class="aag-last-label">In diesem Lauf</div>';
        list.slice(-6).forEach(function (item) {
            html += '<div class="aag-last-row"><b>' + esc(item.article) + '</b><span>' + esc(item.name || '') + '</span><em>' + esc(item.ls || '') + '</em></div>';
        });
        host.innerHTML = html;
    }

    function focusScan() {
        var ask = document.getElementById('aag-retoure-ask');
        if (ask && !ask.hidden) return;
        var input = document.getElementById('aag-retoure-scan');
        if (!input) return;
        setTimeout(function () {
            if (!onRetours()) return;
            if (ask && !ask.hidden) return;
            input.focus();
        }, 60);
    }

    function resetHud() {
        abortMission = true;
        cartHold = false;
        if (askWait) askWait();
        pendingPairEan = '';
        saveQueue([]);
        sessionStorage.removeItem(PENDING_KEY);
        sessionStorage.removeItem(FINISH_KEY);
        sessionStorage.removeItem(ADDED_KEY);
        sessionStorage.removeItem(TRANSLATE_KEY);
        lookupEanNow = '';
        renderChoices([]);
        var input = document.getElementById('aag-retoure-scan');
        if (input) input.value = '';
        var stock = document.getElementById('aag-retoure-stock');
        if (stock) stock.value = '';
        renderAdded();
        setStatus('HUD leer. Name und Sheet-Link bleiben. Der Warenkorb auf der Seite bleibt.', 'neutral');
        setTimeout(function () { abortMission = false; focusScan(); }, 80);
    }

    function hudPosition() {
        try {
            var pos = JSON.parse(localStorage.getItem(HUD_POS_KEY) || 'null');
            if (!pos || typeof pos.left !== 'number' || typeof pos.top !== 'number') return null;
            return pos;
        } catch (e) {
            return null;
        }
    }

    function enableDrag(popup, handle) {
        var state = null;
        function move(e) {
            if (!state) return;
            var maxLeft = Math.max(8, window.innerWidth - popup.offsetWidth - 8);
            var maxTop = Math.max(8, window.innerHeight - popup.offsetHeight - 8);
            var left = Math.max(8, Math.min(maxLeft, e.clientX - state.dx));
            var top = Math.max(8, Math.min(maxTop, e.clientY - state.dy));
            popup.style.left = left + 'px';
            popup.style.top = top + 'px';
            popup.style.right = 'auto';
            state.left = left;
            state.top = top;
        }
        function up() {
            if (state) {
                try { localStorage.setItem(HUD_POS_KEY, JSON.stringify({ left: state.left, top: state.top })); } catch (e) {}
            }
            state = null;
            document.removeEventListener('pointermove', move);
        }
        handle.addEventListener('pointerdown', function (e) {
            if (e.button !== 0 || e.target.closest('button')) return;
            var rect = popup.getBoundingClientRect();
            state = { dx: e.clientX - rect.left, dy: e.clientY - rect.top, left: rect.left, top: rect.top };
            popup.style.left = rect.left + 'px';
            popup.style.top = rect.top + 'px';
            popup.style.right = 'auto';
            document.addEventListener('pointermove', move);
            document.addEventListener('pointerup', up, { once: true });
            e.preventDefault();
        });
    }

    function ensureStyles() {
        if (document.getElementById('aag-retoure-style')) return;
        var style = document.createElement('style');
        style.id = 'aag-retoure-style';
        style.textContent = ''
            + '#aag-retoure-hud{position:fixed;z-index:2147483646;top:84px;right:16px;width:min(380px,calc(100vw - 24px));color:#c9d1d9;font-family:"Segoe UI",Arial,sans-serif;border-radius:22px;overflow:hidden;background:radial-gradient(circle at top left,rgba(220,61,61,.16),transparent 34%),linear-gradient(180deg,rgba(21,27,35,.98),rgba(14,19,26,.98));border:1px solid #2d3642;box-shadow:0 24px 56px rgba(0,0,0,.52);user-select:none}'
            + '#aag-retoure-hud[data-tone="active"]{border-color:rgba(220,61,61,.52)}'
            + '#aag-retoure-hud[data-tone="done"]{border-color:rgba(86,211,100,.55)}'
            + '#aag-retoure-hud[data-tone="error"]{border-color:rgba(248,81,73,.68)}'
            + '#aag-retoure-hud.minimized .aag-retoure-body{display:none}'
            + '#aag-retoure-head{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:14px 14px 12px;cursor:move;border-bottom:1px solid #242d39}'
            + '#aag-retoure-eyebrow{color:#8b949e;font-size:11px;font-weight:800;letter-spacing:1.4px;text-transform:uppercase}'
            + '#aag-retoure-title{color:#f0f6fc;font-size:18px;font-weight:900;letter-spacing:.4px}'
            + '#aag-retoure-hud button{font-family:inherit;cursor:pointer}'
            + '.aag-icon{min-width:36px;height:34px;border-radius:12px;border:1px solid rgba(255,255,255,.12);background:rgba(13,17,23,.72);color:#c9d1d9;font-weight:900}'
            + '.aag-retoure-body{padding:14px}'
            + '#aag-retoure-reset{width:100%;min-height:38px;border:none;border-radius:12px;background:#f0883e;color:#1a1208;font-size:13px;font-weight:900}'
            + '#aag-retoure-reset:hover{background:#ff9d4d}'
            + '#aag-retoure-scan,.aag-field input{width:100%;box-sizing:border-box;margin-top:8px;min-height:40px;border-radius:12px;border:1px solid #3a4452;background:#0d1117;color:#f0f6fc;font-size:15px;font-weight:700;padding:0 12px}'
            + '.aag-field{display:block;margin-top:8px}'
            + '.aag-field span{display:block;color:#8b949e;font-size:10px;font-weight:800;letter-spacing:.6px;text-transform:uppercase;margin-bottom:4px}'
            + '.aag-field input{margin-top:0;min-height:36px;font-size:14px}'
            + '.aag-sheet-row{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-top:8px}'
            + '#aag-retoure-sheet-toggle{min-height:32px;padding:0 12px;border-radius:10px;border:1px solid rgba(255,255,255,.12);background:rgba(13,17,23,.72);color:#d7dee8;font-size:12px;font-weight:800}'
            + '.aag-sheet-state{font-size:12px;font-weight:800}'
            + '.aag-sheet-state.is-ok{color:#86f0b0}'
            + '.aag-sheet-state.is-bad{color:#fecaca}'
            + '#aag-retoure-take{width:100%;min-height:36px;border:none;border-radius:12px;background:#f0883e;color:#1a1208;font-size:13px;font-weight:900}'
            + '.aag-actions{display:flex;gap:8px;margin-top:10px}'
            + '.aag-actions button{flex:1;min-height:40px;border-radius:12px;border:1px solid rgba(255,255,255,.12);background:rgba(13,17,23,.82);color:#d7dee8;font-size:12px;font-weight:900}'
            + '#aag-retoure-finish{border-color:rgba(220,61,61,.55);color:#f07070;background:rgba(220,61,61,.12)}'
            + '#aag-retoure-badge{display:inline-flex;margin-top:12px;min-height:28px;align-items:center;padding:4px 10px;border-radius:999px;border:2px solid rgba(220,61,61,.38);color:#dc3d3d;font-size:11px;font-weight:900;letter-spacing:1px}'
            + '#aag-retoure-message{margin-top:8px;color:#d7dee8;font-size:14px;font-weight:700;line-height:1.4;user-select:text}'
            + '#aag-retoure-last{margin-top:10px}'
            + '.aag-last-label{color:#8b949e;font-size:10px;font-weight:800;letter-spacing:1px;text-transform:uppercase;margin-bottom:4px}'
            + '.aag-last-row{display:flex;gap:8px;align-items:baseline;font-size:12px;padding:3px 0;border-top:1px solid rgba(255,255,255,.06)}'
            + '.aag-last-row b{color:#f07070}'
            + '.aag-last-row span{flex:1;color:#f0f6fc;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}'
            + '.aag-last-row em{color:#8b949e;font-style:normal;font-size:10px}'
            + '#aag-retoure-choices{display:flex;flex-wrap:wrap;gap:8px;margin-top:10px}'
            + '.aag-choice{min-height:36px;padding:6px 10px;border-radius:10px;border:1px solid rgba(220,61,61,.45);background:rgba(220,61,61,.12);color:#f07070;font-size:13px;font-weight:800}'
            + '.aag-foot{display:flex;justify-content:space-between;gap:10px;margin-top:12px;padding-top:10px;border-top:1px solid rgba(255,255,255,.07);color:#8b949e;font-size:12px;font-weight:800}'
            + '.aag-credit{text-align:right;margin-top:4px;color:#6e7681;font-size:10px;font-weight:600}'
            + '.aag-credit a{color:#8b949e}'
            + '@media print{#aag-retoure-hud,#aag-lieferschein-sync-status{display:none !important}}';
        document.documentElement.appendChild(style);
    }

    function ensureHud() {
        if (!document.body) return;
        var hud = document.getElementById('aag-retoure-hud');
        if (!onRetours() && !loadTranslate()) {
            if (hud) hud.style.display = 'none';
            return;
        }
        if (hud) {
            hud.style.display = '';
            return;
        }
        ensureStyles();
        hud = document.createElement('section');
        hud.id = 'aag-retoure-hud';
        hud.setAttribute('data-tone', 'neutral');
        hud.innerHTML = ''
            + '<div id="aag-retoure-head">'
            + '  <div><div id="aag-retoure-eyebrow">AAG Retoure</div><div id="aag-retoure-title">AAG RETOURE v' + BOT_VERSION + '</div></div>'
            + '  <button type="button" class="aag-icon" id="aag-retoure-min">–</button>'
            + '</div>'
            + '<div class="aag-retoure-body">'
            + '  <button type="button" id="aag-retoure-reset">Zurücksetzen</button>'
            + '  <input id="aag-retoure-scan" autocomplete="off" spellcheck="false" placeholder="Artikelnummer oder EAN">'
            + '  <label class="aag-field"><span>Stock-ID</span><input id="aag-retoure-stock" autocomplete="off" spellcheck="false" placeholder="z.B. JR97017"></label>'
            + '  <label class="aag-field"><span>Name</span><input id="aag-retoure-who" autocomplete="off" spellcheck="false" placeholder="z.B. Semmel"></label>'
            + '  <div class="aag-sheet-row"><button type="button" id="aag-retoure-sheet-toggle">Sheet</button><span id="aag-retoure-sheet-state"></span></div>'
            + '  <label class="aag-field" id="aag-retoure-sheet-box"><span>Sheet</span><input id="aag-retoure-sheet" autocomplete="off" spellcheck="false" placeholder="https://script.google.com/.../exec"><span>Barcode Lookup Key</span><input id="aag-retoure-barcode" autocomplete="off" spellcheck="false" placeholder="optional, sonst ean-suche.net"></label>'
            + '  <div id="aag-retoure-ask" hidden><button type="button" id="aag-retoure-take">Zurückgeben</button></div>'
            + '  <div class="aag-actions">'
            + '    <button type="button" id="aag-retoure-cart">Warenkorb</button>'
            + '    <button type="button" id="aag-retoure-finish">Drucken und absenden</button>'
            + '  </div>'
            + '  <div id="aag-retoure-badge">BEREIT</div>'
            + '  <div id="aag-retoure-message">Artikelnummer scannen. Eine EAN wird zur Hersteller-Nr. AAG selbst kennt keine EAN. Geöffnet wird nur die neueste Lieferung.</div>'
            + '  <div id="aag-retoure-choices"></div>'
            + '  <div id="aag-retoure-last"></div>'
            + '  <div class="aag-foot"><span>Bot v' + BOT_VERSION + '</span><span>ESC Stop</span></div>'
            + '  <div class="aag-credit">Data: <a href="https://myproduct.guru" target="_blank" rel="noopener">ProductGuru</a> · <a href="https://ean-suche.net" target="_blank" rel="noopener">ean-suche.net</a> · by Arasaka</div>'
            + '</div>';
        document.body.appendChild(hud);
        var pos = hudPosition();
        if (pos) {
            hud.style.left = pos.left + 'px';
            hud.style.top = pos.top + 'px';
            hud.style.right = 'auto';
        }
        enableDrag(hud, document.getElementById('aag-retoure-head'));
        document.getElementById('aag-retoure-min').addEventListener('click', function () {
            hud.classList.toggle('minimized');
            this.textContent = hud.classList.contains('minimized') ? '+' : '–';
        });
        document.getElementById('aag-retoure-reset').addEventListener('click', resetHud);
        document.getElementById('aag-retoure-cart').addEventListener('click', function () {
            grundTries = 0;
            grundAt = 0;
            setStatus('Warenkorb. Setze ' + GRUND + '...', 'active');
            goCart();
            setTimeout(maybeCartReasons, 500);
        });
        document.getElementById('aag-retoure-finish').addEventListener('click', onFinishClick);
        var who = document.getElementById('aag-retoure-who');
        who.value = workerName();
        who.addEventListener('change', function () { saveWorkerName(who.value); });
        var sheet = document.getElementById('aag-retoure-sheet');
        sheet.value = String(gmGet(SHEET_URL_KEY, '') || '');
        function commitSheet() {
            if (!saveSheetUrl(sheet.value)) {
                setStatus('Sheet-Link muss auf script.google.com enden mit /exec.', 'error');
                var box = document.getElementById('aag-retoure-sheet-box');
                if (box) box.hidden = false;
                return;
            }
            paintSheet();
            if (sheetUrl()) setStatus('Sheet-Link gespeichert.', 'done');
        }
        sheet.addEventListener('change', commitSheet);
        sheet.addEventListener('keydown', function (e) {
            if (e.key !== 'Enter') return;
            e.preventDefault();
            e.stopPropagation();
            commitSheet();
        });
        var barcodeKey = document.getElementById('aag-retoure-barcode');
        if (barcodeKey) {
            barcodeKey.value = String(gmGet(BARCODE_KEY, '') || '');
            barcodeKey.addEventListener('change', function () {
                gmSet(BARCODE_KEY, String(barcodeKey.value || '').trim());
                setStatus(String(barcodeKey.value || '').trim() ? 'Barcode Lookup Key gespeichert.' : 'Barcode Lookup Key entfernt.', 'done');
            });
        }
        document.getElementById('aag-retoure-sheet-toggle').addEventListener('click', function () {
            var box = document.getElementById('aag-retoure-sheet-box');
            if (!box) return;
            box.hidden = !box.hidden;
            if (!box.hidden) sheet.focus();
        });
        paintSheet();
        var scan = document.getElementById('aag-retoure-scan');
        scan.addEventListener('keydown', function (e) {
            if (e.key === 'Escape') return;
            var now = Date.now();
            if (now - lastKeyAt > 80) scanBurst = 0;
            lastKeyAt = now;
            scanBurst += 1;
            if (e.key !== 'Enter') return;
            e.preventDefault();
            e.stopPropagation();
            if (scanTimer) clearTimeout(scanTimer);
            var value = scan.value;
            scan.value = '';
            scanBurst = 0;
            submitScan(value);
        });
        scan.addEventListener('input', function () {
            if (scanTimer) clearTimeout(scanTimer);
            scanTimer = setTimeout(function () {
                if (scanBurst < 4) return;
                if (cleanScan(scan.value).length < 4) return;
                var value = scan.value;
                scan.value = '';
                scanBurst = 0;
                submitScan(value);
            }, 200);
        });
        renderAdded();
        focusScan();
    }

    function boot() {
        if (!document.body) {
            document.addEventListener('DOMContentLoaded', boot);
            return;
        }
        try { sessionStorage.removeItem(TRANSLATE_KEY); } catch (e) {}
        ensureHud();
        window.addEventListener('beforeprint', hideForPrint);
        window.addEventListener('afterprint', showAfterPrint);
        document.addEventListener('keydown', function (e) {
            if (e.key !== 'Escape') return;
            if (!onRetours() && !loadTranslate()) return;
            abortMission = true;
            if (askWait) askWait();
            pendingPairEan = '';
            saveQueue([]);
            sessionStorage.removeItem(PENDING_KEY);
            sessionStorage.removeItem(FINISH_KEY);
            sessionStorage.removeItem(TRANSLATE_KEY);
            lookupEanNow = '';
            renderChoices([]);
            setStatus('Stop. Nächster Scan geht wieder.', 'warn');
            setTimeout(function () { abortMission = false; }, 50);
        }, true);
        setInterval(function () {
            var pathChanged = location.pathname !== lastPath;
            lastPath = location.pathname;
            if (pathChanged || sessionStorage.getItem(FINISH_KEY) === '1' || sessionStorage.getItem(PENDING_KEY)) continueFlow();
            else ensureHud();
            maybeCartReasons();
        }, 700);
        continueFlow();
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
    else boot();
})();
