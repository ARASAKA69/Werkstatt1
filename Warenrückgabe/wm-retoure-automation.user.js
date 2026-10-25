// ==UserScript==
// @name         WM Warenrückgabe Retoure Bot
// @namespace    http://tampermonkey.net/
// @version      5.0
// @description  Automatisiert WM Warenrückgabe per EAN-Scan und schreibt die Teile ins WM-Blatt
// @author       ARASAKA
// @match        *://*.customer-de.wm.de/*
// @match        *://customer-de.wm.de/*
// @updateURL    https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/Warenr%C3%BCckgabe/wm-retoure-automation.user.js
// @downloadURL  https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/Warenr%C3%BCckgabe/wm-retoure-automation.user.js
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      script.google.com
// @connect      script.googleusercontent.com
// ==/UserScript==

(function() {
    'use strict';

    const BOT_VERSION = '5.0';
    const MITTEILUNG_TEXT = 'Passt Nicht.';
    const HUD_POS_KEY = 'wm_retoure_hud_position';
    const PENDING_NEUE_SUCHE_KEY = 'wm_retoure_pending_neue_suche';
    const ACTIVE_FLOW_KEY = 'wm_retoure_active_flow';
    const LAST_EAN_KEY = 'wm_retoure_last_ean';
    const LAST_ADDED_KEY = 'wm_retoure_last_added';
    const ORDER_CTX_KEY = 'wm_retoure_order_ctx';
    const PENDING_SHEET_KEY = 'wm_retoure_pending_sheet';
    const WORKER_KEY = 'wm_retoure_worker';
    const SHEET_URL_KEY = 'wm_retoure_sheet_url';
    const MIN_SEARCH_LEN = 3;
    const MIN_SCAN_LEN = 8;
    const SCAN_IDLE_MS = 2800;
    const PAGE_WAIT_MS = 400;
    const SUBMIT_DELAY_MS = 350;

    let abortMission = false;
    let isRunning = false;
    let searchSubmitted = false;
    let searchListenersReady = false;
    let userEditingSearch = false;
    let pausePageWatcherUntil = 0;
    let lastHudSignature = '';
    let lastInputChangeAt = 0;
    let searchFocusDone = false;
    let searchFocusTimer = null;
    let scanTimer = null;
    let searchInput = null;
    let pageWatcherStarted = false;
    let lastHandledStep = '';
    let printButtonAdded = false;
    let sheetBusyToken = '';
    let qtyOpening = false;

    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            abortMission = true;
            isRunning = false;
            if (scanTimer) clearTimeout(scanTimer);
            sessionStorage.removeItem(PENDING_NEUE_SUCHE_KEY);
            sessionStorage.removeItem(ACTIVE_FLOW_KEY);
            sessionStorage.removeItem(LAST_ADDED_KEY);
            searchSubmitted = false;
            searchListenersReady = false;
            userEditingSearch = false;
            searchFocusDone = false;
            hideQuantityPicker();
            showHud('WM RETOURE STOP', 'Prozess gestoppt. ESC erneut zum Schließen.', true);
        }
    });

    function sleep(ms) {
        return new Promise(function(resolve) {
            setTimeout(resolve, ms);
        });
    }

    function pageWindow() {
        try {
            if (typeof unsafeWindow !== 'undefined' && unsafeWindow) return unsafeWindow;
        } catch (e) {}
        return window;
    }

    function gmGet(key) {
        try { return GM_getValue(key, ''); } catch (e) { return ''; }
    }

    function gmSet(key, value) {
        try { GM_setValue(key, value); } catch (e) {}
    }

    function workerName() {
        return normalizeText(gmGet(WORKER_KEY));
    }

    function saveWorkerName(value) {
        var name = normalizeText(value).slice(0, 40);
        gmSet(WORKER_KEY, name);
        return name;
    }

    function sheetUrl() {
        var saved = String(gmGet(SHEET_URL_KEY) || '').trim();
        if (/^https:\/\/script\.google\.com\/.+\/exec\/?$/i.test(saved)) return saved;
        return '';
    }

    function saveSheetUrl(value) {
        var url = String(value || '').trim();
        if (!url) {
            gmSet(SHEET_URL_KEY, '');
            return true;
        }
        if (!/^https:\/\/script\.google\.com\/.+\/exec\/?$/i.test(url)) return false;
        gmSet(SHEET_URL_KEY, url);
        return true;
    }

    function cleanStockId(value) {
        var raw = String(value || '').toUpperCase().replace(/\s+/g, '');
        if (!raw) return '';
        var match = raw.match(/[A-Z]{2,3}\d{4,8}/);
        if (!match) return null;
        if (match[0] !== raw && raw.length > match[0].length + 2) return null;
        return match[0];
    }

    function forceClick(el) {
        if (!el || abortMission) return;
        el.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
        el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
        el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        try { el.click(); } catch (e) {}
    }

    function saveLastAdded(info) {
        try {
            if (!info) return;
            sessionStorage.setItem(LAST_ADDED_KEY, JSON.stringify({
                part: String(info.part || ''),
                name: String(info.name || ''),
                qty: String(info.qty || ''),
                stockId: String(info.stockId || ''),
                orderId: String(info.orderId || ''),
                who: String(info.who || ''),
                token: String(info.token || ''),
                sheetNote: String(info.sheetNote || ''),
                ts: Date.now()
            }));
        } catch (e) {}
    }

    function loadLastAdded() {
        try {
            var raw = sessionStorage.getItem(LAST_ADDED_KEY);
            if (!raw) return null;
            var data = JSON.parse(raw);
            if (!data || typeof data !== 'object') return null;
            if (!data.part && !data.name) return null;
            return data;
        } catch (e) {
            return null;
        }
    }

    function markSheetNote(note, token) {
        var info = loadLastAdded();
        if (!info) return;
        if (token && info.token && info.token !== token) return;
        info.sheetNote = note;
        try { sessionStorage.setItem(LAST_ADDED_KEY, JSON.stringify(info)); } catch (e) {}
        var el = document.getElementById('wm-sheet-note');
        if (!el) return;
        el.textContent = note;
        el.className = 'wm-sheet-note' + sheetNoteClass(note);
    }

    function sheetNoteClass(note) {
        if (/fehlt|fehl|login|timeout|antwort|abgelehnt|veröffentlich|ungültig|kaputt|nicht erreichbar/i.test(note || '')) return ' is-bad';
        if (/blatt|schon/i.test(note || '')) return ' is-ok';
        return '';
    }

    function loadOrderContext() {
        try {
            var data = JSON.parse(sessionStorage.getItem(ORDER_CTX_KEY) || '{}');
            if (!data || typeof data !== 'object') return {};
            return data;
        } catch (e) {
            return {};
        }
    }

    function rememberContext(partial) {
        var prev = loadOrderContext();
        var next = {
            orderId: (partial && partial.orderId) || prev.orderId || '',
            part: (partial && partial.part) || prev.part || '',
            name: (partial && partial.name) || prev.name || '',
            brand: (partial && partial.brand) || prev.brand || '',
            supplier: (partial && partial.supplier) || prev.supplier || ''
        };
        try { sessionStorage.setItem(ORDER_CTX_KEY, JSON.stringify(next)); } catch (e) {}
    }

    function headerIndexMap(table) {
        var map = {};
        if (!table) return map;
        var rows = table.querySelectorAll('tr');
        var r;
        for (r = 0; r < Math.min(rows.length, 4); r++) {
            var cells = rows[r].querySelectorAll('th, td');
            if (!cells.length) continue;
            var texts = [];
            var i;
            var headerish = 0;
            for (i = 0; i < cells.length; i++) {
                texts.push(normalizeText(cells[i].textContent));
                if (cells[i].tagName === 'TH') headerish++;
            }
            if (headerish || /bestellt|katalog|bezeichnung|lieferschein|beleg/i.test(texts.join(' '))) {
                for (i = 0; i < texts.length; i++) map[i] = texts[i];
                return map;
            }
        }
        return map;
    }

    function pickHeaderValue(headers, cells, needles) {
        var i;
        for (i = 0; i < cells.length; i++) {
            var label = String(headers[i] || '').toLowerCase();
            var n;
            for (n = 0; n < needles.length; n++) {
                if (label.indexOf(needles[n]) !== -1) return cells[i];
            }
        }
        return '';
    }

    function rememberRow(btn) {
        var row = btn && btn.closest ? btn.closest('tr') : null;
        if (!row) return;
        var table = row.closest('table');
        var headers = headerIndexMap(table);
        var tds = row.querySelectorAll('td');
        var cells = [];
        var i;
        for (i = 0; i < tds.length; i++) cells.push(normalizeText(tds[i].innerText));
        var blob = cells.join(' ');
        var orderId = pickHeaderValue(headers, cells, ['bestell', 'beleg', 'lieferschein', 'rechnung', 'auftrag', 'ls-nr', 'ls nr']);
        if (!orderId) {
            var match = blob.match(/\b\d{7,12}\b/);
            if (match) orderId = match[0];
        }
        rememberContext({
            orderId: orderId,
            part: pickHeaderValue(headers, cells, ['katalog', 'artikel-nr', 'artikelnummer', 'artikel nr']),
            name: pickHeaderValue(headers, cells, ['bezeichnung']),
            brand: pickHeaderValue(headers, cells, ['hersteller', 'marke', 'fabrikat']),
            supplier: pickHeaderValue(headers, cells, ['lieferant'])
        });
    }

    function readLabeledFormValue(labelNeedles) {
        var needles = Array.isArray(labelNeedles) ? labelNeedles : [labelNeedles];
        var labels = document.querySelectorAll('label, td, th, span, div, legend, b, strong');
        for (var i = 0; i < labels.length; i++) {
            var txt = normalizeText(labels[i].textContent).toLowerCase().replace(/:$/, '');
            var matched = false;
            for (var n = 0; n < needles.length; n++) {
                var needle = String(needles[n] || '').toLowerCase();
                if (txt === needle || txt.indexOf(needle) === 0) {
                    matched = true;
                    break;
                }
            }
            if (!matched) continue;
            var scope = labels[i].closest('.fmRowSrd, tr, .fmLi, div') || labels[i].parentElement;
            for (var p = 0; p < 5 && scope; p++) {
                var input = scope.querySelector('input[type="text"], input:not([type]), input[readonly], textarea, select');
                if (input) {
                    var val = normalizeText(input.value || (input.options && input.selectedIndex >= 0 ? input.options[input.selectedIndex].text : ''));
                    if (val && val.toLowerCase().indexOf(needles[0].toLowerCase()) === -1) return val;
                }
                var cells = scope.querySelectorAll('td, span, div, font');
                for (var c = 0; c < cells.length; c++) {
                    if (cells[c] === labels[i] || labels[i].contains(cells[c])) continue;
                    var cellText = normalizeText(cells[c].textContent);
                    if (!cellText) continue;
                    var lower = cellText.toLowerCase();
                    var skip = false;
                    for (var s = 0; s < needles.length; s++) {
                        if (lower === String(needles[s]).toLowerCase() || lower.indexOf(String(needles[s]).toLowerCase() + ':') === 0) {
                            skip = true;
                            break;
                        }
                    }
                    if (skip) continue;
                    if (cellText.length > 80) continue;
                    if (cells[c].querySelector('input, select, textarea, button')) continue;
                    return cellText;
                }
                scope = scope.parentElement;
            }
        }
        return '';
    }

    function readQuantityPageArticleInfo(qty) {
        var ctx = loadOrderContext();
        var part = readLabeledFormValue(['Katalog-Nr.', 'Katalog-Nr', 'Katalognummer', 'Artikelnummer', 'Artikel-Nr.', 'Artikel-Nr']);
        var name = readLabeledFormValue(['Bezeichnung', 'Artikelbezeichnung']);
        var brand = readLabeledFormValue(['Hersteller', 'Marke', 'Fabrikat']);
        var supplier = readLabeledFormValue(['Lieferant', 'Großhändler']);
        var orderId = readLabeledFormValue(['Bestellnummer', 'Bestell-Nr.', 'Bestell-Nr', 'Belegnummer', 'Beleg-Nr.', 'Auftragsnummer', 'Lieferschein', 'Rechnungsnummer']);
        var ean = sessionStorage.getItem(LAST_EAN_KEY) || '';
        if (!part) part = ctx.part || '';
        if (!name) name = ctx.name || '';
        if (!brand) brand = ctx.brand || '';
        if (!supplier) supplier = ctx.supplier || '';
        if (!orderId) orderId = ctx.orderId || '';
        if (!part) part = ean;
        return {
            part: part,
            name: name,
            qty: qty != null ? String(qty) : '',
            brand: brand,
            supplier: supplier,
            orderId: orderId,
            ean: ean
        };
    }

    function normalizeText(value) {
        return String(value || '').replace(/\s+/g, ' ').trim();
    }

    function isVisible(el) {
        if (!el || el.disabled) return false;
        var rect = el.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
    }

    function findClickableByText(text, exact) {
        var nodes = document.querySelectorAll('input[type="button"], input[type="submit"], input[type="image"], button, a');
        for (var i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (!isVisible(el)) continue;
            var label = normalizeText(el.value || el.alt || el.textContent);
            if (exact ? label === text : label.indexOf(text) !== -1) return el;
        }
        return null;
    }

    async function waitForClickableByText(text, exact, timeout) {
        var start = Date.now();
        while (Date.now() - start < (timeout || 12000)) {
            if (abortMission) return null;
            var el = findClickableByText(text, exact);
            if (el) return el;
            await sleep(180);
        }
        return null;
    }

    function parseGermanDate(value) {
        var m = String(value || '').match(/(\d{2})\.(\d{2})\.(\d{4})/);
        if (!m) return null;
        return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
    }

    function pageBodyText() {
        if (!document.body) return '';
        var clone = document.body.cloneNode(true);
        var overlays = clone.querySelectorAll('#arasaka-batch-popup, #wm-retoure-qty-popup, #wm-retoure-print-btn');
        for (var i = 0; i < overlays.length; i++) overlays[i].remove();
        return clone.innerText;
    }

    function pageHasText(text) {
        return pageBodyText().indexOf(text) !== -1;
    }

    function isCartPage() {
        return pageHasText('Warenkorb - Rückgabeartikel');
    }

    function detectPage() {
        if (isCartPage()) return 'cart';
        if (pageHasText('Rückgabemenge') && (findClickableByText('Übernehmen', true) || document.querySelector('[id*="ddlQuantity"].RadDropDownList'))) return 'quantity';
        if (pageHasText('Artikel zurückgeben')) return 'orders';
        if (pageHasText('Artikel auswählen')) return 'article';
        if (pageHasText('Artikel/Lieferschein suchen') || findClickableByText('Anzeigen', true)) return 'search';
        return 'unknown';
    }

    function findSearchInput() {
        var anzeigen = findClickableByText('Anzeigen', true);
        if (anzeigen) {
            var form = anzeigen.closest('form');
            if (form) {
                var inputs = form.querySelectorAll('input[type="text"], input:not([type])');
                for (var i = inputs.length - 1; i >= 0; i--) {
                    if (inputs[i].offsetParent !== null) return inputs[i];
                }
            }
        }
        var candidates = document.querySelectorAll('input[type="text"]');
        for (var j = 0; j < candidates.length; j++) {
            var input = candidates[j];
            if (input.offsetParent === null) continue;
            var box = input.closest('td, div, form');
            if (box && normalizeText(box.textContent).indexOf('Katalog') !== -1) return input;
        }
        for (var k = 0; k < candidates.length; k++) {
            if (candidates[k].offsetParent !== null) return candidates[k];
        }
        return null;
    }

    function findOrdersTable() {
        var tables = document.querySelectorAll('table');
        for (var i = 0; i < tables.length; i++) {
            if (normalizeText(tables[i].innerText).indexOf('Bestellt am') !== -1) return tables[i];
        }
        return null;
    }

    function getTableColumnIndex(table, headerText) {
        var headers = table.querySelectorAll('th, td');
        for (var i = 0; i < headers.length; i++) {
            if (normalizeText(headers[i].textContent) === headerText) return i;
        }
        return -1;
    }

    function findNewestRueckgabeRow() {
        var table = findOrdersTable();
        if (!table) return null;
        var dateIdx = getTableColumnIndex(table, 'Bestellt am');
        var rows = table.querySelectorAll('tr');
        var bestRow = null;
        var bestDate = null;
        for (var i = 0; i < rows.length; i++) {
            var cells = rows[i].querySelectorAll('td');
            if (!cells.length) continue;
            var btn = null;
            var inputs = cells[0].querySelectorAll('input[type="button"], input[type="submit"], button');
            for (var b = 0; b < inputs.length; b++) {
                if (normalizeText(inputs[b].value || inputs[b].textContent) === 'Rückgabe') {
                    btn = inputs[b];
                    break;
                }
            }
            if (!btn) continue;
            var dateCell = dateIdx >= 0 && cells[dateIdx] ? cells[dateIdx] : cells[1];
            var date = parseGermanDate(dateCell ? dateCell.textContent : '');
            if (!date) continue;
            if (!bestDate || date.getTime() > bestDate.getTime()) {
                bestDate = date;
                bestRow = btn;
            }
        }
        return bestRow;
    }

    async function waitForNewestRueckgabeRow(timeout) {
        var start = Date.now();
        while (Date.now() - start < (timeout || 10000)) {
            if (abortMission) return null;
            var btn = findNewestRueckgabeRow();
            if (btn) return btn;
            await sleep(180);
        }
        return null;
    }

    function parseNumericLabel(text) {
        var num = parseInt(normalizeText(text), 10);
        return !isNaN(num) && num > 0 ? num : null;
    }

    function getTelerikWidget(controlId) {
        try {
            var page = pageWindow();
            if (typeof page.$find === 'function') return page.$find(controlId);
        } catch (e) {}
        return null;
    }

    function findRueckgabeMengeControl() {
        var telerik = document.querySelector('[id*="RedemptionOfGoods_ddlQuantity"].RadDropDownList, [id*="ddlQuantity"].RadDropDownList');
        if (telerik) return { type: 'telerik', el: telerik, id: telerik.id };

        var rows = document.querySelectorAll('.fmRowSrd, tr, div');
        for (var i = 0; i < rows.length; i++) {
            var rowText = normalizeText(rows[i].textContent);
            if (rowText.indexOf('Rückgabemenge') === -1) continue;
            var ddl = rows[i].querySelector('.RadDropDownList');
            if (ddl) return { type: 'telerik', el: ddl, id: ddl.id };
            var sel = rows[i].querySelector('select');
            if (sel) return { type: 'native', el: sel, id: sel.id || '' };
        }

        var labels = document.querySelectorAll('label.fmLi, td, th, label');
        for (var j = 0; j < labels.length; j++) {
            var txt = normalizeText(labels[j].textContent);
            if (txt !== 'Rückgabemenge' && txt.indexOf('Rückgabemenge') !== 0) continue;
            var parent = labels[j].closest('.fmRowSrd, tr') || labels[j].parentElement;
            for (var p = 0; p < 5 && parent; p++) {
                var ddl2 = parent.querySelector('.RadDropDownList');
                if (ddl2) return { type: 'telerik', el: ddl2, id: ddl2.id };
                var sel2 = parent.querySelector('select');
                if (sel2) return { type: 'native', el: sel2, id: sel2.id || '' };
                parent = parent.parentElement;
            }
        }

        var selects = document.querySelectorAll('select');
        var best = null;
        var bestScore = -1;
        for (var s = 0; s < selects.length; s++) {
            var score = 0;
            for (var o = 0; o < selects[s].options.length; o++) {
                if (parseNumericLabel(selects[s].options[o].textContent || selects[s].options[o].value) !== null) score++;
            }
            if (score > bestScore) {
                bestScore = score;
                best = selects[s];
            }
        }
        if (best) return { type: 'native', el: best, id: best.id || '' };
        return null;
    }

    async function waitForRueckgabeMengeControl(timeout) {
        var start = Date.now();
        while (Date.now() - start < (timeout || 10000)) {
            if (abortMission) return null;
            var ctrl = findRueckgabeMengeControl();
            if (ctrl) return ctrl;
            await sleep(180);
        }
        return null;
    }

    function getTelerikOptionsFromWidget(widget) {
        var out = [];
        if (!widget || !widget.get_items) return out;
        var items = widget.get_items();
        var count = items.get_count ? items.get_count() : 0;
        for (var i = 0; i < count; i++) {
            var item = items.getItem(i);
            var label = item.get_text ? item.get_text() : '';
            var num = parseNumericLabel(label);
            if (num === null) continue;
            out.push({ index: i, numeric: num, label: normalizeText(label) || String(num) });
        }
        return out;
    }

    async function openTelerikDropdown(root) {
        var inner = root.querySelector('.rddlInner, .rddlIcon') || root;
        forceClick(inner);
        await sleep(350);
    }

    function findVisibleTelerikPopup(root) {
        var byId = document.getElementById(root.id + '_DropDown');
        if (byId && isVisible(byId)) return byId;
        var popups = document.querySelectorAll('.rddlPopup, .RadDropDownListDropDown');
        for (var i = 0; i < popups.length; i++) {
            if (isVisible(popups[i])) return popups[i];
        }
        return null;
    }

    async function getTelerikOptionsFromDom(root) {
        await openTelerikDropdown(root);
        var popup = findVisibleTelerikPopup(root);
        if (!popup) return [];
        var nodes = popup.querySelectorAll('.rddlItem, li');
        var out = [];
        for (var i = 0; i < nodes.length; i++) {
            var label = normalizeText(nodes[i].textContent);
            var num = parseNumericLabel(label);
            if (num === null) continue;
            out.push({ index: i, numeric: num, label: label, domNode: nodes[i] });
        }
        return out;
    }

    async function getQuantityOptions(control) {
        if (!control) return [];
        if (control.type === 'native') {
            var out = [];
            var select = control.el;
            for (var i = 0; i < select.options.length; i++) {
                var num = parseNumericLabel(select.options[i].textContent || select.options[i].value);
                if (num === null) continue;
                out.push({ index: i, numeric: num, label: normalizeText(select.options[i].textContent) || String(num) });
            }
            return out;
        }
        var widget = getTelerikWidget(control.id);
        var opts = getTelerikOptionsFromWidget(widget);
        if (opts.length) return opts;
        return getTelerikOptionsFromDom(control.el);
    }

    async function setQuantityValue(control, option) {
        if (!control || !option) return false;
        if (control.type === 'native') {
            control.el.selectedIndex = option.index;
            control.el.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
        }
        var widget = getTelerikWidget(control.id);
        if (widget) {
            try {
                if (widget.selectIndex) widget.selectIndex(option.index);
                else if (widget.get_items) widget.get_items().getItem(option.index).select();
                return true;
            } catch (e) {}
        }
        if (option.domNode) {
            await openTelerikDropdown(control.el);
            forceClick(option.domNode);
            return true;
        }
        await openTelerikDropdown(control.el);
        var popup = findVisibleTelerikPopup(control.el);
        if (!popup) return false;
        var nodes = popup.querySelectorAll('.rddlItem, li');
        var picked = null;
        var pos = 0;
        for (var i = 0; i < nodes.length; i++) {
            var num = parseNumericLabel(nodes[i].textContent);
            if (num === null) continue;
            if (pos === option.index || num === option.numeric) {
                picked = nodes[i];
                break;
            }
            pos++;
        }
        if (!picked) return false;
        forceClick(picked);
        await sleep(200);
        return true;
    }

    function getCurrentQuantityDisplay(control) {
        if (!control) return null;
        if (control.type === 'native') {
            var opt = control.el.options[control.el.selectedIndex];
            return opt ? parseNumericLabel(opt.textContent || opt.value) : null;
        }
        var fake = control.el.querySelector('.rddlFakeInput');
        if (fake) return parseNumericLabel(fake.textContent);
        var widget = getTelerikWidget(control.id);
        if (widget && widget.get_text) return parseNumericLabel(widget.get_text());
        return null;
    }

    let lastBoundSearchInput = null;

    function pauseWatcher(ms) {
        pausePageWatcherUntil = Date.now() + (ms || 3000);
    }

    function shouldPauseWatcher() {
        if (searchSubmitted) return false;
        if (Date.now() < pausePageWatcherUntil) return true;
        if (!userEditingSearch || !searchInput) return false;
        if (document.activeElement === searchInput) return true;
        if (normalizeText(searchInput.value) && !searchSubmitted) return true;
        return false;
    }

    function canAutoFocusSearch() {
        if (abortMission || userEditingSearch || searchSubmitted) return false;
        var input = findSearchInput();
        if (!input || !isVisible(input)) return false;
        if (normalizeText(input.value)) return false;
        if (detectPage() !== 'search') return false;
        return true;
    }

    function scheduleSearchAutoFocus(force) {
        if (!force && searchFocusDone) return;
        if (searchFocusTimer) clearTimeout(searchFocusTimer);
        searchFocusTimer = setTimeout(function() {
            searchFocusTimer = null;
            if (!canAutoFocusSearch()) return;
            var input = findSearchInput();
            if (!input) return;
            searchInput = input;
            bindSearchInput(searchInput);
            if (document.activeElement === input) {
                searchFocusDone = true;
                return;
            }
            try {
                input.focus({ preventScroll: true });
            } catch (e) {
                try { input.focus(); } catch (e2) {}
            }
            searchFocusDone = true;
        }, force ? 120 : 320);
    }

    function handleSearchInput(e) {
        if (abortMission) return;
        searchInput = e.target;
        userEditingSearch = true;
        pauseWatcher(6000);
        lastInputChangeAt = Date.now();
        var value = normalizeText(searchInput.value);
        if (!value) searchSubmitted = false;
        if (searchSubmitted) return;
        scheduleScanSearch();
    }

    function handleSearchKeydown(e) {
        userEditingSearch = true;
        pauseWatcher(6000);
        if (e.key !== 'Enter') return;
        e.preventDefault();
        e.stopPropagation();
        searchInput = e.target;
        if (scanTimer) {
            clearTimeout(scanTimer);
            scanTimer = null;
        }
        triggerSearch(true);
    }

    function handleSearchPointerDown() {
        userEditingSearch = true;
        pauseWatcher(6000);
    }

    function bindSearchInput(el) {
        if (!el) return;
        if (el !== lastBoundSearchInput) {
            lastBoundSearchInput = el;
            el.addEventListener('input', handleSearchInput);
            el.addEventListener('keydown', handleSearchKeydown);
            el.addEventListener('pointerdown', handleSearchPointerDown);
        }
    }

    function isInputStable() {
        return Date.now() - lastInputChangeAt >= SCAN_IDLE_MS - 100;
    }

    function scheduleScanSearch() {
        if (scanTimer) clearTimeout(scanTimer);
        scanTimer = setTimeout(function() {
            scanTimer = null;
            if (!searchInput || searchSubmitted || abortMission) return;
            if (!isInputStable()) {
                scheduleScanSearch();
                return;
            }
            var len = normalizeText(searchInput.value).length;
            if (len >= MIN_SCAN_LEN) triggerSearch(false);
        }, SCAN_IDLE_MS);
    }

    async function triggerSearch(manual) {
        if (abortMission || !searchInput || searchSubmitted) return;
        if (!manual && !isInputStable()) return;
        var value = normalizeText(searchInput.value);
        if (!value) return;
        var minLen = manual ? MIN_SEARCH_LEN : MIN_SCAN_LEN;
        if (value.length < minLen) return;
        if (scanTimer) {
            clearTimeout(scanTimer);
            scanTimer = null;
        }
        await sleep(manual ? SUBMIT_DELAY_MS : SUBMIT_DELAY_MS + 150);
        if (abortMission || !searchInput || searchSubmitted) return;
        value = normalizeText(searchInput.value);
        if (!value || value.length < minLen) return;
        if (!manual && !isInputStable()) return;
        searchSubmitted = true;
        userEditingSearch = false;
        pauseWatcher(1500);
        sessionStorage.setItem(LAST_EAN_KEY, value);
        sessionStorage.removeItem(ORDER_CTX_KEY);
        sessionStorage.removeItem('wm_retoure_qty_hold');
        sessionStorage.setItem(ACTIVE_FLOW_KEY, '1');
        lastHandledStep = '';
        showHud('WM RETOURE LÄUFT', 'Suche Artikel für Nr.: ' + value, false);
        var btn = findClickableByText('Anzeigen', true);
        if (!btn) {
            searchSubmitted = false;
            showHud('WM RETOURE FEHLER', 'Anzeigen-Button nicht gefunden.', true);
            return;
        }
        forceClick(btn);
        setTimeout(function() {
            if (!shouldPauseWatcher()) tickFlow(true);
        }, 700);
    }

    function ensureSearchListeners(clearField) {
        var input = findSearchInput();
        if (!input) return false;
        searchInput = input;
        bindSearchInput(searchInput);
        if (clearField) {
            searchSubmitted = false;
            userEditingSearch = false;
            searchFocusDone = false;
            if (scanTimer) {
                clearTimeout(scanTimer);
                scanTimer = null;
            }
            searchInput.value = '';
            lastInputChangeAt = 0;
        }
        if (!searchListenersReady || clearField) {
            showHud('WM RETOURE BEREIT', 'EAN scannen oder Artikelnummer tippen — Enter bestätigt, Scanner startet automatisch nach Pause.', false, true);
        }
        searchListenersReady = true;
        if (clearField) searchFocusDone = false;
        scheduleSearchAutoFocus(!!clearField);
        return true;
    }

    function setupSearchListeners(clearField) {
        return ensureSearchListeners(clearField);
    }

    async function reloadForNextProduct() {
        if (abortMission) return;
        sessionStorage.removeItem(PENDING_NEUE_SUCHE_KEY);
        sessionStorage.setItem(ACTIVE_FLOW_KEY, '1');
        searchSubmitted = false;
        searchListenersReady = false;
        searchFocusDone = false;
        userEditingSearch = false;
        lastHandledStep = '';
        await sleep(400);
        location.reload();
    }

    async function handleReturnToSearch() {
        showHud('WM RETOURE BEREIT', 'Artikel hinzugefügt — Seite wird neu geladen...', false);
        await sleep(500);
        await reloadForNextProduct();
    }

    async function finishReturnToSearch() {
        sessionStorage.removeItem(PENDING_NEUE_SUCHE_KEY);
        sessionStorage.setItem(ACTIVE_FLOW_KEY, '1');
        lastHandledStep = '';
        searchSubmitted = false;
        searchListenersReady = false;
        searchFocusDone = false;
        userEditingSearch = false;
        if (ensureSearchListeners(true)) return;
        await sleep(400);
        ensureSearchListeners(true);
    }

    function isNoArticleFoundPage() {
        if (!pageHasText('Artikel auswählen')) return false;
        var text = pageBodyText().toLowerCase();
        if (text.indexOf('es wurden keine artikel gefunden') !== -1) return true;
        if (text.indexOf('keine artikel gefunden') !== -1) return true;
        if (findClickableByText('Auswählen', true)) return false;
        var tables = document.querySelectorAll('table');
        for (var i = 0; i < tables.length; i++) {
            var t = normalizeText(tables[i].innerText).toLowerCase();
            if (t.indexOf('katalog-nr') === -1) continue;
            if (t.indexOf('es wurden keine artikel gefunden') !== -1) return true;
        }
        return false;
    }

    function isNoOrdersFoundPage() {
        if (!pageHasText('Artikel zurückgeben')) return false;
        var text = pageBodyText().toLowerCase();
        if (text.indexOf('es wurden keine lieferscheine/rechnungen gefunden') !== -1) return true;
        if (text.indexOf('keine lieferscheine/rechnungen gefunden') !== -1) return true;
        if (text.indexOf('keine lieferscheine') !== -1 && text.indexOf('rechnungen gefunden') !== -1) return true;
        if (findNewestRueckgabeRow()) return false;
        var tables = document.querySelectorAll('table');
        for (var i = 0; i < tables.length; i++) {
            var t = normalizeText(tables[i].innerText).toLowerCase();
            if (t.indexOf('bestellt am') === -1) continue;
            if (t.indexOf('es wurden keine lieferscheine/rechnungen gefunden') !== -1) return true;
        }
        return false;
    }

    function showNoResultsWarning(message) {
        showHud('WM RETOURE WARNUNG', message, false, false, 'backToSearch');
    }

    async function clickZurueckZurSuche() {
        showHud('WM RETOURE LÄUFT', 'Zurück zur Suche...', false);
        var btn = findClickableByText('Neue Suche', true);
        if (!btn) btn = await waitForClickableByText('Neue Suche', true, 5000);
        if (!btn || abortMission) {
            showHud('WM RETOURE FEHLER', 'Neue Suche-Button nicht gefunden.', true);
            return false;
        }
        searchSubmitted = false;
        searchListenersReady = false;
        searchFocusDone = false;
        userEditingSearch = false;
        lastHandledStep = '';
        await sleep(250);
        forceClick(btn);
        setTimeout(function() {
            ensureSearchListeners(true);
        }, 700);
        return true;
    }

    async function handleArticlePage() {
        if (sessionStorage.getItem(PENDING_NEUE_SUCHE_KEY) === '1') {
            await handleReturnToSearch();
            return;
        }
        if (isNoArticleFoundPage()) {
            showNoResultsWarning('Es wurden keine Artikel gefunden.');
            return;
        }
        showHud('WM RETOURE LÄUFT', 'Artikel gefunden — wähle Auswählen...', false);
        var btn = await waitForClickableByText('Auswählen', true, 10000);
        if (!btn || abortMission) {
            lastHandledStep = '';
            if (isNoArticleFoundPage()) {
                showNoResultsWarning('Es wurden keine Artikel gefunden.');
                return;
            }
            showHud('WM RETOURE FEHLER', 'Auswählen-Button nicht gefunden.', true);
            return;
        }
        await sleep(250);
        rememberRow(btn);
        forceClick(btn);
    }

    async function handleOrdersPage() {
        if (sessionStorage.getItem(PENDING_NEUE_SUCHE_KEY) === '1') {
            await handleReturnToSearch();
            return;
        }
        if (isNoOrdersFoundPage()) {
            showNoResultsWarning('Es wurden keine Lieferscheine/Rechnungen gefunden.');
            return;
        }
        showHud('WM RETOURE LÄUFT', 'Suche neuesten Auftrag mit Rückgabe-Button...', false);
        var btn = await waitForNewestRueckgabeRow(10000);
        if (!btn || abortMission) {
            lastHandledStep = '';
            if (isNoOrdersFoundPage()) {
                showNoResultsWarning('Es wurden keine Lieferscheine/Rechnungen gefunden.');
                return;
            }
            showHud('WM RETOURE FEHLER', 'Kein Rückgabe-Button für aktuellen Auftrag gefunden.', true);
            return;
        }
        await sleep(300);
        rememberRow(btn);
        forceClick(btn);
    }

    function hideQuantityPicker() {
        var existing = document.getElementById('wm-retoure-qty-popup');
        if (existing) existing.remove();
    }

    function berlinNow() {
        var fmt = new Intl.DateTimeFormat('de-DE', {
            timeZone: 'Europe/Berlin',
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            hourCycle: 'h23'
        });
        var map = {};
        fmt.formatToParts(new Date()).forEach(function(part) { map[part.type] = part.value; });
        var hour = map.hour === '24' ? '00' : map.hour;
        return {
            date: map.day + '.' + map.month + '.' + map.year,
            time: hour + ':' + map.minute
        };
    }

    function sheetPayload(info) {
        var stamp = berlinNow();
        return {
            token: info.token || '',
            orderId: info.orderId || '',
            stockId: info.stockId || '',
            article: info.part || '',
            name: info.name || '',
            supplier: info.supplier || '',
            brand: info.brand || '',
            qty: info.qty || '',
            ean: info.ean || '',
            who: info.who || workerName(),
            date: info.date || stamp.date,
            time: info.time || stamp.time
        };
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
        return new Promise(function(resolve, reject) {
            GM_xmlhttpRequest({
                method: 'GET',
                url: url,
                anonymous: false,
                timeout: 30000,
                onload: function(res) {
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
                    else if (data && data.ok === true) reject(new Error('Web-App hat nichts geschrieben. Neue Version veröffentlichen.'));
                    else reject(new Error('Sheet Antwort ' + (res && res.status)));
                },
                onerror: function() { reject(new Error('Sheet nicht erreichbar')); },
                ontimeout: function() { reject(new Error('Sheet Timeout')); }
            });
        });
    }

    function postWmSheet(row) {
        var url = sheetUrl();
        if (!url) return Promise.resolve({ ok: false, error: 'Sheet-Link fehlt' });
        if (!row || (!row.article && !row.name)) return Promise.resolve({ ok: false, error: 'Teil fehlt' });
        if (row.token && sheetBusyToken === row.token) return Promise.resolve({ ok: false, error: 'läuft' });
        sheetBusyToken = row.token || 'x';
        var payload = JSON.stringify({ action: 'wmAppend', entries: [row] });
        var join = url.indexOf('?') === -1 ? '?' : '&';
        return sheetReq(url + join + 'payload=' + encodeURIComponent(payload), 0).then(function(data) {
            if (sheetBusyToken === row.token) sheetBusyToken = '';
            if (data.added || data.skipped) dropPending(row.token);
            markSheetNote(data.added ? 'Im Blatt WM' : (data.skipped ? 'Schon im Blatt' : 'Nichts ins Blatt geschrieben'), row.token);
            return data;
        }).catch(function(err) {
            if (sheetBusyToken === row.token) sheetBusyToken = '';
            var message = err && err.message ? err.message : 'Sheet fehlgeschlagen';
            if (message === 'läuft') return { ok: false, error: message };
            if (message === 'unbekannt') message = 'Web-App neu veröffentlichen';
            markSheetNote(message, row.token);
            return { ok: false, error: message };
        });
    }

    function loadPending() {
        try {
            var data = JSON.parse(sessionStorage.getItem(PENDING_SHEET_KEY) || '[]');
            if (Array.isArray(data)) return data.filter(function(item) { return item && item.token; });
            if (data && data.token) return [data];
        } catch (e) {}
        return [];
    }

    function savePending(list) {
        try {
            if (!list.length) sessionStorage.removeItem(PENDING_SHEET_KEY);
            else sessionStorage.setItem(PENDING_SHEET_KEY, JSON.stringify(list));
        } catch (e) {}
    }

    function rememberPending(row) {
        if (!row || !row.token) return;
        var list = loadPending().filter(function(item) { return item.token !== row.token; });
        list.push(row);
        savePending(list);
    }

    function dropPending(token) {
        if (!token) return;
        savePending(loadPending().filter(function(item) { return item.token !== token; }));
    }

    function retryPendingSheet() {
        if (!sheetUrl()) return;
        loadPending().forEach(function(row) { postWmSheet(row); });
    }

    function showQuantityPicker(control, options, onPick) {
        var main = document.getElementById('arasaka-batch-popup');
        if (main) main.remove();
        lastHudSignature = '';
        hideQuantityPicker();
        hudEnsureStyles();
        var info = readQuantityPageArticleInfo('');
        var selected = 0;
        var s;
        for (s = 0; s < options.length; s++) {
            if (options[s].numeric === 1) {
                selected = s;
                break;
            }
        }
        var buttonsHtml = options.map(function(opt, idx) {
            return '<button type="button" class="wm-retoure-qty-btn' + (idx === selected ? ' is-on' : '') + '" data-index="' + opt.index + '" data-numeric="' + opt.numeric + '">' + hudEscape(opt.label) + '</button>';
        }).join('');
        var metaBits = [];
        if (info.orderId) metaBits.push('Auftrag ' + info.orderId);
        if (info.brand) metaBits.push(info.brand);
        var popup = document.createElement('div');
        popup.id = 'wm-retoure-qty-popup';
        popup.setAttribute('data-tone', 'active');
        popup.innerHTML = ''
            + '<div class="arasaka-hud-head" id="wm-retoure-qty-drag">'
            + '  <div>'
            + '    <div class="arasaka-hud-eyebrow">WM Retoure · Teil hinzufügen</div>'
            + '    <div class="arasaka-hud-title">In die Rückgabe</div>'
            + '  </div>'
            + '  <div class="arasaka-hud-actions">'
            + '    <button type="button" class="arasaka-hud-btn" id="wm-retoure-qty-close">X</button>'
            + '  </div>'
            + '</div>'
            + '<div class="arasaka-hud-body">'
            + '  <div class="wm-part-card">'
            + '    <div class="wm-part-no">' + hudEscape(info.part || info.ean || 'Teil') + '</div>'
            + (info.name ? '<div class="wm-part-name">' + hudEscape(info.name) + '</div>' : '')
            + (metaBits.length ? '<div class="wm-part-meta">' + hudEscape(metaBits.join(' · ')) + '</div>' : '')
            + '  </div>'
            + '  <div class="wm-field">'
            + '    <span>Menge</span>'
            + '    <div class="wm-retoure-qty-grid">' + buttonsHtml + '</div>'
            + '  </div>'
            + '  <label class="wm-field">'
            + '    <span>Stock-ID <em>optional</em></span>'
            + '    <input id="wm-add-stock" type="text" maxlength="16" autocomplete="off" spellcheck="false" placeholder="z.B. JR97017">'
            + '  </label>'
            + '  <label class="wm-field">'
            + '    <span>Dein Name</span>'
            + '    <input id="wm-add-name" type="text" maxlength="40" autocomplete="off" placeholder="z.B. Semmel" value="' + hudEscape(workerName()) + '">'
            + '  </label>'
            + '  <div class="wm-form-note" id="wm-add-note">Ohne Stock-ID landet das Teil trotzdem im Blatt.</div>'
            + '  <button type="button" class="wm-go" id="wm-add-go">Hinzufügen</button>'
            + '  ' + hudFooterHtml()
            + '</div>';
        document.body.appendChild(popup);
        hudApplyPosition(popup);
        hudEnableDrag(popup, popup.querySelector('#wm-retoure-qty-drag'));
        var stockInput = popup.querySelector('#wm-add-stock');
        var nameInput = popup.querySelector('#wm-add-name');
        var note = popup.querySelector('#wm-add-note');
        var goBtn = popup.querySelector('#wm-add-go');
        function pickedOption() {
            var on = popup.querySelector('.wm-retoure-qty-btn.is-on') || popup.querySelector('.wm-retoure-qty-btn');
            if (!on) return options[0];
            var idx = parseInt(on.getAttribute('data-index'), 10);
            var numeric = parseInt(on.getAttribute('data-numeric'), 10);
            var i;
            for (i = 0; i < options.length; i++) {
                if (options[i].index === idx || options[i].numeric === numeric) return options[i];
            }
            return options[0];
        }
        popup.querySelector('#wm-retoure-qty-close').addEventListener('click', function() {
            popup.remove();
            sessionStorage.setItem('wm_retoure_qty_hold', '1');
            showHud('WM RETOURE', 'Teil ist noch nicht im Warenkorb.', false, false, 'reopenQty');
        });
        popup.querySelectorAll('.wm-retoure-qty-btn').forEach(function(btn) {
            btn.addEventListener('click', function() {
                popup.querySelectorAll('.wm-retoure-qty-btn').forEach(function(other) {
                    other.classList.remove('is-on');
                });
                btn.classList.add('is-on');
                stockInput.focus();
            });
        });
        function submitPick() {
            if (goBtn.disabled) return;
            var stock = cleanStockId(stockInput.value);
            if (stock === null) {
                note.textContent = 'Stock-ID passt nicht. Leer lassen oder z.B. JR97017.';
                note.className = 'wm-form-note is-bad';
                stockInput.focus();
                return;
            }
            var who = saveWorkerName(nameInput.value);
            if (!who) {
                note.textContent = 'Trag deinen Namen ein. Der gilt für alle Teile, die du hinzufügst.';
                note.className = 'wm-form-note is-bad';
                nameInput.focus();
                return;
            }
            goBtn.disabled = true;
            note.textContent = 'Wird hinzugefügt…';
            note.className = 'wm-form-note';
            popup.remove();
            onPick(pickedOption(), control, { stockId: stock, who: who });
        }
        goBtn.addEventListener('click', submitPick);
        stockInput.addEventListener('keydown', function(e) {
            if (e.key !== 'Enter') return;
            e.preventDefault();
            e.stopPropagation();
            submitPick();
        });
        nameInput.addEventListener('input', function() { saveWorkerName(nameInput.value); });
        nameInput.addEventListener('keydown', function(e) {
            if (e.key !== 'Enter') return;
            e.preventDefault();
            e.stopPropagation();
            submitPick();
        });
        setTimeout(function() {
            try { stockInput.focus(); } catch (e) {}
        }, 40);
    }

    async function submitQuantity(control, option, extra) {
        extra = extra || {};
        var current = getCurrentQuantityDisplay(control);
        if (current !== option.numeric) {
            var ok = await setQuantityValue(control, option);
            if (!ok) {
                lastHandledStep = '';
                showHud('WM RETOURE FEHLER', 'Rückgabemenge konnte nicht gesetzt werden.', true);
                return;
            }
        }
        var info = readQuantityPageArticleInfo(option.numeric);
        info.stockId = extra.stockId || '';
        info.who = extra.who || workerName();
        info.token = 'wm' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
        info.sheetNote = sheetUrl() ? 'Übertrage…' : 'Sheet-Link fehlt';
        saveLastAdded(info);
        var row = sheetPayload(info);
        rememberPending(row);
        sessionStorage.removeItem('wm_retoure_qty_hold');
        if (sheetUrl()) postWmSheet(row);
        showHud('WM RETOURE LÄUFT', 'Menge ' + option.numeric + (info.stockId ? ' · ' + info.stockId : '') + ' — übernehme das Teil...', false);
        await sleep(250);
        var submit = findClickableByText('Übernehmen', true);
        if (!submit) submit = await waitForClickableByText('Übernehmen', true, 8000);
        if (!submit || abortMission) {
            showHud('WM RETOURE FEHLER', 'Übernehmen-Button nicht gefunden.', true);
            return;
        }
        forceClick(submit);
        await sleep(1200);
        await handleReturnToSearch();
    }

    function findWarenkorbButton() {
        var nodes = document.querySelectorAll('input[type="button"], input[type="submit"], button, a');
        for (var i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (!isVisible(el)) continue;
            var label = normalizeText(el.value || el.alt || el.textContent);
            if (label.indexOf('Warenkorb') === 0) return el;
        }
        return null;
    }

    async function clickWarenkorb() {
        showHud('WM RETOURE LÄUFT', 'Öffne Warenkorb...', false);
        var btn = findWarenkorbButton();
        if (!btn) btn = await waitForClickableByText('Warenkorb', false, 8000);
        if (!btn || abortMission) {
            showHud('WM RETOURE FEHLER', 'Warenkorb-Button nicht gefunden.', true, true);
            return false;
        }
        await sleep(250);
        forceClick(btn);
        setTimeout(function() {
            tickFlow(true);
        }, 700);
        return true;
    }

    function findElementByTextMatch(matchFn) {
        var nodes = document.querySelectorAll('h1, h2, h3, h4, h5, label, legend, span, div, td, th, p, b, strong, font');
        var best = null;
        var bestLen = Infinity;
        for (var i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (!isVisible(el)) continue;
            if (el.closest('#arasaka-batch-popup, #wm-retoure-qty-popup, #wm-retoure-print-btn')) continue;
            var t = normalizeText(el.textContent);
            if (!matchFn(t)) continue;
            var len = t.length;
            if (len < bestLen) {
                bestLen = len;
                best = el;
            }
        }
        return best;
    }

    function findWarenkorbTitleElement() {
        return findElementByTextMatch(function(t) {
            return t.indexOf('Warenkorb - Rückgabeartikel') !== -1;
        });
    }

    function findKundenLsNrElement() {
        return findElementByTextMatch(function(t) {
            return t.indexOf('Kunden LS-Nr') !== -1;
        });
    }

    function findCartTable() {
        var tables = document.querySelectorAll('table');
        for (var i = 0; i < tables.length; i++) {
            var t = normalizeText(tables[i].innerText);
            if (t.indexOf('Katalog-Nr') !== -1 && t.indexOf('Bezeichnung') !== -1) return tables[i];
        }
        return null;
    }

    function findMailkopieField() {
        var labels = document.querySelectorAll('label, td, th, span, div, legend');
        for (var i = 0; i < labels.length; i++) {
            var txt = normalizeText(labels[i].textContent);
            if (txt.indexOf('Mailkopie senden an') === -1) continue;
            var scope = labels[i].closest('tr, .fmRowSrd, fieldset, form, table') || labels[i].parentElement;
            for (var p = 0; p < 6 && scope; p++) {
                var input = scope.querySelector('input[type="text"], input:not([type])');
                if (input && isVisible(input)) return input;
                scope = scope.parentElement;
            }
        }
        return null;
    }

    function findKundenLsNrField() {
        var label = findKundenLsNrElement();
        if (!label) return null;
        var scope = label.closest('tr, .fmRowSrd, fieldset, form, table, div') || label.parentElement;
        for (var p = 0; p < 6 && scope; p++) {
            var input = scope.querySelector('input[type="text"], input:not([type])');
            if (input && isVisible(input)) return input;
            scope = scope.parentElement;
        }
        return null;
    }

    function cloneTableForPrint(table) {
        var clone = table.cloneNode(true);
        clone.querySelectorAll('a').forEach(function(a) {
            var t = normalizeText(a.textContent);
            if (t === 'Bearbeiten' || t === 'Entfernen') {
                var cell = a.closest('td, th');
                if (cell) cell.remove();
            }
        });
        clone.querySelectorAll('input, button, select, script, style').forEach(function(el) {
            el.remove();
        });
        return clone.outerHTML;
    }

    function buildWarenkorbPrintHtml() {
        var titleEl = findWarenkorbTitleElement();
        if (!titleEl) return '';
        var html = '<h2 style="color:#2d7a2d;margin:0 0 12px;font-size:18px;">' + hudEscape(normalizeText(titleEl.textContent)) + '</h2>';
        var table = findCartTable();
        if (table) html += cloneTableForPrint(table);
        var mitteilung = findMitteilungField();
        if (mitteilung) {
            html += '<div style="margin:14px 0;">'
                + '<div style="font-weight:bold;margin-bottom:6px;">Ihre Mitteilung an den Großhändler</div>'
                + '<div style="border:1px solid #666;padding:8px;min-height:56px;white-space:pre-wrap;">'
                + hudEscape(mitteilung.value) + '</div></div>';
        }
        var mailField = findMailkopieField();
        if (mailField) {
            html += '<div style="margin:14px 0;">'
                + '<div style="font-weight:bold;margin-bottom:6px;">Mailkopie senden an</div>'
                + '<div style="border:1px solid #666;padding:4px 8px;display:inline-block;min-width:280px;">'
                + hudEscape(mailField.value) + '</div></div>';
        }
        html += '<hr style="border:none;border-top:1px solid #666;margin:18px 0;">';
        var lsLabel = findKundenLsNrElement();
        var lsInput = findKundenLsNrField();
        html += '<div style="margin:10px 0;">'
            + '<span style="font-weight:bold;">' + hudEscape(lsLabel ? normalizeText(lsLabel.textContent) : '*Kunden LS-Nr.') + '</span>';
        if (lsInput && normalizeText(lsInput.value)) {
            html += ' <span style="border:1px solid #666;padding:2px 8px;margin-left:8px;">' + hudEscape(lsInput.value) + '</span>';
        }
        html += '</div>';
        return html;
    }

    function getPrintFrame() {
        var frame = document.getElementById('wm-retoure-print-frame');
        if (!frame) {
            frame = document.createElement('iframe');
            frame.id = 'wm-retoure-print-frame';
            frame.setAttribute('title', 'WM Retoure Druck');
            frame.style.cssText = 'position:fixed;left:0;top:0;width:0;height:0;border:0;visibility:hidden;pointer-events:none;';
            document.body.appendChild(frame);
        }
        return frame;
    }

    function printWarenkorbSection() {
        var html = buildWarenkorbPrintHtml();
        if (!html) {
            showHud('WM RETOURE FEHLER', 'Druckbereich nicht gefunden.', true);
            return;
        }
        hideHud();
        var frame = getPrintFrame();
        var doc = frame.contentWindow || frame.contentDocument;
        if (!doc) {
            showHud('WM RETOURE FEHLER', 'Druckvorschau konnte nicht erstellt werden.', true);
            return;
        }
        if (doc.document) doc = doc.document;
        doc.open();
        doc.write('<!DOCTYPE html><html><head><meta charset="utf-8"><title>Warenkorb Rückgabeartikel</title><style>');
        doc.write('body{font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#000;margin:20px;padding:0;}');
        doc.write('table{border-collapse:collapse;width:100%;margin:10px 0;} th,td{border:1px solid #666;padding:5px 8px;text-align:left;vertical-align:top;} th{background:#eee;font-weight:bold;}');
        doc.write('@media print{body{margin:12mm;} @page{margin:12mm;}}</style></head><body>');
        doc.write(html);
        doc.write('</body></html>');
        doc.close();
        setTimeout(function() {
            try {
                var win = frame.contentWindow;
                win.focus();
                win.print();
            } catch (e) {
                showHud('WM RETOURE FEHLER', 'Drucken fehlgeschlagen.', true);
            }
        }, 350);
    }

    function setupPrintListeners() {}

    function hideHud() {
        hideQuantityPicker();
        var popup = document.getElementById('arasaka-batch-popup');
        if (popup) popup.remove();
        lastHudSignature = '';
    }

    function bindPrintButton(btn) {
        if (!btn || btn.__wmPrintBound) return;
        btn.__wmPrintBound = true;
        styleWmPrintButton(btn);
        btn.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            printWarenkorbSection();
        }, true);
    }

    function styleWmPrintButton(btn) {
        if (!btn) return;
        hudEnsureStyles();
        btn.classList.add('wm-retoure-print-wm');
        var ref = findClickableByText('Absenden', true) || findClickableByText('Anzeigen', true) || findWarenkorbButton();
        if (ref) {
            try {
                var cs = window.getComputedStyle(ref);
                btn.style.boxSizing = cs.boxSizing || 'border-box';
                btn.style.display = cs.display === 'inline' ? 'inline-block' : cs.display;
                btn.style.minHeight = cs.minHeight !== '0px' ? cs.minHeight : cs.height;
                btn.style.height = cs.height;
                btn.style.padding = cs.padding;
                btn.style.margin = cs.margin;
                btn.style.fontSize = cs.fontSize;
                btn.style.fontWeight = cs.fontWeight;
                btn.style.fontFamily = cs.fontFamily;
                btn.style.lineHeight = cs.lineHeight;
                btn.style.borderRadius = cs.borderRadius;
                btn.style.textAlign = 'center';
                btn.style.border = cs.border;
                btn.style.appearance = 'none';
                btn.style.webkitAppearance = 'none';
                btn.style.cursor = 'pointer';
            } catch (e) {}
        }
    }

    function findCartActionBar() {
        var absenden = findClickableByText('Absenden', true);
        if (absenden && absenden.parentElement) return absenden.parentElement;
        return null;
    }

    function ensurePrintButton() {
        setupPrintListeners();
        var existing = document.getElementById('wm-retoure-print-btn');
        if (existing) {
            styleWmPrintButton(existing);
            bindPrintButton(existing);
            return;
        }
        var host = findCartActionBar();
        if (!host) host = document.body;
        var btn = document.createElement('button');
        btn.id = 'wm-retoure-print-btn';
        btn.type = 'button';
        btn.textContent = 'Drucken';
        styleWmPrintButton(btn);
        bindPrintButton(btn);
        var absenden = findClickableByText('Absenden', true);
        if (absenden && absenden.parentElement === host) {
            host.insertBefore(btn, absenden);
        } else {
            host.appendChild(btn);
        }
        printButtonAdded = true;
    }

    function findMitteilungField() {
        var labels = document.querySelectorAll('label, td, th, span, div, legend');
        for (var i = 0; i < labels.length; i++) {
            var txt = normalizeText(labels[i].textContent);
            if (txt.indexOf('Ihre Mitteilung an den Großhändler') === -1) continue;
            var scope = labels[i].closest('tr, .fmRowSrd, fieldset, form, table') || labels[i].parentElement;
            for (var p = 0; p < 6 && scope; p++) {
                var ta = scope.querySelector('textarea');
                if (ta && isVisible(ta)) return ta;
                scope = scope.parentElement;
            }
        }
        var textareas = document.querySelectorAll('textarea');
        for (var j = 0; j < textareas.length; j++) {
            if (!isVisible(textareas[j])) continue;
            var box = textareas[j].closest('tr, div, td');
            if (box && normalizeText(box.textContent).indexOf('Mitteilung') !== -1) return textareas[j];
        }
        for (var k = 0; k < textareas.length; k++) {
            if (isVisible(textareas[k])) return textareas[k];
        }
        return null;
    }

    function fillMitteilungField() {
        var field = findMitteilungField();
        if (!field) return false;
        if (normalizeText(field.value) === MITTEILUNG_TEXT) return true;
        field.value = MITTEILUNG_TEXT;
        field.dispatchEvent(new Event('input', { bubbles: true }));
        field.dispatchEvent(new Event('change', { bubbles: true }));
        return true;
    }

    async function handleCartPage() {
        fillMitteilungField();
        ensurePrintButton();
        hideHud();
    }

    async function handleQuantityPage() {
        if (qtyOpening || document.getElementById('wm-retoure-qty-popup')) return;
        if (sessionStorage.getItem('wm_retoure_qty_hold') === '1') {
            showHud('WM RETOURE', 'Teil ist noch nicht im Warenkorb.', false, false, 'reopenQty');
            return;
        }
        qtyOpening = true;
        try {
            showHud('WM RETOURE LÄUFT', 'Rückgabemenge — warte auf Dropdown...', false);
            var control = await waitForRueckgabeMengeControl(10000);
            if (!control || abortMission) {
                lastHandledStep = '';
                showHud('WM RETOURE FEHLER', 'Rückgabemenge-Dropdown nicht gefunden.', true);
                return;
            }
            var options = await getQuantityOptions(control);
            if (!options.length) {
                lastHandledStep = '';
                showHud('WM RETOURE FEHLER', 'Keine gültige Rückgabemenge im Dropdown.', true);
                return;
            }
            showQuantityPicker(control, options, function(option, ctrl, extra) {
                submitQuantity(ctrl, option, extra || {});
            });
        } finally {
            qtyOpening = false;
        }
    }

    function shouldAutoRun(page) {
        if (page === 'search' || page === 'cart') return true;
        if (sessionStorage.getItem(ACTIVE_FLOW_KEY) === '1') return true;
        if (sessionStorage.getItem(PENDING_NEUE_SUCHE_KEY) === '1') return true;
        return false;
    }

    async function runPageFlow(force) {
        if (abortMission) return;
        if (isRunning && !force) return;
        var page = detectPage();
        if (page === 'unknown') return;
        var pendingReturn = sessionStorage.getItem(PENDING_NEUE_SUCHE_KEY) === '1';
        if (!shouldAutoRun(page) && page !== 'search' && !pendingReturn) return;
        var stepKey = page + '|' + (pendingReturn ? 'return' : '') + '|' + (sessionStorage.getItem(ACTIVE_FLOW_KEY) || '') + '|' + (searchSubmitted ? '1' : '');
        if (!force && lastHandledStep === stepKey) return;
        isRunning = true;
        try {
            await sleep(PAGE_WAIT_MS);
            if (abortMission) return;
            if (pendingReturn && page === 'search') {
                isRunning = false;
                await finishReturnToSearch();
                return;
            }
            if (page === 'search') {
                var currentInput = findSearchInput();
                if (currentInput) bindSearchInput(currentInput);
                if (shouldPauseWatcher()) {
                    isRunning = false;
                    return;
                }
                if (searchListenersReady && !force && !pendingReturn) {
                    isRunning = false;
                    return;
                }
                lastHandledStep = stepKey;
                isRunning = false;
                sessionStorage.setItem(ACTIVE_FLOW_KEY, '1');
                ensureSearchListeners(false);
                return;
            }
            if (page === 'article') {
                lastHandledStep = stepKey;
                await handleArticlePage();
                return;
            }
            if (page === 'orders') {
                lastHandledStep = stepKey;
                await handleOrdersPage();
                return;
            }
            if (page === 'quantity') {
                lastHandledStep = stepKey;
                await handleQuantityPage();
                return;
            }
            if (page === 'cart') {
                lastHandledStep = stepKey;
                await handleCartPage();
                return;
            }
        } finally {
            setTimeout(function() { isRunning = false; }, 800);
        }
    }

    function tickFlow(force) {
        if (!pageHasText('Warenrückgabe')) return;
        runPageFlow(!!force);
    }

    function startPageWatcher() {
        if (pageWatcherStarted) return;
        pageWatcherStarted = true;
        var debounceTimer = null;
        function scheduleTick() {
            if (shouldPauseWatcher()) return;
            if (debounceTimer) clearTimeout(debounceTimer);
            debounceTimer = setTimeout(function() {
                if (shouldPauseWatcher()) return;
                tickFlow(false);
            }, 250);
        }
        var observer = new MutationObserver(scheduleTick);
        observer.observe(document.documentElement, { childList: true, subtree: true });
        setInterval(function() {
            if (shouldPauseWatcher()) return;
            if (sessionStorage.getItem(ACTIVE_FLOW_KEY) === '1' || sessionStorage.getItem(PENDING_NEUE_SUCHE_KEY) === '1') {
                tickFlow(false);
            }
        }, 900);
    }

    function hudEscape(value) {
        return String(value == null ? '' : value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function hudFooterHtml() {
        return ''
            + '<div class="arasaka-hud-footer"><span>Bot v' + hudEscape(BOT_VERSION) + '</span><span>ESC Stop</span></div>'
            + '<div class="arasaka-hud-credit">by Arasaka</div>';
    }

    function lastAddedHudHtml() {
        var info = loadLastAdded();
        if (!info) return '';
        var lines = [];
        if (info.name) lines.push('<div class="arasaka-hud-last-name">' + hudEscape(info.name) + '</div>');
        var meta = [];
        if (info.part) meta.push(hudEscape(info.part));
        if (info.qty) meta.push('Menge ' + hudEscape(info.qty));
        if (info.stockId) meta.push('<span class="wm-stock">' + hudEscape(info.stockId) + '</span>');
        else meta.push('ohne Stock-ID');
        if (meta.length) lines.push('<div class="arasaka-hud-last-meta">' + meta.join(' · ') + '</div>');
        if (info.sheetNote) lines.push('<div class="wm-sheet-note' + sheetNoteClass(info.sheetNote) + '" id="wm-sheet-note">' + hudEscape(info.sheetNote) + '</div>');
        if (!lines.length) return '';
        return ''
            + '<div class="arasaka-hud-last">'
            + '  <div class="arasaka-hud-last-label">Zuletzt hinzugefügt</div>'
            + lines.join('')
            + '</div>';
    }

    function hudMessageHtml(message) {
        return hudEscape(message).replace(/\n/g, '<br>');
    }

    function hudTone(title, message, isEnd) {
        var text = String(title || '') + ' ' + String(message || '');
        var lower = text.toLowerCase();
        if (/warnung/.test(lower)) return 'warn';
        if (/fehler|error/.test(lower)) return 'error';
        if (/nicht gefunden/.test(lower)) return 'error';
        if (/stop|gestoppt/.test(lower) || isEnd) return 'done';
        if (/läuft|suche|wähle|menge|bereit/.test(lower)) return 'active';
        return 'neutral';
    }

    function hudToneLabel(tone) {
        if (tone === 'error') return 'FEHLER';
        if (tone === 'warn') return 'WARNUNG';
        if (tone === 'done') return 'FERTIG';
        if (tone === 'active') return 'LÄUFT';
        return 'INFO';
    }

    function hudClamp(value, min, max) {
        return Math.max(min, Math.min(max, value));
    }

    function hudPosition() {
        try {
            var raw = localStorage.getItem(HUD_POS_KEY);
            if (!raw) return null;
            var pos = JSON.parse(raw);
            if (typeof pos.left !== 'number' || typeof pos.top !== 'number') return null;
            return pos;
        } catch (e) {
            return null;
        }
    }

    function hudSavePosition(left, top) {
        try {
            localStorage.setItem(HUD_POS_KEY, JSON.stringify({ left: left, top: top }));
        } catch (e) {}
    }

    function hudApplyPosition(popup) {
        var pos = hudPosition();
        if (!pos) {
            popup.style.left = '50%';
            popup.style.top = '50%';
            popup.style.transform = 'translate(-50%, -50%)';
            return;
        }
        var maxLeft = Math.max(12, window.innerWidth - popup.offsetWidth - 12);
        var maxTop = Math.max(12, window.innerHeight - popup.offsetHeight - 12);
        popup.style.left = hudClamp(pos.left, 12, maxLeft) + 'px';
        popup.style.top = hudClamp(pos.top, 12, maxTop) + 'px';
        popup.style.transform = 'none';
    }

    function hudEnableDrag(popup, handle) {
        var state = null;
        function move(e) {
            if (!state) return;
            var maxLeft = Math.max(12, window.innerWidth - popup.offsetWidth - 12);
            var maxTop = Math.max(12, window.innerHeight - popup.offsetHeight - 12);
            var left = hudClamp(e.clientX - state.dx, 12, maxLeft);
            var top = hudClamp(e.clientY - state.dy, 12, maxTop);
            popup.style.left = left + 'px';
            popup.style.top = top + 'px';
            popup.style.transform = 'none';
            state.left = left;
            state.top = top;
        }
        function up() {
            if (state) hudSavePosition(state.left, state.top);
            state = null;
            document.removeEventListener('pointermove', move);
        }
        handle.addEventListener('pointerdown', function(e) {
            if (e.button !== 0 || e.target.closest('button, input, textarea, a, label, select')) return;
            var rect = popup.getBoundingClientRect();
            state = { dx: e.clientX - rect.left, dy: e.clientY - rect.top, left: rect.left, top: rect.top };
            popup.style.left = rect.left + 'px';
            popup.style.top = rect.top + 'px';
            popup.style.transform = 'none';
            document.addEventListener('pointermove', move);
            document.addEventListener('pointerup', up, { once: true });
            e.preventDefault();
        });
    }

    function hudEnsureStyles() {
        if (document.getElementById('wm-retoure-hud-style')) return;
        var style = document.createElement('style');
        style.id = 'wm-retoure-hud-style';
        style.textContent = ''
            + '#arasaka-batch-popup, #wm-retoure-qty-popup {'
            + '  position: fixed; z-index: 9999999; color: #e7f6ec;'
            + '  font-family: "Segoe UI", Arial, sans-serif; border-radius: 24px; overflow: hidden;'
            + '  background:'
            + '    radial-gradient(120% 80% at 0% 0%, rgba(72, 220, 130, 0.22), transparent 42%),'
            + '    radial-gradient(90% 70% at 100% 0%, rgba(12, 70, 40, 0.45), transparent 46%),'
            + '    linear-gradient(165deg, rgba(10, 28, 18, 0.78), rgba(5, 12, 9, 0.88));'
            + '  border: 1px solid rgba(150, 235, 180, 0.28);'
            + '  box-shadow: 0 28px 60px rgba(0, 0, 0, 0.45), inset 0 1px 0 rgba(255, 255, 255, 0.16);'
            + '  backdrop-filter: blur(22px) saturate(1.35); -webkit-backdrop-filter: blur(22px) saturate(1.35);'
            + '  user-select: none; }'
            + '#arasaka-batch-popup { width: min(420px, calc(100vw - 24px)); }'
            + '#wm-retoure-qty-popup { width: min(460px, calc(100vw - 24px)); }'
            + '#arasaka-batch-popup[data-tone="active"], #wm-retoure-qty-popup[data-tone="active"] { border-color: rgba(90, 220, 140, 0.55); }'
            + '#arasaka-batch-popup[data-tone="done"] { border-color: rgba(120, 235, 160, 0.7); }'
            + '#arasaka-batch-popup[data-tone="error"] { border-color: rgba(248, 113, 113, 0.72); }'
            + '#arasaka-batch-popup[data-tone="warn"] { border-color: rgba(251, 191, 36, 0.7); }'
            + '.arasaka-hud-head { display: flex; align-items: center; justify-content: space-between; gap: 14px;'
            + '  padding: 16px 16px 14px; cursor: move;'
            + '  background: linear-gradient(180deg, rgba(255,255,255,0.06), rgba(255,255,255,0.02));'
            + '  border-bottom: 1px solid rgba(160, 230, 185, 0.14); }'
            + '.arasaka-hud-eyebrow { color: #8ddeb0; font-size: 11px; font-weight: 800; letter-spacing: 1.4px; text-transform: uppercase; margin-bottom: 4px; }'
            + '.arasaka-hud-title { color: #f4fff7; font-size: 18px; font-weight: 800; letter-spacing: 0.2px; line-height: 1.15; }'
            + '.arasaka-hud-actions { display: flex; align-items: center; gap: 8px; flex-shrink: 0; }'
            + '.arasaka-hud-btn { border: 1px solid rgba(180, 240, 200, 0.18); border-radius: 12px;'
            + '  background: rgba(4, 16, 10, 0.45); color: #d7f5e3; min-width: 40px; height: 36px; padding: 0 12px;'
            + '  font-size: 13px; font-weight: 800; cursor: pointer; }'
            + '.arasaka-hud-body { padding: 16px; display: flex; flex-direction: column; gap: 12px; }'
            + '.arasaka-hud-main { display: flex; flex-direction: column; gap: 10px; align-items: stretch; }'
            + '.arasaka-hud-badge { align-self: flex-start; display: inline-flex; align-items: center; min-height: 28px; padding: 4px 10px;'
            + '  border-radius: 999px; background: rgba(8, 24, 14, 0.45); border: 1px solid rgba(110, 220, 150, 0.35); color: #b6f5cf;'
            + '  font-size: 11px; font-weight: 800; letter-spacing: 1.1px; }'
            + '#arasaka-batch-popup[data-tone="error"] .arasaka-hud-badge { border-color: rgba(248, 113, 113, 0.55); color: #fecaca; }'
            + '#arasaka-batch-popup[data-tone="warn"] .arasaka-hud-badge { border-color: rgba(251, 191, 36, 0.55); color: #fde68a; }'
            + '.arasaka-hud-message { color: #e7f6ec; font-size: 15px; font-weight: 650; line-height: 1.45; word-break: break-word; user-select: text; }'
            + '.arasaka-hud-footer { display: flex; justify-content: space-between; gap: 10px; padding-top: 10px;'
            + '  border-top: 1px solid rgba(160, 230, 185, 0.12); color: #8eae98; font-size: 11px; font-weight: 700; }'
            + '.arasaka-hud-credit { text-align: right; color: #6f907c; font-size: 10px; font-weight: 600; }'
            + '.arasaka-hud-last { padding: 12px; border-radius: 16px; background: rgba(8, 32, 18, 0.42); border: 1px solid rgba(120, 220, 155, 0.22); }'
            + '.arasaka-hud-last-label { color: #8ddeb0; font-size: 10px; font-weight: 800; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 6px; }'
            + '.arasaka-hud-last-name { color: #f4fff7; font-size: 15px; font-weight: 750; line-height: 1.3; }'
            + '.arasaka-hud-last-meta { color: #c9ead6; font-size: 13px; font-weight: 700; margin-top: 4px; }'
            + '.wm-stock { color: #86f0b0; font-weight: 800; }'
            + '.wm-sheet-note { margin-top: 8px; font-size: 12px; font-weight: 750; color: #d7f5e3; }'
            + '.wm-sheet-note.is-ok { color: #86f0b0; }'
            + '.wm-sheet-note.is-bad { color: #fecaca; }'
            + '.wm-part-card { padding: 12px 14px; border-radius: 16px; background: rgba(8, 32, 18, 0.42); border: 1px solid rgba(120, 220, 155, 0.22); }'
            + '.wm-part-no { color: #86f0b0; font-size: 18px; font-weight: 800; letter-spacing: 0.2px; }'
            + '.wm-part-name { color: #f4fff7; font-size: 14px; font-weight: 700; margin-top: 4px; line-height: 1.35; }'
            + '.wm-part-meta { color: #8eae98; font-size: 12px; font-weight: 700; margin-top: 4px; }'
            + '.wm-field { display: flex; flex-direction: column; gap: 6px; }'
            + '.wm-field > span { color: #8ddeb0; font-size: 10px; font-weight: 800; letter-spacing: 1px; text-transform: uppercase; }'
            + '.wm-field > span em { font-style: normal; letter-spacing: 0; text-transform: none; color: #c9ead6; font-weight: 700; }'
            + '.wm-field input { width: 100%; box-sizing: border-box; border-radius: 12px; border: 1px solid rgba(150, 230, 180, 0.22);'
            + '  background: rgba(3, 14, 9, 0.55); color: #f4fff7; font: 700 15px/1.2 "Segoe UI", Arial, sans-serif;'
            + '  padding: 11px 12px; outline: none; user-select: text; -webkit-user-select: text; }'
            + '.wm-field input:focus { border-color: rgba(110, 220, 150, 0.7); box-shadow: 0 0 0 3px rgba(70, 200, 120, 0.16); }'
            + '.wm-field input::placeholder { color: #6f907c; font-weight: 600; }'
            + '.wm-form-note { color: #8eae98; font-size: 12px; font-weight: 650; line-height: 1.4; }'
            + '.wm-form-note.is-bad { color: #fecaca; }'
            + '.wm-go, .wm-retoure-hud-action { width: 100%; min-height: 46px; border-radius: 14px; cursor: pointer;'
            + '  border: 1px solid rgba(170, 245, 195, 0.35); color: #05210f; font-size: 14px; font-weight: 800;'
            + '  background: linear-gradient(180deg, #8cf0b4, #3dbe72); }'
            + '.wm-go:disabled { opacity: 0.6; cursor: default; }'
            + '.wm-retoure-hud-action-warn { color: #fde68a; border-color: rgba(251, 191, 36, 0.45);'
            + '  background: rgba(120, 70, 10, 0.28); }'
            + '.wm-retoure-qty-grid { display: flex; flex-wrap: wrap; gap: 8px; }'
            + '.wm-retoure-qty-btn { min-width: 64px; min-height: 48px; padding: 8px 14px; border-radius: 14px; cursor: pointer;'
            + '  border: 1px solid rgba(140, 220, 170, 0.28); background: rgba(4, 16, 10, 0.45); color: #d7f5e3;'
            + '  font-size: 20px; font-weight: 800; }'
            + '.wm-retoure-qty-btn.is-on, .wm-retoure-qty-btn:hover { color: #05210f; background: linear-gradient(180deg, #b6f7cf, #5ed890); border-color: transparent; }'
            + '.wm-identity { display: flex; flex-direction: column; gap: 8px; padding-top: 4px; }'
            + '.wm-sheet-row { display: flex; align-items: center; justify-content: space-between; gap: 8px; }'
            + '.wm-sheet-toggle { border: 1px solid rgba(150, 230, 180, 0.22); background: rgba(4, 16, 10, 0.35); color: #d7f5e3;'
            + '  border-radius: 999px; padding: 6px 12px; font-size: 12px; font-weight: 800; cursor: pointer; }'
            + '.wm-sheet-state { font-size: 12px; font-weight: 750; }'
            + '.wm-sheet-state.is-ok { color: #86f0b0; }'
            + '.wm-sheet-state.is-bad { color: #fecaca; }'
            + '.wm-sheet-hint { color: #6f907c; font-size: 11px; font-weight: 600; line-height: 1.35; }'
            + '.wm-retoure-hud-actions { display: flex; }'
            + '#wm-retoure-print-btn, button.wm-retoure-print-wm {'
            + '  background: #18884a !important; background-color: #18884a !important; color: #fff !important;'
            + '  border: none !important; box-shadow: none !important; cursor: pointer; text-decoration: none !important;'
            + '  font-weight: 700 !important; margin-right: 12px; }'
            + '#wm-retoure-print-btn:hover, button.wm-retoure-print-wm:hover {'
            + '  background: #146e3c !important; background-color: #146e3c !important; color: #fff !important; }';
        document.head.appendChild(style);
    }

    function bindDynamic(popup) {
        var cartBtn = popup.querySelector('#wm-retoure-go-cart');
        if (cartBtn) cartBtn.addEventListener('click', function() { clickWarenkorb(); });
        var backBtn = popup.querySelector('#wm-retoure-back-search');
        if (backBtn) backBtn.addEventListener('click', function() { clickZurueckZurSuche(); });
        var reopen = popup.querySelector('#wm-retoure-reopen-qty');
        if (reopen) {
            reopen.addEventListener('click', function() {
                sessionStorage.removeItem('wm_retoure_qty_hold');
                handleQuantityPage();
            });
        }
    }

    function paintSheetState(extra) {
        var state = document.getElementById('wm-sheet-state');
        if (!state) return;
        if (extra) {
            state.textContent = extra;
            state.className = 'wm-sheet-state is-bad';
            return;
        }
        var ok = !!sheetUrl();
        state.textContent = ok ? 'verbunden' : 'Link fehlt';
        state.className = 'wm-sheet-state' + (ok ? ' is-ok' : ' is-bad');
    }

    function bindIdentity(popup) {
        var name = popup.querySelector('#wm-worker-name');
        var url = popup.querySelector('#wm-sheet-url');
        var box = popup.querySelector('#wm-sheet-box');
        var toggle = popup.querySelector('#wm-sheet-toggle');
        if (name) {
            name.addEventListener('input', function() { saveWorkerName(name.value); });
            name.addEventListener('keydown', function(e) {
                if (e.key === 'Enter') e.preventDefault();
            });
        }
        if (url) {
            url.addEventListener('change', function() {
                if (!saveSheetUrl(url.value)) {
                    paintSheetState('Link ungültig');
                    return;
                }
                paintSheetState();
                if (sheetUrl()) {
                    if (box) box.hidden = true;
                    retryPendingSheet();
                }
            });
            url.addEventListener('keydown', function(e) {
                if (e.key !== 'Enter') return;
                e.preventDefault();
                url.blur();
            });
        }
        if (toggle && box) {
            toggle.addEventListener('click', function() {
                box.hidden = !box.hidden;
                if (!box.hidden && url) url.focus();
            });
        }
        paintSheetState();
    }

    function bindHudChrome(popup) {
        if (popup.getAttribute('data-bound') === '1') return;
        popup.setAttribute('data-bound', '1');
        var close = popup.querySelector('#wm-retoure-hud-close');
        if (close) {
            close.addEventListener('click', function() {
                popup.remove();
                lastHudSignature = '';
            });
        }
        hudEnableDrag(popup, popup.querySelector('#wm-retoure-hud-drag'));
        bindIdentity(popup);
    }

    function shouldShowCartButton(explicit) {
        if (explicit != null) return !!explicit;
        return detectPage() === 'search' && sessionStorage.getItem(ACTIVE_FLOW_KEY) === '1';
    }

    function buildHudActionHtml(showCart, hudAction) {
        if (hudAction === 'backToSearch') {
            return '<div class="wm-retoure-hud-actions"><button type="button" class="wm-retoure-hud-action wm-retoure-hud-action-warn" id="wm-retoure-back-search">Zurück zur Suche</button></div>';
        }
        if (hudAction === 'reopenQty') {
            return '<div class="wm-retoure-hud-actions"><button type="button" class="wm-retoure-hud-action" id="wm-retoure-reopen-qty">Menge wählen</button></div>';
        }
        if (showCart) {
            return '<div class="wm-retoure-hud-actions"><button type="button" class="wm-retoure-hud-action" id="wm-retoure-go-cart">Zum Warenkorb</button></div>';
        }
        return '';
    }

    function hudDynamicHtml(tone, message, lastHtml, actionHtml) {
        return ''
            + '<div class="arasaka-hud-main">'
            + '  <div class="arasaka-hud-badge">' + hudEscape(hudToneLabel(tone)) + '</div>'
            + '  <div class="arasaka-hud-message">' + hudMessageHtml(message) + '</div>'
            + lastHtml
            + '</div>'
            + actionHtml;
    }

    function identityHtml() {
        var url = sheetUrl();
        return ''
            + '<div class="wm-identity" id="wm-retoure-identity">'
            + '  <label class="wm-field"><span>Dein Name</span>'
            + '    <input id="wm-worker-name" type="text" maxlength="40" autocomplete="off" placeholder="z.B. Semmel" value="' + hudEscape(workerName()) + '">'
            + '  </label>'
            + '  <div class="wm-sheet-row">'
            + '    <button type="button" class="wm-sheet-toggle" id="wm-sheet-toggle">Sheet</button>'
            + '    <span class="wm-sheet-state" id="wm-sheet-state"></span>'
            + '  </div>'
            + '  <div id="wm-sheet-box"' + (url ? ' hidden' : '') + '>'
            + '    <label class="wm-field"><span>Web-App</span>'
            + '      <input id="wm-sheet-url" type="url" autocomplete="off" spellcheck="false" placeholder="https://script.google.com/.../exec" value="' + hudEscape(url) + '">'
            + '    </label>'
            + '    <div class="wm-sheet-hint">Gleicher Link wie bei Alfah. Nach dem Update die Web-App neu veröffentlichen.</div>'
            + '  </div>'
            + '</div>';
    }

    function showHud(title, message, isEnd, showCartButton, hudAction) {
        hudEnsureStyles();
        var showCart = shouldShowCartButton(showCartButton);
        var lastHtml = lastAddedHudHtml();
        var signature = title + '|' + message + '|' + !!isEnd + '|' + !!showCart + '|' + (hudAction || '') + '|' + lastHtml;
        var existing = document.getElementById('arasaka-batch-popup');
        if (existing && !existing.querySelector('#wm-retoure-dynamic')) {
            existing.remove();
            existing = null;
        }
        if (existing && signature === lastHudSignature) return;
        lastHudSignature = signature;
        var tone = hudTone(title, message, isEnd);
        var actionHtml = buildHudActionHtml(showCart, hudAction);
        var dynamic = hudDynamicHtml(tone, message, lastHtml, actionHtml);
        if (existing) {
            existing.setAttribute('data-tone', tone);
            var eye = existing.querySelector('.arasaka-hud-eyebrow');
            var heading = existing.querySelector('.arasaka-hud-title');
            if (eye) eye.textContent = 'WM Retoure · ' + hudToneLabel(tone);
            if (heading) heading.textContent = title;
            var slot = existing.querySelector('#wm-retoure-dynamic');
            if (slot) slot.innerHTML = dynamic;
            bindDynamic(existing);
            return;
        }
        var popup = document.createElement('div');
        popup.id = 'arasaka-batch-popup';
        popup.setAttribute('data-tone', tone);
        popup.innerHTML = ''
            + '<div class="arasaka-hud-head" id="wm-retoure-hud-drag">'
            + '  <div>'
            + '    <div class="arasaka-hud-eyebrow">WM Retoure · ' + hudEscape(hudToneLabel(tone)) + '</div>'
            + '    <div class="arasaka-hud-title">' + hudEscape(title) + '</div>'
            + '  </div>'
            + '  <div class="arasaka-hud-actions">'
            + '    <button type="button" class="arasaka-hud-btn" id="wm-retoure-hud-close">X</button>'
            + '  </div>'
            + '</div>'
            + '<div class="arasaka-hud-body">'
            + '  <div id="wm-retoure-dynamic">' + dynamic + '</div>'
            + identityHtml()
            + '  ' + hudFooterHtml()
            + '</div>';
        document.body.appendChild(popup);
        hudApplyPosition(popup);
        bindHudChrome(popup);
        bindDynamic(popup);
    }

    function boot() {
        if (!pageHasText('Warenrückgabe')) return;
        retryPendingSheet();
        var page = detectPage();
        if (page === 'article' && sessionStorage.getItem(ACTIVE_FLOW_KEY) !== '1' && findClickableByText('Auswählen', true)) {
            sessionStorage.setItem(ACTIVE_FLOW_KEY, '1');
        }
        if (page === 'quantity' && sessionStorage.getItem(ACTIVE_FLOW_KEY) !== '1' && findRueckgabeMengeControl()) {
            sessionStorage.setItem(ACTIVE_FLOW_KEY, '1');
        }
        startPageWatcher();
        scheduleSearchAutoFocus(true);
        if (sessionStorage.getItem(ACTIVE_FLOW_KEY) === '1' && page === 'search') {
            showHud('WM RETOURE BEREIT', 'EAN scannen oder Artikelnummer tippen — Enter bestätigt, Scanner startet automatisch nach Pause.', false, true);
        }
        tickFlow(true);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', boot);
    } else {
        boot();
    }
})();
