// ==UserScript==
// @name         WM Warenrückgabe Retoure Bot
// @namespace    http://tampermonkey.net/
// @version      1.8
// @description  Automatisiert WM Warenrückgabe per EAN-Scan
// @author       ARASAKA
// @match        *://*.customer-de.wm.de/*
// @match        *://customer-de.wm.de/*
// @updateURL    https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/Warenr%C3%BCckgabe/wm-retoure-automation.user.js
// @downloadURL  https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/Warenr%C3%BCckgabe/wm-retoure-automation.user.js
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    const BOT_VERSION = '1.8';
    const HUD_POS_KEY = 'wm_retoure_hud_position';
    const PENDING_NEUE_SUCHE_KEY = 'wm_retoure_pending_neue_suche';
    const ACTIVE_FLOW_KEY = 'wm_retoure_active_flow';
    const LAST_EAN_KEY = 'wm_retoure_last_ean';
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

    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            abortMission = true;
            isRunning = false;
            if (scanTimer) clearTimeout(scanTimer);
            sessionStorage.removeItem(PENDING_NEUE_SUCHE_KEY);
            sessionStorage.removeItem(ACTIVE_FLOW_KEY);
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

    function forceClick(el) {
        if (!el || abortMission) return;
        el.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
        el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
        el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        try { el.click(); } catch (e) {}
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
        var overlays = clone.querySelectorAll('#arasaka-batch-popup, #wm-retoure-qty-popup');
        for (var i = 0; i < overlays.length; i++) overlays[i].remove();
        return clone.innerText;
    }

    function pageHasText(text) {
        return pageBodyText().indexOf(text) !== -1;
    }

    function detectPage() {
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
            if (typeof window.$find === 'function') return window.$find(controlId);
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
            showHud('WM RETOURE BEREIT', 'EAN scannen oder Artikelnummer tippen — Enter bestätigt, Scanner startet automatisch nach Pause.', false);
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

    async function handleArticlePage() {
        if (sessionStorage.getItem(PENDING_NEUE_SUCHE_KEY) === '1') {
            await handleReturnToSearch();
            return;
        }
        showHud('WM RETOURE LÄUFT', 'Artikel gefunden — wähle Auswählen...', false);
        var btn = await waitForClickableByText('Auswählen', true, 10000);
        if (!btn || abortMission) {
            lastHandledStep = '';
            showHud('WM RETOURE FEHLER', 'Auswählen-Button nicht gefunden.', true);
            return;
        }
        await sleep(250);
        forceClick(btn);
    }

    async function handleOrdersPage() {
        if (sessionStorage.getItem(PENDING_NEUE_SUCHE_KEY) === '1') {
            await handleReturnToSearch();
            return;
        }
        showHud('WM RETOURE LÄUFT', 'Suche neuesten Auftrag mit Rückgabe-Button...', false);
        var btn = await waitForNewestRueckgabeRow(10000);
        if (!btn || abortMission) {
            lastHandledStep = '';
            showHud('WM RETOURE FEHLER', 'Kein Rückgabe-Button für aktuellen Auftrag gefunden.', true);
            return;
        }
        await sleep(300);
        forceClick(btn);
    }

    function hideQuantityPicker() {
        var existing = document.getElementById('wm-retoure-qty-popup');
        if (existing) existing.remove();
    }

    function showQuantityPicker(control, options, onPick) {
        hideQuantityPicker();
        hudEnsureStyles();
        var popup = document.createElement('div');
        popup.id = 'wm-retoure-qty-popup';
        popup.setAttribute('data-tone', 'active');
        var buttonsHtml = options.map(function(opt) {
            return '<button type="button" class="wm-retoure-qty-btn" data-index="' + opt.index + '" data-numeric="' + opt.numeric + '">' + hudEscape(opt.label) + '</button>';
        }).join('');
        popup.innerHTML = ''
            + '<div class="arasaka-hud-head" id="wm-retoure-qty-drag">'
            + '  <div>'
            + '    <div class="arasaka-hud-eyebrow">WM Retoure · MENGE WÄHLEN</div>'
            + '    <div class="arasaka-hud-title">Rückgabemenge</div>'
            + '  </div>'
            + '  <div class="arasaka-hud-actions">'
            + '    <button type="button" class="arasaka-hud-btn" id="wm-retoure-qty-close">X</button>'
            + '  </div>'
            + '</div>'
            + '<div class="arasaka-hud-body">'
            + '  <div class="arasaka-hud-main">'
            + '    <div class="arasaka-hud-badge">MENGE</div>'
            + '    <div class="arasaka-hud-message">Wie viele Stück zurückgeben?</div>'
            + '  </div>'
            + '  <div class="wm-retoure-qty-grid">' + buttonsHtml + '</div>'
            + '  <div class="arasaka-hud-footer"><span>Bot v' + hudEscape(BOT_VERSION) + '</span><span>ESC Stop</span></div>'
            + '</div>';
        document.body.appendChild(popup);
        hudApplyPosition(popup);
        hudEnableDrag(popup, popup.querySelector('#wm-retoure-qty-drag'));
        popup.querySelector('#wm-retoure-qty-close').addEventListener('click', function() {
            popup.remove();
        });
        popup.querySelectorAll('.wm-retoure-qty-btn').forEach(function(btn) {
            btn.addEventListener('click', function() {
                var idx = parseInt(btn.getAttribute('data-index'), 10);
                var numeric = parseInt(btn.getAttribute('data-numeric'), 10);
                var picked = null;
                for (var i = 0; i < options.length; i++) {
                    if (options[i].index === idx || options[i].numeric === numeric) {
                        picked = options[i];
                        break;
                    }
                }
                popup.remove();
                onPick(picked || options[0], control);
            });
        });
    }

    async function submitQuantity(control, option) {
        var current = getCurrentQuantityDisplay(control);
        if (current !== option.numeric) {
            var ok = await setQuantityValue(control, option);
            if (!ok) {
                lastHandledStep = '';
                showHud('WM RETOURE FEHLER', 'Rückgabemenge konnte nicht gesetzt werden.', true);
                return;
            }
        }
        showHud('WM RETOURE LÄUFT', 'Menge ' + option.numeric + ' — klicke Übernehmen...', false);
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

    async function handleQuantityPage() {
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
        if (options.length === 1) {
            await submitQuantity(control, options[0]);
            return;
        }
        showQuantityPicker(control, options, function(option, ctrl) {
            submitQuantity(ctrl, option);
        });
    }

    function shouldAutoRun(page) {
        if (page === 'search') return true;
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

    function hudMessageHtml(message) {
        return hudEscape(message).replace(/\n/g, '<br>');
    }

    function hudTone(title, message, isEnd) {
        var text = String(title || '') + ' ' + String(message || '');
        var lower = text.toLowerCase();
        if (/fehler|error|nicht gefunden/.test(lower)) return 'error';
        if (/warnung|wartet/.test(lower)) return 'warn';
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
            if (e.button !== 0 || e.target.closest('button')) return;
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
            + '  position: fixed; z-index: 9999999; width: min(560px, calc(100vw - 32px)); color: #c9d1d9;'
            + '  font-family: "Segoe UI", Arial, sans-serif; border-radius: 22px; overflow: hidden;'
            + '  background: radial-gradient(circle at top left, rgba(61, 158, 220, 0.14), transparent 34%),'
            + '    radial-gradient(circle at top right, rgba(86, 211, 100, 0.1), transparent 32%),'
            + '    linear-gradient(180deg, rgba(21, 27, 35, 0.98) 0%, rgba(14, 19, 26, 0.98) 100%);'
            + '  border: 1px solid #2d3642; box-shadow: 0 24px 56px rgba(0, 0, 0, 0.52);'
            + '  backdrop-filter: blur(12px); -webkit-backdrop-filter: blur(12px); user-select: none;'
            + '}'
            + '#arasaka-batch-popup[data-tone="active"], #wm-retoure-qty-popup[data-tone="active"] { border-color: rgba(61, 158, 220, 0.52); }'
            + '#arasaka-batch-popup[data-tone="done"] { border-color: rgba(86, 211, 100, 0.55); }'
            + '#arasaka-batch-popup[data-tone="error"] { border-color: rgba(248, 81, 73, 0.68); }'
            + '.arasaka-hud-head { display: flex; align-items: center; justify-content: space-between; gap: 18px;'
            + '  padding: 18px 20px 16px; background: linear-gradient(180deg, rgba(22, 27, 34, 0.99) 0%, rgba(18, 23, 30, 0.99) 100%);'
            + '  border-bottom: 1px solid #242d39; cursor: move; }'
            + '.arasaka-hud-eyebrow { color: #8b949e; font-size: 12px; font-weight: 800; letter-spacing: 1.6px; text-transform: uppercase; margin-bottom: 4px; }'
            + '.arasaka-hud-title { color: #f0f6fc; font-size: 20px; font-weight: 900; letter-spacing: 0.5px; line-height: 1.15; text-transform: uppercase; }'
            + '.arasaka-hud-actions { display: flex; align-items: center; gap: 8px; flex-shrink: 0; }'
            + '.arasaka-hud-btn { border: 1px solid rgba(255,255,255,0.12); border-radius: 14px; background: rgba(13, 17, 23, 0.72);'
            + '  color: #c9d1d9; min-width: 42px; height: 40px; padding: 0 14px; font-size: 13px; font-weight: 900; cursor: pointer; }'
            + '.arasaka-hud-body { padding: 22px; }'
            + '.arasaka-hud-main { display: flex; flex-direction: column; gap: 16px; align-items: start; }'
            + '.arasaka-hud-badge { display: inline-flex; align-items: center; justify-content: center; min-height: 40px; padding: 9px 16px;'
            + '  border-radius: 999px; background: rgba(13, 17, 23, 0.7); border: 2px solid rgba(61, 158, 220, 0.38); color: #3FA0DB;'
            + '  font-size: 13px; font-weight: 900; letter-spacing: 1.2px; text-transform: uppercase; }'
            + '.arasaka-hud-message { color: #d7dee8; font-size: 16px; font-weight: 700; line-height: 1.52; white-space: normal; word-break: break-word; user-select: text; }'
            + '.arasaka-hud-footer { display: flex; justify-content: space-between; gap: 10px; margin-top: 16px; padding-top: 12px;'
            + '  border-top: 1px solid rgba(255,255,255,0.07); color: #8b949e; font-size: 12px; font-weight: 800; }'
            + '.wm-retoure-qty-grid { display: flex; flex-wrap: wrap; gap: 12px; margin-top: 18px; }'
            + '.wm-retoure-qty-btn { min-width: 72px; min-height: 56px; padding: 10px 18px; border-radius: 16px; border: 2px solid rgba(86, 211, 100, 0.42);'
            + '  background: rgba(13, 17, 23, 0.82); color: #56d364; font-size: 24px; font-weight: 900; cursor: pointer; transition: transform 0.14s, background 0.14s; }'
            + '.wm-retoure-qty-btn:hover { transform: translateY(-2px); background: rgba(86, 211, 100, 0.14); }';
        document.head.appendChild(style);
    }

    function showHud(title, message, isEnd) {
        hudEnsureStyles();
        var signature = title + '|' + message + '|' + !!isEnd;
        var existing = document.getElementById('arasaka-batch-popup');
        if (existing && signature === lastHudSignature) return;
        lastHudSignature = signature;
        var tone = hudTone(title, message, isEnd);
        var html = ''
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
            + '  <div class="arasaka-hud-main">'
            + '    <div class="arasaka-hud-badge">' + hudEscape(hudToneLabel(tone)) + '</div>'
            + '    <div class="arasaka-hud-message">' + hudMessageHtml(message) + '</div>'
            + '  </div>'
            + '  <div class="arasaka-hud-footer"><span>Bot v' + hudEscape(BOT_VERSION) + '</span><span>ESC Stop</span></div>'
            + '</div>';
        if (existing) {
            existing.setAttribute('data-tone', tone);
            existing.innerHTML = html;
            existing.querySelector('#wm-retoure-hud-close').addEventListener('click', function() {
                existing.remove();
                lastHudSignature = '';
            });
            hudEnableDrag(existing, existing.querySelector('#wm-retoure-hud-drag'));
            return;
        }
        var popup = document.createElement('div');
        popup.id = 'arasaka-batch-popup';
        popup.setAttribute('data-tone', tone);
        popup.innerHTML = html;
        document.body.appendChild(popup);
        hudApplyPosition(popup);
        hudEnableDrag(popup, popup.querySelector('#wm-retoure-hud-drag'));
        popup.querySelector('#wm-retoure-hud-close').addEventListener('click', function() {
            popup.remove();
            lastHudSignature = '';
        });
    }

    function boot() {
        if (!pageHasText('Warenrückgabe')) return;
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
            showHud('WM RETOURE BEREIT', 'EAN scannen oder Artikelnummer tippen — Enter bestätigt, Scanner startet automatisch nach Pause.', false);
        }
        tickFlow(true);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', boot);
    } else {
        boot();
    }
})();
