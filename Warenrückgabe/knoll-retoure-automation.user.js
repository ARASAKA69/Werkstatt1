// ==UserScript==
// @name         KNOLL Warenrückgabe Retoure Bot
// @namespace    http://tampermonkey.net/
// @version      4.1
// @description  Automatisiert KNOLL Warenrückgabe per EAN-Scan
// @author       ARASAKA
// @match        *://shop.knoll.de/*
// @match        *://*.knoll.de/*
// @match        *://*.knoll-online.com/*
// @match        *://knoll-online.com/*
// @run-at       document-idle
// @updateURL    https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/Warenr%C3%BCckgabe/knoll-retoure-automation.user.js
// @downloadURL  https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/Warenr%C3%BCckgabe/knoll-retoure-automation.user.js
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    const BOT_VERSION = '4.1';
    const RETURN_ARTICLE_URL = 'https://shop.knoll.de/shop/my-account/return-article';
    const PENDING_RELOAD_KEY = 'knoll_retoure_pending_reload';
    const RELOAD_RETRY_KEY = 'knoll_retoure_reload_retry';
    const MAX_RELOAD_RETRIES = 12;
    const HUD_POS_KEY = 'knoll_retoure_hud_position';
    const ACTIVE_FLOW_KEY = 'knoll_retoure_active_flow';
    const FLOW_STATE_KEY = 'knoll_retoure_flow_state';
    const LAST_EAN_KEY = 'knoll_retoure_last_ean';
    const LAST_ADDED_KEY = 'knoll_retoure_last_added';
    const PENDING_ARTICLE_KEY = 'knoll_retoure_pending_article';
    const EXCLUSION_OK_KEY = 'knoll_retoure_exclusion_ok';
    const GRUND_TEXT = 'Passt nicht.';
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
    let lastStableValue = '';
    let scanTimer = null;
    let searchInput = null;
    let pageWatcherStarted = false;
    let lastHandledStep = '';
    let exclusionDialogOpen = false;
    let printButtonAdded = false;
    let grundFilled = false;
    let searchFocusDone = false;
    let searchFocusTimer = null;
    let orderHandledForSearch = false;
    let orderProcessing = false;
    let ordersFlowActive = false;
    let pendingReload = false;

    function isPendingReload() {
        return pendingReload || sessionStorage.getItem(PENDING_RELOAD_KEY) === '1';
    }

    function getReloadRetryCount() {
        return parseInt(sessionStorage.getItem(RELOAD_RETRY_KEY) || '0', 10) || 0;
    }

    function incrementReloadRetry() {
        var next = getReloadRetryCount() + 1;
        sessionStorage.setItem(RELOAD_RETRY_KEY, String(next));
        return next;
    }

    function clearReloadRetry() {
        sessionStorage.removeItem(RELOAD_RETRY_KEY);
    }

    function confirmCleanSearchPage() {
        sessionStorage.removeItem(PENDING_RELOAD_KEY);
        clearReloadRetry();
        pendingReload = false;
        orderHandledForSearch = false;
    }

    function hasStaleOrderListAfterAdd() {
        if (isReturnListPage() || loadFlowState().step === 'returnList') return false;
        if (!isFlowActive() || !isOrdersPage()) return false;
        var state = loadFlowState();
        if (state.step === 'orders' || state.step === 'articlePick' || state.searchSubmitted) return false;
        if (state.step && state.step !== 'search') return false;
        var input = findKnollSearchInput();
        return !input || !normalizeText(input.value);
    }

    function expectsCleanSearchPage() {
        if (isReturnListPage() || loadFlowState().step === 'returnList') return false;
        if (isPendingReload()) return true;
        return hasStaleOrderListAfterAdd();
    }

    function buildCleanRetoureReloadUrl() {
        try {
            var u = new URL(window.location.href);
            u.hash = '';
            u.searchParams.set('_knr', Date.now().toString(36));
            return u.toString();
        } catch (e) {
            return window.location.href.split('#')[0];
        }
    }

    function buildReturnArticleUrl() {
        try {
            var u = new URL(RETURN_ARTICLE_URL);
            u.searchParams.set('_knr', Date.now().toString(36));
            return u.toString();
        } catch (e) {
            return RETURN_ARTICLE_URL;
        }
    }

    function canAutoFocusSearch() {
        if (abortMission || userEditingSearch || searchSubmitted) return false;
        var input = findKnollSearchInput();
        if (!input || !isVisible(input)) return false;
        if (normalizeText(input.value)) return false;
        if (orderHandledForSearch) return true;
        if (isOrdersPage() || findReturnListTable()) return false;
        var page = detectPage();
        if (page !== 'search' && !(page === 'unknown' && isRetoureUrl())) return false;
        return true;
    }

    function focusReturnSearchInput(attempt) {
        if (abortMission || searchSubmitted) return;
        var input = findKnollSearchInput();
        if (!input || !isVisible(input)) {
            if (attempt < 15) setTimeout(function() { focusReturnSearchInput(attempt + 1); }, 180);
            return;
        }
        searchInput = input;
        bindSearchInput(searchInput);
        if (document.activeElement !== input) {
            try {
                input.focus({ preventScroll: true });
            } catch (e) {
                try { input.focus(); } catch (e2) {}
            }
            input.click();
        }
        if (document.activeElement === input) {
            searchFocusDone = true;
            return;
        }
        if (attempt < 15) setTimeout(function() { focusReturnSearchInput(attempt + 1); }, 180);
    }

    function scheduleSearchAutoFocus(force) {
        if (!force && searchFocusDone) return;
        if (searchFocusTimer) clearTimeout(searchFocusTimer);
        var delay = force ? (orderHandledForSearch ? 400 : 120) : 320;
        searchFocusTimer = setTimeout(function() {
            searchFocusTimer = null;
            if (!canAutoFocusSearch()) {
                if (orderHandledForSearch) focusReturnSearchInput(0);
                return;
            }
            focusReturnSearchInput(0);
        }, delay);
    }

    function loadFlowState() {
        try {
            var raw = sessionStorage.getItem(FLOW_STATE_KEY);
            if (!raw) return {};
            var state = JSON.parse(raw);
            return state && typeof state === 'object' ? state : {};
        } catch (e) {
            return {};
        }
    }

    function saveFlowState(updates) {
        try {
            var current = loadFlowState();
            var next = {
                active: current.active || sessionStorage.getItem(ACTIVE_FLOW_KEY) === '1',
                step: current.step || '',
                searchSubmitted: !!current.searchSubmitted,
                grundFilled: !!current.grundFilled,
                lastEan: current.lastEan || sessionStorage.getItem(LAST_EAN_KEY) || '',
                orderHandled: !!current.orderHandled
            };
            if (updates) {
                if (updates.active != null) next.active = !!updates.active;
                if (updates.step != null) next.step = String(updates.step || '');
                if (updates.searchSubmitted != null) next.searchSubmitted = !!updates.searchSubmitted;
                if (updates.grundFilled != null) next.grundFilled = !!updates.grundFilled;
                if (updates.lastEan != null) next.lastEan = String(updates.lastEan || '');
                if (updates.orderHandled != null) next.orderHandled = !!updates.orderHandled;
            }
            next.ts = Date.now();
            sessionStorage.setItem(FLOW_STATE_KEY, JSON.stringify(next));
            if (next.active) sessionStorage.setItem(ACTIVE_FLOW_KEY, '1');
            if (next.lastEan) sessionStorage.setItem(LAST_EAN_KEY, next.lastEan);
        } catch (e) {}
    }

    function clearFlowState() {
        sessionStorage.removeItem(FLOW_STATE_KEY);
        sessionStorage.removeItem(ACTIVE_FLOW_KEY);
        sessionStorage.removeItem(LAST_EAN_KEY);
        sessionStorage.removeItem(LAST_ADDED_KEY);
        sessionStorage.removeItem(PENDING_ARTICLE_KEY);
        sessionStorage.removeItem(EXCLUSION_OK_KEY);
        sessionStorage.removeItem(PENDING_RELOAD_KEY);
        sessionStorage.removeItem(RELOAD_RETRY_KEY);
        pendingReload = false;
    }

    function restoreFlowState() {
        var state = loadFlowState();
        if (state.active || sessionStorage.getItem(ACTIVE_FLOW_KEY) === '1') {
            sessionStorage.setItem(ACTIVE_FLOW_KEY, '1');
            searchSubmitted = !!state.searchSubmitted;
            grundFilled = !!state.grundFilled;
            orderHandledForSearch = !!state.orderHandled;
            if (state.step === 'orders' || state.step === 'returnList' || state.step === 'excluded' || state.step === 'submitPending') {
                searchSubmitted = false;
            }
            if (searchSubmitted || state.step === 'orders' || state.step === 'articlePick' || state.step === 'returnList' || state.step === 'submitPending') {
                searchListenersReady = true;
            }
        }
        lastHandledStep = '';
    }

    function isFlowActive() {
        return sessionStorage.getItem(ACTIVE_FLOW_KEY) === '1' || !!loadFlowState().active;
    }

    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            abortMission = true;
            isRunning = false;
            if (scanTimer) clearTimeout(scanTimer);
            clearFlowState();
            searchSubmitted = false;
            searchListenersReady = false;
            userEditingSearch = false;
            orderHandledForSearch = false;
            orderProcessing = false;
            ordersFlowActive = false;
            pendingReload = false;
            hideQuantityPicker();
            hideExclusionDialog();
            showHud('KNOLL RETOURE STOP', 'Prozess gestoppt. ESC erneut zum Schließen.', true);
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

    function saveLastAdded(info) {
        try {
            if (!info) return;
            sessionStorage.setItem(LAST_ADDED_KEY, JSON.stringify({
                part: String(info.part || ''),
                name: String(info.name || ''),
                qty: String(info.qty || ''),
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

    function savePendingArticle(info) {
        try {
            if (!info) return;
            sessionStorage.setItem(PENDING_ARTICLE_KEY, JSON.stringify({
                part: String(info.part || ''),
                name: String(info.name || '')
            }));
        } catch (e) {}
    }

    function loadPendingArticle() {
        try {
            var raw = sessionStorage.getItem(PENDING_ARTICLE_KEY);
            if (!raw) return null;
            var data = JSON.parse(raw);
            return data && typeof data === 'object' ? data : null;
        } catch (e) {
            return null;
        }
    }

    function parseArticleLabel(text) {
        var raw = normalizeText(text);
        if (!raw) return { part: '', name: '' };
        var m = raw.match(/^(.+?)\s*:\s*(.+)$/);
        if (m) return { part: normalizeText(m[1]), name: normalizeText(m[2]) };
        return { part: raw, name: '' };
    }

    function rememberArticleFromTarget(target) {
        if (!target) return;
        var parsed = parseArticleLabel(target.textContent || '');
        var ean = sessionStorage.getItem(LAST_EAN_KEY) || '';
        savePendingArticle({
            part: parsed.part || ean,
            name: parsed.name
        });
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
        var raw = String(value || '').trim();
        var m = raw.match(/(\d{2})\.(\d{2})\.(\d{2,4})/);
        if (!m) return null;
        var year = parseInt(m[3], 10);
        if (year < 100) year += 2000;
        return new Date(year, parseInt(m[2], 10) - 1, parseInt(m[1], 10));
    }

    function pageBodyText() {
        if (!document.body) return '';
        var clone = document.body.cloneNode(true);
        var overlays = clone.querySelectorAll('#knoll-retoure-popup, #knoll-retoure-qty-popup, #knoll-retoure-exclusion-popup, #knoll-retoure-print-btn');
        for (var i = 0; i < overlays.length; i++) overlays[i].remove();
        return clone.innerText;
    }

    function pageHasText(text) {
        return pageBodyText().indexOf(text) !== -1;
    }

    function isRetoureUrl() {
        var path = String(location.pathname || '').toLowerCase();
        return path.indexOf('return-article') !== -1 || path.indexOf('return') !== -1 && path.indexOf('account') !== -1;
    }

    function isSubmitConfirmationPage() {
        var text = pageBodyText().toLowerCase();
        if (text.indexOf('vielen dank') !== -1 && text.indexOf('anfrage') !== -1) return true;
        if (text.indexOf('ihre anfrage erhalten') !== -1) return true;
        if (text.indexOf('wir haben ihre anfrage') !== -1) return true;
        return false;
    }

    function shouldReturnToSearchAfterSubmit() {
        if (!isSubmitConfirmationPage()) return false;
        if (!isFlowActive()) return false;
        var state = loadFlowState();
        return state.step === 'returnList' || state.step === 'submitPending' || state.step === 'submitDone';
    }

    function isRetoureContext() {
        if (isRetoureUrl()) return true;
        if (shouldReturnToSearchAfterSubmit()) return true;
        if (pageHasText('ARTIKELNUMMER SUCHEN')) return true;
        if (pageHasText('Zur Rückgabeliste')) return true;
        if (pageHasText('Warenrückgabe') || pageHasText('WARENRÜCKGABE')) return true;
        if (findReturnListTable() || findOrdersTable()) return true;
        if (findKnollSearchInput()) return true;
        return false;
    }

    function hasExclusionWarning() {
        var text = pageBodyText().toLowerCase();
        if (text.indexOf('rücksprache') === -1 && text.indexOf('ruecksprache') === -1) return false;
        return text.indexOf('umtausch') !== -1 || text.indexOf('rückgabe') !== -1 || text.indexOf('rueckgabe') !== -1;
    }

    function findReturnTableItems() {
        var nodes = document.querySelectorAll('.return-table__item');
        var out = [];
        for (var i = 0; i < nodes.length; i++) {
            var item = nodes[i];
            if (!item.classList.contains('return-table__item')) continue;
            if (!isVisible(item)) continue;
            if (item.closest('.return-table__header, [class*="return-table__head"]')) continue;
            if (!item.querySelector('[class*="return-table__item--date"], input.returnQtyInput, input.returnQty, input[class*="returnQty"]')) continue;
            out.push(item);
        }
        return out;
    }

    function findOrdersTable() {
        var tables = document.querySelectorAll('table');
        for (var i = 0; i < tables.length; i++) {
            var t = normalizeText(tables[i].innerText).toLowerCase();
            if (t.indexOf('bestellnummer') === -1) continue;
            if (t.indexOf('ret') !== -1 && t.indexOf('menge') !== -1) return tables[i];
            if (t.indexOf('rückgabe') !== -1 || t.indexOf('rueckgabe') !== -1) return tables[i];
            if (t.indexOf('datum') !== -1 && tableHasOrderDates(tables[i])) return tables[i];
        }
        var items = findReturnTableItems();
        if (items.length) {
            return items[0].closest('.return-table, [class*="return-table"]') || items[0].parentElement;
        }
        return null;
    }

    function isDivOrderList(root) {
        if (!root) return findReturnTableItems().length > 0;
        if (root.matches && root.matches('.return-table__item, [class*="return-table__item"]')) return true;
        return findReturnTableItems().length > 0 && !root.querySelector('tr td');
    }

    function hasOrdersListContent() {
        if (findReturnTableItems().length >= 1) return true;
        if (findOrdersTable() && !isDivOrderList(findOrdersTable())) return true;
        var text = pageBodyText().toLowerCase();
        if (text.indexOf('bestellnummer') === -1 && findReturnTableItems().length === 0) return false;
        if (text.indexOf('neuteil') === -1 && text.indexOf('rückgabeart') === -1 && text.indexOf('ret.') === -1 && findReturnTableItems().length === 0) return false;
        var dates = pageBodyText().match(/\b\d{2}\.\d{2}\.\d{2,4}\b/g);
        return !!(dates && dates.length >= 2);
    }

    function isOrdersPage() {
        if (isReturnListPage()) return false;
        return hasOrdersListContent();
    }

    function rowLooksLikeOrderRow(el) {
        var text = normalizeText(el.textContent || '');
        if (!/(\d{2})\.(\d{2})\.(\d{2,4})/.test(text)) return false;
        return /\b\d{10}\b/.test(text.replace(/\s+/g, '')) || text.toLowerCase().indexOf('neuteil') !== -1;
    }

    function tableHasOrderDates(table) {
        var rows = table.querySelectorAll('tr');
        for (var i = 0; i < rows.length; i++) {
            if (/(\d{2})\.(\d{2})\.(\d{2,4})/.test(rows[i].textContent || '')) return true;
        }
        return false;
    }

    function findReturnListTable() {
        var tables = document.querySelectorAll('table');
        for (var i = 0; i < tables.length; i++) {
            var t = normalizeText(tables[i].innerText).toLowerCase();
            if (t.indexOf('grund') !== -1 && t.indexOf('bestellnummer') !== -1) return tables[i];
        }
        return null;
    }

    function findGrundFields() {
        var nodes = document.querySelectorAll('textarea.returnReasonComment, textarea[name="comment"][class*="returnReason"], textarea[id^="comment-"]');
        var out = [];
        for (var i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (!isVisible(el)) continue;
            if (el.closest('#knoll-retoure-popup, #knoll-retoure-qty-popup, #knoll-retoure-exclusion-popup')) continue;
            out.push(el);
        }
        return out;
    }

    function isReturnListPage() {
        if (!findClickableByText('Absenden', true)) return false;
        if (findGrundFields().length) return true;
        if (findReturnListTable()) return true;
        if (pageHasText('Grund') && document.querySelector('.return-table, [class*="return-table"]')) return true;
        return false;
    }

    function getTableHeaderRow(table) {
        return table.querySelector('thead tr') || table.querySelector('tr');
    }

    function getTableColumnIndex(table, headerText) {
        var want = normalizeText(headerText).toLowerCase();
        var headerRow = getTableHeaderRow(table);
        if (!headerRow) return -1;
        var headers = headerRow.querySelectorAll('th, td');
        for (var i = 0; i < headers.length; i++) {
            var h = normalizeText(headers[i].textContent).toLowerCase();
            if (h === want || h.indexOf(want) !== -1) return i;
        }
        return -1;
    }

    function detectPage() {
        if (shouldReturnToSearchAfterSubmit()) return 'submitDone';
        if (isReturnListPage()) return 'returnList';
        if (isOrdersPage()) return 'orders';
        if (hasExclusionWarning()) return 'excluded';
        if (searchSubmitted && !isOrdersPage() && (findAutocompleteItems().length || findInlineArticleResult())) return 'articlePick';
        if (pageHasText('ARTIKELNUMMER SUCHEN') || findKnollSearchInput()) return 'search';
        if (isRetoureUrl()) return 'search';
        return 'unknown';
    }

    function findAutocompleteItems() {
        var nodes = document.querySelectorAll('.ui-autocomplete li, .ui-menu li, .ui-menu-item, ul[id*="autocomplete"] li');
        var out = [];
        for (var i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (!isVisible(el)) continue;
            var t = normalizeText(el.textContent).toLowerCase();
            if (!t || t.indexOf('kein artikel') !== -1) continue;
            if (t.indexOf('es gibt noch mehr artikel') !== -1) continue;
            out.push(el);
        }
        return out;
    }

    function findInlineArticleResult() {
        if (isOrdersPage()) return null;
        var input = findKnollSearchInput();
        if (!input || !normalizeText(input.value)) return null;
        var searchVal = normalizeText(input.value).toLowerCase();
        var inputRect = input.getBoundingClientRect();
        var nodes = document.querySelectorAll('.ui-autocomplete li, .ui-menu li, a, button, li, div, span');
        for (var i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (!isVisible(el)) continue;
            if (el === input || el.contains(input)) continue;
            if (el.closest('#knoll-retoure-popup, #knoll-retoure-qty-popup, #knoll-retoure-exclusion-popup')) continue;
            if (el.closest('.return-table__item, .return-table, [class*="return-table"]')) continue;
            if (rowLooksLikeOrderRow(el)) continue;
            var t = normalizeText(el.textContent);
            if (t.indexOf(' : ') === -1) continue;
            if (t.toLowerCase().indexOf('artikelnummer suchen') !== -1) continue;
            if (searchVal && t.toLowerCase().indexOf(searchVal) === -1 && searchVal.length >= MIN_SCAN_LEN) {
                var compact = t.replace(/\s+/g, '');
                if (compact.toLowerCase().indexOf(searchVal) === -1) continue;
            }
            var rect = el.getBoundingClientRect();
            if (rect.top < inputRect.bottom - 5) continue;
            if (rect.top > inputRect.bottom + 260) continue;
            return el.closest('.ui-menu-item, a, li, button') || el;
        }
        return null;
    }

    async function goToOrdersFromArticlePick() {
        if (isPendingReload()) return;
        if (expectsCleanSearchPage()) {
            await forceCleanSearchReload('Seite wird neu geladen...');
            return;
        }
        if (orderHandledForSearch) {
            await reloadForNextProduct();
            return;
        }
        if (orderProcessing || ordersFlowActive) {
            ensureSearchListeners(true);
            return;
        }
        searchSubmitted = false;
        saveFlowState({ active: true, step: 'orders', searchSubmitted: false, orderHandled: false });
        lastHandledStep = '';
        await handleOrdersPage();
    }

    async function waitForSearchResult(timeout) {
        var start = Date.now();
        while (Date.now() - start < (timeout || 10000)) {
            if (abortMission) return null;
            if (isOrdersPage() || hasExclusionWarning()) return null;
            var items = findAutocompleteItems();
            if (items.length) return items[0];
            var card = findInlineArticleResult();
            if (card) return card;
            await sleep(180);
        }
        return findInlineArticleResult();
    }

    async function selectSearchResult(target) {
        if (!target) return false;
        rememberArticleFromTarget(target);
        var clickTarget = target.querySelector('a, div.ui-menu-item-wrapper, .ui-menu-item') || target;
        forceClick(clickTarget);
        await sleep(350);
        if (findOrdersTable() || hasExclusionWarning()) return true;
        forceClick(target);
        await sleep(350);
        return !!(findOrdersTable() || hasExclusionWarning());
    }

    async function selectSearchResultViaKeyboard() {
        if (!searchInput) searchInput = findKnollSearchInput();
        if (!searchInput) return false;
        searchInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', code: 'ArrowDown', keyCode: 40, which: 40, bubbles: true }));
        await sleep(280);
        searchInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
        searchInput.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
        await sleep(700);
        return !!(findOrdersTable() || isOrdersPage() || hasExclusionWarning());
    }

    async function handleArticlePickPage() {
        if (isOrdersPage()) {
            if (orderHandledForSearch || isPendingReload()) {
                reloadForNextProduct();
                return;
            }
            await goToOrdersFromArticlePick();
            return;
        }
        if (hasExclusionWarning()) {
            lastHandledStep = '';
            tickFlow(true);
            return;
        }
        showHud('KNOLL RETOURE LÄUFT', 'Wähle Artikel aus Trefferliste...', false, true);
        var target = await waitForSearchResult(10000);
        if (target) rememberArticleFromTarget(target);
        if (abortMission) return;
        if (isOrdersPage()) {
            if (orderHandledForSearch || isPendingReload()) {
                reloadForNextProduct();
                return;
            }
            await goToOrdersFromArticlePick();
            return;
        }
        if (pageHasText('Kein Artikel gefunden')) {
            searchSubmitted = false;
            saveFlowState({ searchSubmitted: false, step: 'search' });
            lastHandledStep = '';
            showHud('KNOLL RETOURE FEHLER', 'Kein Artikel gefunden.', true, true);
            return;
        }
        var ok = false;
        if (target) ok = await selectSearchResult(target);
        if (!ok && !isOrdersPage() && !hasExclusionWarning()) ok = await selectSearchResultViaKeyboard();
        var waitStart = Date.now();
        while (Date.now() - waitStart < 8000) {
            if (abortMission) return;
            if (isOrdersPage() || hasExclusionWarning()) {
                ok = true;
                break;
            }
            await sleep(200);
        }
        lastHandledStep = '';
        if (isOrdersPage()) {
            if (orderHandledForSearch || isPendingReload()) {
                reloadForNextProduct();
                return;
            }
            await goToOrdersFromArticlePick();
            return;
        }
        if (hasExclusionWarning()) {
            saveFlowState({ active: true, step: 'excluded', searchSubmitted: false });
            tickFlow(true);
            return;
        }
        if (!ok) {
            showHud('KNOLL RETOURE FEHLER', 'Artikel-Treffer konnte nicht ausgewählt werden.', true, true);
            searchSubmitted = false;
            saveFlowState({ searchSubmitted: false, step: 'search' });
        }
    }

    function isHeaderSearchInput(input) {
        if (!input) return true;
        var id = normalizeText(input.id).toLowerCase();
        if (id === 'js-return-article-quick-search-input') return false;
        if (id.indexOf('return-article') !== -1) return false;
        var ph = normalizeText(input.placeholder).toLowerCase();
        if (ph.indexOf('sortiment') !== -1) return true;
        if (ph.indexOf('schnellerfassung') !== -1) return true;
        var meta = normalizeText((input.name || '') + ' ' + (input.id || '') + ' ' + (input.className || '')).toLowerCase();
        if (meta.indexOf('global') !== -1 || meta.indexOf('header') !== -1) return true;
        if (meta.indexOf('quick') !== -1 && meta.indexOf('return-article') === -1) return true;
        if (input.closest('header, .header, .site-header, .top-header, .main-header, .navbar, .nav-bar')) return true;
        return false;
    }

    function isArtikelnummerLabel(node) {
        var t = normalizeText(node.textContent).toLowerCase();
        return t === 'artikelnummer suchen' || (t.indexOf('artikelnummer suchen') === 0 && t.length <= 24);
    }

    function findKnollSearchInput() {
        var direct = document.getElementById('js-return-article-quick-search-input');
        if (direct && isVisible(direct)) return direct;
        var bySelector = document.querySelector('input[id*="return-article"][id*="search"], input[name="text"][autocomplete="off"]');
        if (bySelector && isVisible(bySelector) && !isHeaderSearchInput(bySelector)) return bySelector;
        var marker = null;
        var nodes = document.querySelectorAll('label, div, span, td, th, p, h1, h2, h3, h4, strong, b');
        for (var i = 0; i < nodes.length; i++) {
            if (!isArtikelnummerLabel(nodes[i])) continue;
            if (marker && marker.contains(nodes[i])) continue;
            marker = nodes[i];
        }
        if (marker) {
            var labelFor = marker.getAttribute('for');
            if (labelFor) {
                var linked = document.getElementById(labelFor);
                if (linked && isVisible(linked) && !isHeaderSearchInput(linked)) return linked;
            }
            var notInList = findClickableByText('Der gewünschte Artikel ist nicht in der Liste', false);
            var allInputs = document.querySelectorAll('input[type="text"], input:not([type])');
            for (var k = 0; k < allInputs.length; k++) {
                var candidate = allInputs[k];
                if (!isVisible(candidate) || isHeaderSearchInput(candidate)) continue;
                if (!(marker.compareDocumentPosition(candidate) & Node.DOCUMENT_POSITION_FOLLOWING)) continue;
                if (notInList && !(notInList.compareDocumentPosition(candidate) & Node.DOCUMENT_POSITION_PRECEDING)) continue;
                return candidate;
            }
        }
        var links = document.querySelectorAll('a');
        for (var l = 0; l < links.length; l++) {
            if (normalizeText(links[l].textContent).indexOf('Der gewünschte Artikel ist nicht in der Liste') === -1) continue;
            var prev = links[l].previousElementSibling;
            while (prev) {
                if (prev.matches && prev.matches('input[type="text"], input:not([type])') && isVisible(prev) && !isHeaderSearchInput(prev)) return prev;
                var nested = prev.querySelector('input[type="text"], input:not([type])');
                if (nested && isVisible(nested) && !isHeaderSearchInput(nested)) return nested;
                prev = prev.previousElementSibling;
            }
            var block = links[l].parentElement;
            if (block) {
                var blockInputs = block.querySelectorAll('input[type="text"], input:not([type])');
                for (var b = 0; b < blockInputs.length; b++) {
                    if (isVisible(blockInputs[b]) && !isHeaderSearchInput(blockInputs[b])) return blockInputs[b];
                }
            }
        }
        return null;
    }

    function getOrderRows(root) {
        var items = findReturnTableItems();
        if (items.length) return items;
        root = root || findOrdersTable();
        if (!root) return [];
        var rows = root.querySelectorAll('tr');
        var out = [];
        for (var i = 0; i < rows.length; i++) {
            var cells = rows[i].querySelectorAll('td');
            if (cells.length < 4) continue;
            out.push(rows[i]);
        }
        return out;
    }

    function getRowDateText(row) {
        var dateEl = row.querySelector('[class*="return-table__item--date"], [class*="item--date"]');
        if (dateEl) return normalizeText(dateEl.textContent);
        var cells = row.querySelectorAll('td, [class*="return-table__item--"]');
        for (var i = 0; i < cells.length; i++) {
            var t = normalizeText(cells[i].textContent);
            if (/^\d{2}\.\d{2}\.\d{2,4}$/.test(t)) return t;
        }
        var m = normalizeText(row.textContent).match(/\b(\d{2}\.\d{2}\.\d{2,4})\b/);
        return m ? m[1] : '';
    }

    function isRowReturnable(table, row) {
        var input = getRowQuantityInput(row);
        if (!input) return false;
        var current = parseInt(input.value, 10);
        if (isNaN(current)) current = 0;
        if (current > 0) return true;
        var buttons = findRowQuantityButtons(row, input);
        if (buttons.plus) {
            if (buttons.plus.disabled || buttons.plus.classList.contains('disabled')) return false;
            return true;
        }
        return (getRowMaxQuantity(table, row) || 0) > 0;
    }

    function findBestOrderRow(minQty) {
        var rows = getOrderRows();
        if (!rows.length) return null;
        var table = findOrdersTable();
        var need = minQty && minQty > 0 ? minQty : 1;
        var bestRow = null;
        var bestDate = null;
        var fallbackRow = null;
        var fallbackDate = null;
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            if (!isRowReturnable(table, row)) continue;
            var date = parseGermanDate(getRowDateText(row));
            if (!date) continue;
            if (!fallbackDate || date.getTime() > fallbackDate.getTime()) {
                fallbackDate = date;
                fallbackRow = row;
            }
            var maxQty = getRowMaxQuantity(table, row) || 0;
            if (maxQty >= need) {
                if (!bestDate || date.getTime() > bestDate.getTime()) {
                    bestDate = date;
                    bestRow = row;
                }
            }
        }
        return bestRow || fallbackRow;
    }

    function findNewestOrderRow() {
        var rows = getOrderRows();
        var minQty = rows.length && isBrakeArticle(rows[0]) ? 2 : 1;
        return findBestOrderRow(minQty);
    }

    async function waitForNewestOrderRow(timeout) {
        var rows = getOrderRows();
        var minQty = rows.length && isBrakeArticle(rows[0]) ? 2 : 1;
        var start = Date.now();
        while (Date.now() - start < (timeout || 10000)) {
            if (abortMission) return null;
            var row = findBestOrderRow(minQty);
            if (row) return row;
            await sleep(180);
        }
        return null;
    }

    function parseNumericLabel(text) {
        var t = normalizeText(text);
        if (!t) return null;
        if (/^\d{2}\.\d{2}\.\d{2,4}$/.test(t)) return null;
        if (!/^\d{1,2}$/.test(t)) return null;
        var num = parseInt(t, 10);
        return !isNaN(num) && num > 0 && num <= 99 ? num : null;
    }

    function getRowMaxReturnCell(row) {
        var maxCell = row.querySelector('.return-table__item--max-return, [class*="return-table__item--max-return"]');
        if (!maxCell) return null;
        var inner = maxCell.querySelector('div[id]');
        if (inner) {
            var fromInner = parseNumericLabel(inner.textContent);
            if (fromInner) return fromInner;
        }
        return parseNumericLabel(maxCell.textContent);
    }

    function getRowRetMengeValue(row, table) {
        var fromMaxReturn = getRowMaxReturnCell(row);
        if (fromMaxReturn) return fromMaxReturn;
        var clsCells = row.querySelectorAll('[class*="return-table__item--"]');
        for (var i = 0; i < clsCells.length; i++) {
            var clsCell = clsCells[i];
            if (clsCell.querySelector('input, button, select')) continue;
            var clsName = normalizeText(clsCell.className).toLowerCase();
            if (clsName.indexOf('max-return') !== -1 || clsName.indexOf('maxreturn') !== -1) {
                var fromMaxClass = parseNumericLabel(clsCell.textContent);
                if (fromMaxClass) return fromMaxClass;
            }
            if (clsName.indexOf('date') !== -1 || clsName.indexOf('orderid') !== -1 || clsName.indexOf('order') !== -1 && clsName.indexOf('menge') === -1) continue;
            if (clsName.indexOf('name') !== -1 || clsName.indexOf('article') !== -1 || clsName.indexOf('articlenumber') !== -1 || clsName.indexOf('product') !== -1) continue;
            if (clsName.indexOf('button') !== -1 || clsName.indexOf('reason') !== -1 || clsName.indexOf('quantity') !== -1) continue;
            if (clsName.indexOf('ret') !== -1 && clsName.indexOf('menge') !== -1) {
                var fromRetClass = parseNumericLabel(clsCell.textContent);
                if (fromRetClass) return fromRetClass;
            }
        }
        if (table) {
            var idx = getTableColumnIndex(table, 'Ret. menge');
            if (idx < 0) idx = getTableColumnIndex(table, 'Ret.menge');
            if (idx >= 0) {
                var tds = row.querySelectorAll('td');
                if (tds[idx]) {
                    var fromCol = parseNumericLabel(tds[idx].textContent);
                    if (fromCol) return fromCol;
                }
            }
        }
        return null;
    }

    function isBrakeArticle(row) {
        var text = normalizeText(row.textContent).toLowerCase();
        return text.indexOf('bremsscheibe') !== -1 ||
            text.indexOf('bremsbelag') !== -1 ||
            text.indexOf('bremsklotz') !== -1 ||
            text.indexOf('bremssattel') !== -1;
    }

    function resolveReturnQuantity(row, maxQty) {
        var max = maxQty && maxQty > 0 ? maxQty : 1;
        if (isBrakeArticle(row) && max >= 2) return Math.min(2, max);
        return 1;
    }

    function getRowMaxQuantity(table, row) {
        var retMenge = getRowRetMengeValue(row, table);
        if (retMenge) return retMenge;
        var input = getRowQuantityInput(row);
        if (input) {
            var maxAttr = parseNumericLabel(input.getAttribute('max') || input.getAttribute('data-max'));
            if (maxAttr) return maxAttr;
        }
        return 1;
    }

    function getRowQuantityInput(row) {
        var qtyCell = row.querySelector('.return-table__item--quantity, [class*="return-table__item--quantity"]');
        if (qtyCell) {
            var scoped = qtyCell.querySelector('input.returnQtyInput, input.returnQty, input[class*="returnQty"], input[id*="returnableQuantity"], input[type="text"], input[type="number"]');
            if (scoped && isVisible(scoped)) return scoped;
        }
        var input = row.querySelector('input.returnQtyInput, input.returnQty, input[class*="returnQty"], input[id*="returnableQuantity"]');
        if (input && isVisible(input)) return input;
        var inputs = row.querySelectorAll('input[type="text"], input[type="number"]');
        for (var i = 0; i < inputs.length; i++) {
            if (isVisible(inputs[i])) return inputs[i];
        }
        return null;
    }

    function findRowQuantityButtons(row, input) {
        var qtyHost = row.querySelector('.return-table__item--quantity, [class*="return-table__item--quantity"]') || row;
        var plus = qtyHost.querySelector('.btn-plus, .js-qty-selector-plus-button, .btn-increment, button.update-entry-quantity-button.btn-plus');
        var minus = qtyHost.querySelector('.btn-minus, .js-qty-selector-minus-button, .btn-decrement, button.update-entry-quantity-button.btn-minus');
        if (!plus || !minus) {
            var btns = qtyHost.querySelectorAll('button');
            for (var i = 0; i < btns.length; i++) {
                var btn = btns[i];
                if (btn.type === 'submit' || btn.classList.contains('btn-primary')) continue;
                var label = normalizeText(btn.textContent || btn.value);
                if (label === '+' || label === '−' || label === '-') {
                    if (label === '+') plus = plus || btn;
                    else minus = minus || btn;
                }
            }
        }
        return { minus: minus, plus: plus };
    }

    async function setRowQuantity(row, targetQty) {
        var input = getRowQuantityInput(row);
        if (!input) return false;
        var current = parseInt(input.value, 10);
        if (isNaN(current)) current = 0;
        if (current === targetQty) return true;
        var buttons = findRowQuantityButtons(row, input);
        var guard = 0;
        while (current < targetQty && buttons.plus && guard < 30) {
            forceClick(buttons.plus);
            await sleep(280);
            current = parseInt(input.value, 10);
            if (isNaN(current)) current = 0;
            if (current >= targetQty) break;
            guard++;
        }
        guard = 0;
        while (current > targetQty && buttons.minus && guard < 30) {
            forceClick(buttons.minus);
            await sleep(220);
            current = parseInt(input.value, 10);
            if (isNaN(current)) current = 0;
            guard++;
        }
        if (current !== targetQty) {
            input.focus();
            input.value = String(targetQty);
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            input.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));
            input.dispatchEvent(new Event('blur', { bubbles: true }));
            await sleep(350);
        }
        current = parseInt(input.value, 10);
        if (current === targetQty) return true;
        if (buttons.plus && current < targetQty) {
            forceClick(buttons.plus);
            await sleep(350);
            current = parseInt(input.value, 10);
        }
        if (current === targetQty) return true;
        return false;
    }

    function findRowAddButton(row) {
        var host = row.querySelector('[class*="return-table__item--button"], .return-table__item--button');
        if (host) {
            var hosted = host.querySelector('button[type="submit"], button.btn-primary');
            if (hosted && isVisible(hosted)) return hosted;
        }
        var submits = row.querySelectorAll('button[type="submit"].btn-primary, button.btn-primary[type="submit"], button.btn-primary');
        for (var i = 0; i < submits.length; i++) {
            var btn = submits[i];
            if (!isVisible(btn)) continue;
            if (btn.classList.contains('btn-plus') || btn.classList.contains('btn-minus')) continue;
            if (btn.classList.contains('btn-increment') || btn.classList.contains('btn-decrement')) continue;
            return btn;
        }
        var cells = row.querySelectorAll('td, [class*="return-table__item--button"]');
        for (var c = cells.length - 1; c >= 0; c--) {
            var btns = cells[c].querySelectorAll('button[type="submit"], button.btn-primary, input[type="submit"]');
            for (var j = 0; j < btns.length; j++) {
                if (isVisible(btns[j])) return btns[j];
            }
        }
        return null;
    }

    function buildQuantityOptions(maxQty) {
        var out = [];
        var max = maxQty && maxQty > 0 ? maxQty : 1;
        for (var i = 1; i <= max; i++) {
            out.push({ index: i - 1, numeric: i, label: String(i) });
        }
        return out;
    }

    let lastBoundSearchInput = null;

    function pauseWatcher(ms) {
        pausePageWatcherUntil = Date.now() + (ms || 3000);
    }

    function shouldPauseWatcher() {
        if (isPendingReload()) return true;
        if (searchSubmitted) return false;
        if (Date.now() < pausePageWatcherUntil) return true;
        if (!userEditingSearch || !searchInput) return false;
        if (document.activeElement === searchInput) return true;
        if (normalizeText(searchInput.value) && !searchSubmitted) return true;
        return false;
    }

    function handleSearchInput(e) {
        if (abortMission) return;
        searchInput = e.target;
        userEditingSearch = true;
        pauseWatcher(6000);
        lastInputChangeAt = Date.now();
        lastStableValue = normalizeText(searchInput.value);
        var value = lastStableValue;
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
            var value = normalizeText(searchInput.value);
            if (value.length >= MIN_SCAN_LEN) triggerSearch(false);
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
        orderHandledForSearch = false;
        orderProcessing = false;
        ordersFlowActive = false;
        lastHandledStep = '';
        pauseWatcher(1500);
        saveFlowState({ active: true, searchSubmitted: true, step: 'articlePick', lastEan: value, orderHandled: false });
        savePendingArticle({ part: value, name: '' });
        sessionStorage.removeItem(EXCLUSION_OK_KEY);
        lastHandledStep = '';
        grundFilled = false;
        printButtonAdded = false;
        showHud('KNOLL RETOURE LÄUFT', 'Suche Artikel für Nr.: ' + value, false, true);
        await sleep(500);
        if (abortMission || !searchInput) return;
        searchInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
        searchInput.dispatchEvent(new KeyboardEvent('keypress', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
        searchInput.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
        setTimeout(function() {
            tickFlow(true);
        }, 700);
    }

    function resetForNextScan(clearField) {
        searchSubmitted = false;
        userEditingSearch = false;
        searchListenersReady = false;
        searchFocusDone = false;
        if (!orderHandledForSearch) {
            lastHandledStep = '';
        }
        lastHudSignature = '';
        grundFilled = false;
        saveFlowState({
            active: true,
            searchSubmitted: false,
            step: 'search',
            grundFilled: false,
            lastEan: '',
            orderHandled: orderHandledForSearch
        });
        sessionStorage.removeItem(EXCLUSION_OK_KEY);
        if (ensureSearchListeners(clearField)) return;
        setTimeout(function() {
            ensureSearchListeners(clearField);
        }, 400);
    }

    function ensureSearchListeners(clearField) {
        var input = findKnollSearchInput();
        if (!input) return false;
        searchInput = input;
        bindSearchInput(searchInput);
        if (clearField) {
            searchSubmitted = false;
            userEditingSearch = false;
            if (scanTimer) {
                clearTimeout(scanTimer);
                scanTimer = null;
            }
            searchInput.value = '';
            lastStableValue = '';
            lastInputChangeAt = 0;
        }
        if (!searchListenersReady || clearField) {
            showHud('KNOLL RETOURE BEREIT', 'EAN scannen oder Artikelnummer tippen — Enter bestätigt, Scanner startet automatisch nach Pause.', false, true);
        }
        searchListenersReady = true;
        if (clearField) searchFocusDone = false;
        scheduleSearchAutoFocus(!!clearField);
        return true;
    }

    function setupSearchListeners(clearField) {
        return ensureSearchListeners(clearField);
    }

    function hideHud() {
        hideQuantityPicker();
        hideExclusionDialog();
        var popup = document.getElementById('knoll-retoure-popup');
        if (popup) popup.remove();
        lastHudSignature = '';
    }

    function hideBotOverlaysForPrint() {
        hideHud();
        document.querySelectorAll('#knoll-retoure-popup, #knoll-retoure-qty-popup, #knoll-retoure-exclusion-popup').forEach(function(el) {
            el.setAttribute('data-knoll-print-hidden', '1');
            el.style.display = 'none';
        });
    }

    function restoreBotOverlaysAfterPrint() {
        document.querySelectorAll('[data-knoll-print-hidden="1"]').forEach(function(el) {
            el.style.display = '';
            el.removeAttribute('data-knoll-print-hidden');
        });
    }

    function printPageClean() {
        hideBotOverlaysForPrint();
        window.print();
    }

    function setupPrintListeners() {
        if (window.__knollRetourePrintListeners) return;
        window.__knollRetourePrintListeners = true;
        window.addEventListener('beforeprint', hideBotOverlaysForPrint);
        window.addEventListener('afterprint', restoreBotOverlaysAfterPrint);
    }

    function bindPrintButton(btn) {
        if (!btn || btn.__knollPrintBound) return;
        btn.__knollPrintBound = true;
        styleKnollPrintButton(btn);
        btn.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            printPageClean();
        }, true);
    }

    function styleKnollPrintButton(btn) {
        if (!btn) return;
        hudEnsureStyles();
        btn.classList.add('knoll-retoure-print-knoll');
        var ref = findClickableByText('Zurück', true) || findClickableByText('Absenden', true);
        if (ref) {
            var skip = { 'btn-primary': 1, 'btn-danger': 1, 'btn-success': 1 };
            (ref.className || '').split(/\s+/).forEach(function(cls) {
                if (cls && !skip[cls]) btn.classList.add(cls);
            });
            if (!btn.classList.contains('btn')) btn.classList.add('btn');
            try {
                var cs = window.getComputedStyle(ref);
                btn.style.boxSizing = cs.boxSizing || 'border-box';
                btn.style.display = cs.display === 'inline' ? 'inline-block' : cs.display;
                btn.style.minHeight = cs.minHeight !== '0px' ? cs.minHeight : cs.height;
                btn.style.height = cs.height;
                btn.style.width = cs.width;
                btn.style.flex = cs.flex;
                btn.style.flexGrow = cs.flexGrow;
                btn.style.flexShrink = cs.flexShrink;
                btn.style.flexBasis = cs.flexBasis;
                btn.style.padding = cs.padding;
                btn.style.margin = cs.margin;
                btn.style.fontSize = cs.fontSize;
                btn.style.fontWeight = cs.fontWeight;
                btn.style.fontFamily = cs.fontFamily;
                btn.style.lineHeight = cs.lineHeight;
                btn.style.textTransform = cs.textTransform;
                btn.style.letterSpacing = cs.letterSpacing;
                btn.style.borderRadius = cs.borderRadius;
                btn.style.textAlign = 'center';
                btn.style.verticalAlign = cs.verticalAlign;
                btn.style.appearance = 'none';
                btn.style.webkitAppearance = 'none';
            } catch (e) {}
        }
    }

    function findReturnListActionBar() {
        var absenden = findClickableByText('Absenden', true);
        var zurueck = findClickableByText('Zurück', true);
        if (absenden && zurueck && absenden.parentElement === zurueck.parentElement) return absenden.parentElement;
        if (absenden && absenden.parentElement) return absenden.parentElement;
        if (zurueck && zurueck.parentElement) return zurueck.parentElement;
        return null;
    }
    function hideQuantityPicker() {
        var existing = document.getElementById('knoll-retoure-qty-popup');
        if (existing) existing.remove();
    }

    function hideExclusionDialog() {
        exclusionDialogOpen = false;
        var existing = document.getElementById('knoll-retoure-exclusion-popup');
        if (existing) existing.remove();
    }

    function showExclusionDialog(onContinue, onSkip) {
        hideExclusionDialog();
        hudEnsureStyles();
        exclusionDialogOpen = true;
        var popup = document.createElement('div');
        popup.id = 'knoll-retoure-exclusion-popup';
        popup.setAttribute('data-tone', 'warn');
        popup.innerHTML = ''
            + '<div class="arasaka-hud-head" id="knoll-retoure-exclusion-drag">'
            + '  <div>'
            + '    <div class="arasaka-hud-eyebrow">KNOLL Retoure · RÜCKGABE EINGESCHRÄNKT</div>'
            + '    <div class="arasaka-hud-title">Artikel nicht ohne Rücksprache</div>'
            + '  </div>'
            + '  <div class="arasaka-hud-actions">'
            + '    <button type="button" class="arasaka-hud-btn" id="knoll-retoure-exclusion-close">X</button>'
            + '  </div>'
            + '</div>'
            + '<div class="arasaka-hud-body">'
            + '  <div class="arasaka-hud-main">'
            + '    <div class="arasaka-hud-badge">WARNUNG</div>'
            + '    <div class="arasaka-hud-message">Umtausch/Rückgabe nur nach Rücksprache mit der Filiale. Diesen Artikel kannst du so nicht zurücksenden.<br><br>Weiter zur Rückgabeliste oder nächsten Scan?</div>'
            + '  </div>'
            + '  <div class="knoll-retoure-action-row">'
            + '    <button type="button" class="knoll-retoure-action-btn knoll-retoure-action-primary" id="knoll-retoure-exclusion-continue">Zur Rückgabeliste</button>'
            + '    <button type="button" class="knoll-retoure-action-btn" id="knoll-retoure-exclusion-skip">Nächster Scan</button>'
            + '  </div>'
            + '  ' + hudFooterHtml()
            + '</div>';
        document.body.appendChild(popup);
        hudApplyPosition(popup);
        hudEnableDrag(popup, popup.querySelector('#knoll-retoure-exclusion-drag'));
        popup.querySelector('#knoll-retoure-exclusion-close').addEventListener('click', function() {
            hideExclusionDialog();
            onSkip();
        });
        popup.querySelector('#knoll-retoure-exclusion-skip').addEventListener('click', function() {
            hideExclusionDialog();
            onSkip();
        });
        popup.querySelector('#knoll-retoure-exclusion-continue').addEventListener('click', function() {
            hideExclusionDialog();
            sessionStorage.setItem(EXCLUSION_OK_KEY, '1');
            onContinue();
        });
    }

    function showQuantityPicker(options, onPick) {
        hideQuantityPicker();
        hudEnsureStyles();
        var popup = document.createElement('div');
        popup.id = 'knoll-retoure-qty-popup';
        popup.setAttribute('data-tone', 'active');
        var buttonsHtml = options.map(function(opt) {
            return '<button type="button" class="knoll-retoure-qty-btn" data-index="' + opt.index + '" data-numeric="' + opt.numeric + '">' + hudEscape(opt.label) + '</button>';
        }).join('');
        popup.innerHTML = ''
            + '<div class="arasaka-hud-head" id="knoll-retoure-qty-drag">'
            + '  <div>'
            + '    <div class="arasaka-hud-eyebrow">KNOLL Retoure · MENGE WÄHLEN</div>'
            + '    <div class="arasaka-hud-title">Rückgabemenge</div>'
            + '  </div>'
            + '  <div class="arasaka-hud-actions">'
            + '    <button type="button" class="arasaka-hud-btn" id="knoll-retoure-qty-close">X</button>'
            + '  </div>'
            + '</div>'
            + '<div class="arasaka-hud-body">'
            + '  <div class="arasaka-hud-main">'
            + '    <div class="arasaka-hud-badge">MENGE</div>'
            + '    <div class="arasaka-hud-message">Wie viele Stück zurückgeben?</div>'
            + '  </div>'
            + '  <div class="knoll-retoure-qty-grid">' + buttonsHtml + '</div>'
            + '  ' + hudFooterHtml()
            + '</div>';
        document.body.appendChild(popup);
        hudApplyPosition(popup);
        hudEnableDrag(popup, popup.querySelector('#knoll-retoure-qty-drag'));
        popup.querySelector('#knoll-retoure-qty-close').addEventListener('click', function() {
            popup.remove();
            if (!orderHandledForSearch && !orderProcessing) ordersFlowActive = false;
        });
        popup.querySelectorAll('.knoll-retoure-qty-btn').forEach(function(btn) {
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
                onPick(picked || options[0]);
            });
        });
    }

    async function clickZurRueckgabeliste() {
        confirmCleanSearchPage();
        saveFlowState({ active: true, step: 'returnList', searchSubmitted: false, orderHandled: false });
        showHud('KNOLL RETOURE LÄUFT', 'Öffne Rückgabeliste...', false, true);
        var btn = findClickableByText('Zur Rückgabeliste', false);
        if (!btn) btn = await waitForClickableByText('Zur Rückgabeliste', false, 8000);
        if (!btn || abortMission) {
            if (isSubmitConfirmationPage() || isFlowActive()) {
                await returnToWarenrueckgabeSearch('Zurück zur Artikelsuche...');
                return true;
            }
            showHud('KNOLL RETOURE FEHLER', 'Zur Rückgabeliste nicht gefunden.', true, true);
            return false;
        }
        await sleep(250);
        forceClick(btn);
        setTimeout(function() {
            tickFlow(true);
        }, 700);
        return true;
    }

    async function reloadForNextProduct() {
        if (abortMission) return;
        pendingReload = true;
        sessionStorage.setItem(PENDING_RELOAD_KEY, '1');
        orderHandledForSearch = false;
        orderProcessing = false;
        ordersFlowActive = false;
        searchSubmitted = false;
        searchListenersReady = false;
        searchFocusDone = false;
        userEditingSearch = false;
        lastHandledStep = '';
        saveFlowState({
            active: true,
            step: 'search',
            searchSubmitted: false,
            orderHandled: false,
            grundFilled: false,
            lastEan: ''
        });
        sessionStorage.setItem(ACTIVE_FLOW_KEY, '1');
        pauseWatcher(15000);
        await sleep(200);
        location.replace(buildCleanRetoureReloadUrl());
    }

    async function returnToWarenrueckgabeSearch(message) {
        if (abortMission) return;
        pendingReload = true;
        sessionStorage.setItem(PENDING_RELOAD_KEY, '1');
        orderHandledForSearch = false;
        orderProcessing = false;
        ordersFlowActive = false;
        searchSubmitted = false;
        searchListenersReady = false;
        searchFocusDone = false;
        userEditingSearch = false;
        lastHandledStep = '';
        grundFilled = false;
        saveFlowState({
            active: true,
            step: 'search',
            searchSubmitted: false,
            orderHandled: false,
            grundFilled: false,
            lastEan: ''
        });
        sessionStorage.setItem(ACTIVE_FLOW_KEY, '1');
        pauseWatcher(15000);
        showHud('KNOLL RETOURE BEREIT', message || 'Zurück zur Artikelsuche...', false, false);
        await sleep(450);
        location.replace(buildReturnArticleUrl());
    }

    async function forceCleanSearchReload(message) {
        if (abortMission) return false;
        if (!expectsCleanSearchPage() && !isOrdersPage()) {
            confirmCleanSearchPage();
            return false;
        }
        var retries = incrementReloadRetry();
        if (retries > MAX_RELOAD_RETRIES) {
            confirmCleanSearchPage();
            showHud('KNOLL RETOURE FEHLER', 'Seite bitte manuell mit F5 neu laden, dann erneut scannen.', true, true);
            return false;
        }
        showHud('KNOLL RETOURE BEREIT', message || 'Seite wird neu geladen...', false, true);
        await sleep(350);
        await reloadForNextProduct();
        return true;
    }

    async function finishOrderAndReturnToSearch() {
        showHud('KNOLL RETOURE BEREIT', 'Artikel hinzugefügt — Seite wird neu geladen...', false, true);
        await sleep(500);
        await reloadForNextProduct();
    }

    async function submitOrderRow(row, quantity) {
        if (isPendingReload() || orderHandledForSearch || orderProcessing) return false;
        var table = findOrdersTable();
        if (!table || !row) return false;
        orderProcessing = true;
        showHud('KNOLL RETOURE LÄUFT', 'Setze Menge ' + quantity + ' und füge Artikel hinzu...', false, true);
        try {
            var ok = await setRowQuantity(row, quantity);
            if (!ok) {
                orderProcessing = false;
                ordersFlowActive = false;
                showHud('KNOLL RETOURE FEHLER', 'Rückgabemenge konnte nicht gesetzt werden.', true, true);
                return false;
            }
            await sleep(250);
            var addBtn = findRowAddButton(row);
            if (!addBtn) {
                orderProcessing = false;
                ordersFlowActive = false;
                showHud('KNOLL RETOURE FEHLER', 'Hinzufügen-Button in der Zeile nicht gefunden.', true, true);
                return false;
            }
            saveFlowState({ active: true, step: 'search', searchSubmitted: false, orderHandled: false });
            lastHandledStep = 'order-done';
            var pending = loadPendingArticle() || {};
            saveLastAdded({
                part: pending.part || sessionStorage.getItem(LAST_EAN_KEY) || '',
                name: pending.name || '',
                qty: quantity
            });
            sessionStorage.removeItem(PENDING_ARTICLE_KEY);
            forceClick(addBtn);
            await sleep(400);
            await finishOrderAndReturnToSearch();
            return true;
        } catch (e) {
            if (!isPendingReload()) {
                orderProcessing = false;
                ordersFlowActive = false;
            }
            return false;
        }
    }

    async function handleExcludedPage() {
        if (sessionStorage.getItem(EXCLUSION_OK_KEY) === '1') {
            await clickZurRueckgabeliste();
            return;
        }
        if (exclusionDialogOpen) return;
        showHud('KNOLL RETOURE WARNUNG', 'Rückgabe/Umtausch nur nach Rücksprache — Entscheidung nötig.', false, true);
        showExclusionDialog(
            function() {
                clickZurRueckgabeliste();
            },
            function() {
                resetForNextScan(true);
            }
        );
    }

    async function handleOrdersPage() {
        if (isPendingReload()) return;
        if (expectsCleanSearchPage()) {
            await forceCleanSearchReload('Alte Liste erkannt — Seite wird neu geladen...');
            return;
        }
        if (orderHandledForSearch) {
            await reloadForNextProduct();
            return;
        }
        if (orderProcessing) return;
        if (ordersFlowActive) return;
        ordersFlowActive = true;
        saveFlowState({ active: true, step: 'orders', searchSubmitted: false });
        searchSubmitted = false;
        try {
            if (hasExclusionWarning() && sessionStorage.getItem(EXCLUSION_OK_KEY) !== '1') {
                ordersFlowActive = false;
                await handleExcludedPage();
                return;
            }
            var previewRows = getOrderRows();
            var wantBrakePair = previewRows.length && isBrakeArticle(previewRows[0]);
            showHud('KNOLL RETOURE LÄUFT', 'Wähle neuesten Auftrag mit Menge ' + (wantBrakePair ? '2' : '1') + '...', false, true);
            var row = await waitForNewestOrderRow(12000);
            if (!row || abortMission || orderHandledForSearch) {
                if (!orderHandledForSearch) {
                    ordersFlowActive = false;
                    lastHandledStep = '';
                    if (hasStaleOrderListAfterAdd() || expectsCleanSearchPage()) {
                        await forceCleanSearchReload('Keine Zeile mehr verfügbar — Seite wird neu geladen...');
                        return;
                    }
                    showHud('KNOLL RETOURE FEHLER', 'Keine passende Bestellzeile gefunden.', true, true);
                }
                return;
            }
            var table = findOrdersTable();
            var maxQty = getRowMaxQuantity(table, row) || 1;
            var chosenQty = resolveReturnQuantity(row, maxQty);
            showHud('KNOLL RETOURE LÄUFT', 'Auftrag ' + getRowDateText(row) + ' — Menge ' + chosenQty + ' (max ' + maxQty + ')...', false, true);
            try {
                row.scrollIntoView({ block: 'center', behavior: 'auto' });
            } catch (e) {}
            await sleep(250);
            if (orderHandledForSearch) return;
            maxQty = getRowMaxQuantity(table, row) || 1;
            chosenQty = resolveReturnQuantity(row, maxQty);
            if (maxQty === 1 || (isBrakeArticle(row) && chosenQty <= maxQty)) {
                showHud('KNOLL RETOURE LÄUFT', 'Menge ' + chosenQty + ' — füge Artikel hinzu...', false, true);
                await submitOrderRow(row, chosenQty);
                return;
            }
            var options = buildQuantityOptions(maxQty);
            if (options.length === 1) {
                await submitOrderRow(row, options[0].numeric);
                return;
            }
            showQuantityPicker(options, function(option) {
                if (orderHandledForSearch || orderProcessing) return;
                submitOrderRow(row, option.numeric);
            });
        } catch (e) {
            if (!orderHandledForSearch) ordersFlowActive = false;
        }
    }

    function fillGrundFields() {
        var fields = findGrundFields();
        if (!fields.length) {
            var table = findReturnListTable();
            if (!table) return false;
            var grundIdx = getTableColumnIndex(table, 'Grund');
            var rows = getOrderRows(table);
            var changed = false;
            for (var i = 0; i < rows.length; i++) {
                var cells = rows[i].querySelectorAll('td');
                var input = null;
                if (grundIdx >= 0 && cells[grundIdx]) {
                    input = cells[grundIdx].querySelector('input[type="text"], textarea');
                }
                if (!input) input = rows[i].querySelector('input[type="text"], textarea');
                if (!input || !isVisible(input)) continue;
                if (normalizeText(input.value) === GRUND_TEXT) continue;
                input.value = GRUND_TEXT;
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
                changed = true;
            }
            return changed;
        }
        var changed = false;
        for (var j = 0; j < fields.length; j++) {
            var field = fields[j];
            if (normalizeText(field.value) === GRUND_TEXT) continue;
            field.value = GRUND_TEXT;
            field.dispatchEvent(new Event('input', { bubbles: true }));
            field.dispatchEvent(new Event('change', { bubbles: true }));
            changed = true;
        }
        return changed;
    }

    function allGrundFieldsFilled() {
        var fields = findGrundFields();
        if (!fields.length) return false;
        for (var i = 0; i < fields.length; i++) {
            if (normalizeText(fields[i].value) !== GRUND_TEXT) return false;
        }
        return true;
    }

    function ensurePrintButton() {
        var existing = document.getElementById('knoll-retoure-print-btn');
        if (existing) {
            styleKnollPrintButton(existing);
            bindPrintButton(existing);
            return;
        }
        var native = findClickableByText('Seite drucken', false);
        if (native) {
            if (!native.id) native.id = 'knoll-retoure-print-btn';
            var absendenNative = findClickableByText('Absenden', true);
            var actionBar = findReturnListActionBar();
            if (actionBar && native.parentElement !== actionBar) {
                if (absendenNative && absendenNative.parentElement === actionBar) {
                    actionBar.insertBefore(native, absendenNative);
                } else {
                    actionBar.appendChild(native);
                }
            }
            styleKnollPrintButton(native);
            bindPrintButton(native);
            printButtonAdded = true;
            return;
        }
        var absenden = findClickableByText('Absenden', true);
        var zurueck = findClickableByText('Zurück', true);
        var host = findReturnListActionBar();
        if (!host) {
            var table = findReturnListTable();
            host = table ? table.parentElement : null;
        }
        if (!host) {
            var grund = findGrundFields()[0];
            host = grund ? grund.closest('.return-table, [class*="return-table"]') : null;
            if (host && host.parentElement) host = host.parentElement;
        }
        if (!host) host = document.body;
        var btn = document.createElement('button');
        btn.id = 'knoll-retoure-print-btn';
        btn.type = 'button';
        btn.textContent = 'Seite drucken';
        styleKnollPrintButton(btn);
        bindPrintButton(btn);
        if (absenden && absenden.parentElement === host) {
            host.insertBefore(btn, absenden);
        } else if (zurueck && zurueck.parentElement === host) {
            host.insertBefore(btn, zurueck.nextSibling);
        } else {
            host.appendChild(btn);
        }
        printButtonAdded = true;
    }

    function bindAbsendenButton() {
        var absenden = findClickableByText('Absenden', true);
        if (!absenden || absenden.__knollAbsendenBound) return;
        absenden.__knollAbsendenBound = true;
        absenden.addEventListener('click', function() {
            saveFlowState({ active: true, step: 'submitPending', searchSubmitted: false, orderHandled: false, grundFilled: grundFilled });
            sessionStorage.setItem(ACTIVE_FLOW_KEY, '1');
        }, true);
    }

    async function handleSubmitDonePage() {
        if (abortMission || isPendingReload()) return;
        showHud('KNOLL RETOURE BEREIT', 'Absenden bestätigt — zurück zur Artikelsuche...', false, false);
        await sleep(600);
        await returnToWarenrueckgabeSearch();
    }

    async function handleReturnListPage() {
        confirmCleanSearchPage();
        saveFlowState({ active: true, step: 'returnList', searchSubmitted: false, orderHandled: false });
        fillGrundFields();
        grundFilled = allGrundFieldsFilled();
        saveFlowState({ grundFilled: grundFilled });
        ensurePrintButton();
        bindAbsendenButton();
        hideHud();
    }

    function shouldAutoRun(page) {
        if (page === 'submitDone') return true;
        if (page === 'returnList') return true;
        if (page === 'search' || page === 'articlePick' || page === 'orders') return true;
        if (isFlowActive()) return true;
        return false;
    }

    async function runPageFlow(force) {
        if (abortMission) return;
        if (isPendingReload()) {
            isRunning = false;
            return;
        }
        if (expectsCleanSearchPage() && isOrdersPage()) {
            isRunning = false;
            await forceCleanSearchReload('Seite wird neu geladen...');
            return;
        }
        if (isRunning && !force) return;
        if (orderProcessing && !force) return;
        if (ordersFlowActive && !force) return;
        if (!isRetoureContext()) return;
        var page = detectPage();
        if (page === 'search' && isOrdersPage()) page = 'orders';
        if (page === 'articlePick' && isOrdersPage()) page = 'orders';
        if (page === 'unknown') return;
        if (!shouldAutoRun(page) && page !== 'search' && page !== 'articlePick' && page !== 'returnList' && page !== 'submitDone') return;
        var stepKey = page + '|' + (sessionStorage.getItem(ACTIVE_FLOW_KEY) || '') + '|' + (sessionStorage.getItem(EXCLUSION_OK_KEY) || '') + '|' + (searchSubmitted ? '1' : '') + '|' + (orderHandledForSearch ? 'done' : '');
        if (!force && lastHandledStep === stepKey) return;
        if (orderHandledForSearch && (page === 'orders' || (page === 'search' && isOrdersPage()))) {
            reloadForNextProduct();
            return;
        }
        isRunning = true;
        try {
            await sleep(PAGE_WAIT_MS);
            if (abortMission) return;
            if (page === 'articlePick') {
                lastHandledStep = stepKey;
                await handleArticlePickPage();
                return;
            }
            if (page === 'search') {
                var currentInput = findKnollSearchInput();
                if (currentInput) bindSearchInput(currentInput);
                if (currentInput && !normalizeText(currentInput.value) && !isFlowActive()) searchSubmitted = false;
                if (isOrdersPage()) {
                    if (orderHandledForSearch || isPendingReload()) {
                        reloadForNextProduct();
                        return;
                    }
                    lastHandledStep = '';
                    await handleOrdersPage();
                    return;
                }
                if (searchSubmitted) {
                    lastHandledStep = stepKey;
                    await handleArticlePickPage();
                    return;
                }
                if ((searchListenersReady && !force) || shouldPauseWatcher()) {
                    if (!isFlowActive()) {
                        isRunning = false;
                        return;
                    }
                }
                lastHandledStep = stepKey;
                isRunning = false;
                saveFlowState({ active: true, step: 'search', searchSubmitted: false });
                ensureSearchListeners(false);
                return;
            }
            if (page === 'excluded') {
                lastHandledStep = stepKey;
                await handleExcludedPage();
                return;
            }
            if (page === 'orders') {
                if (orderHandledForSearch || isPendingReload()) {
                    reloadForNextProduct();
                    return;
                }
                lastHandledStep = stepKey;
                await handleOrdersPage();
                return;
            }
            if (page === 'submitDone') {
                lastHandledStep = stepKey;
                await handleSubmitDonePage();
                return;
            }
            if (page === 'returnList') {
                lastHandledStep = stepKey + '|' + findGrundFields().length + '|' + (allGrundFieldsFilled() ? '1' : '0');
                await handleReturnListPage();
                return;
            }
        } finally {
            setTimeout(function() { isRunning = false; }, 800);
        }
    }

    function tickFlow(force) {
        runPageFlow(!!force);
    }

    function startPageWatcher() {
        if (pageWatcherStarted) return;
        pageWatcherStarted = true;
        setupPrintListeners();
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
            if (isRetoureUrl() || isFlowActive()) tickFlow(false);
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
        if (meta.length) lines.push('<div class="arasaka-hud-last-meta">' + meta.join(' · ') + '</div>');
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
        if (/fehler|error|nicht gefunden/.test(lower)) return 'error';
        if (/warnung|eingeschränkt|rücksprache/.test(lower)) return 'warn';
        if (/stop|gestoppt|fertig/.test(lower) || isEnd) return 'done';
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
        if (document.getElementById('knoll-retoure-hud-style')) return;
        var style = document.createElement('style');
        style.id = 'knoll-retoure-hud-style';
        style.textContent = ''
            + '#knoll-retoure-popup, #knoll-retoure-qty-popup, #knoll-retoure-exclusion-popup {'
            + '  position: fixed; z-index: 9999999; width: min(560px, calc(100vw - 32px)); color: #c9d1d9;'
            + '  font-family: "Segoe UI", Arial, sans-serif; border-radius: 22px; overflow: hidden;'
            + '  background: radial-gradient(circle at top left, rgba(220, 61, 61, 0.14), transparent 34%),'
            + '    radial-gradient(circle at top right, rgba(211, 86, 86, 0.1), transparent 32%),'
            + '    linear-gradient(180deg, rgba(21, 27, 35, 0.98) 0%, rgba(14, 19, 26, 0.98) 100%);'
            + '  border: 1px solid #2d3642; box-shadow: 0 24px 56px rgba(0, 0, 0, 0.52);'
            + '  backdrop-filter: blur(12px); -webkit-backdrop-filter: blur(12px); user-select: none;'
            + '}'
            + '#knoll-retoure-popup[data-tone="active"], #knoll-retoure-qty-popup[data-tone="active"] { border-color: rgba(220, 61, 61, 0.52); }'
            + '#knoll-retoure-popup[data-tone="done"] { border-color: rgba(86, 211, 100, 0.55); }'
            + '#knoll-retoure-popup[data-tone="error"], #knoll-retoure-exclusion-popup[data-tone="warn"] { border-color: rgba(248, 81, 73, 0.68); }'
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
            + '  border-radius: 999px; background: rgba(13, 17, 23, 0.7); border: 2px solid rgba(220, 61, 61, 0.38); color: #dc3d3d;'
            + '  font-size: 13px; font-weight: 900; letter-spacing: 1.2px; text-transform: uppercase; }'
            + '.arasaka-hud-message { color: #d7dee8; font-size: 16px; font-weight: 700; line-height: 1.52; white-space: normal; word-break: break-word; user-select: text; }'
            + '.arasaka-hud-footer { display: flex; justify-content: space-between; gap: 10px; margin-top: 16px; padding-top: 12px;'
            + '  border-top: 1px solid rgba(255,255,255,0.07); color: #8b949e; font-size: 12px; font-weight: 800; }'
            + '.arasaka-hud-credit { text-align: right; margin-top: 4px; color: #6e7681; font-size: 10px; font-weight: 600; letter-spacing: 0.3px; }'
            + '.arasaka-hud-last { width: 100%; margin-top: 14px; padding: 12px 14px; border-radius: 14px;'
            + '  background: rgba(220, 61, 61, 0.1); border: 1px solid rgba(220, 61, 61, 0.28); }'
            + '.arasaka-hud-last-label { color: #8b949e; font-size: 11px; font-weight: 800; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 6px; }'
            + '.arasaka-hud-last-name { color: #f0f6fc; font-size: 15px; font-weight: 800; line-height: 1.3; }'
            + '.arasaka-hud-last-meta { color: #f07070; font-size: 13px; font-weight: 800; margin-top: 4px; }'
            + '.knoll-retoure-qty-grid { display: flex; flex-wrap: wrap; gap: 12px; margin-top: 18px; }'
            + '.knoll-retoure-qty-btn { min-width: 72px; min-height: 56px; padding: 10px 18px; border-radius: 16px; border: 2px solid rgba(220, 61, 61, 0.42);'
            + '  background: rgba(13, 17, 23, 0.82); color: #f07070; font-size: 24px; font-weight: 900; cursor: pointer; transition: transform 0.14s, background 0.14s; }'
            + '.knoll-retoure-qty-btn:hover { transform: translateY(-2px); background: rgba(220, 61, 61, 0.14); }'
            + '.knoll-retoure-action-row { display: flex; flex-wrap: wrap; gap: 12px; margin-top: 18px; }'
            + '.knoll-retoure-action-btn { min-height: 48px; padding: 10px 18px; border-radius: 14px; border: 2px solid rgba(255,255,255,0.14);'
            + '  background: rgba(13, 17, 23, 0.82); color: #d7dee8; font-size: 14px; font-weight: 900; cursor: pointer; }'
            + '.knoll-retoure-action-primary { border-color: rgba(220, 61, 61, 0.55); color: #f07070; }'
            + '.knoll-retoure-hud-action { margin-top: 14px; width: 100%; min-height: 48px; border-radius: 14px; border: 2px solid rgba(220, 61, 61, 0.55);'
            + '  background: rgba(220, 61, 61, 0.12); color: #f07070; font-size: 14px; font-weight: 900; cursor: pointer; }'
            + '#knoll-retoure-print-btn, button.knoll-retoure-print-knoll {'
            + '  background: #1565a8 !important; background-color: #1565a8 !important; color: #fff !important;'
            + '  border: none !important; box-shadow: none !important; cursor: pointer; text-decoration: none !important;'
            + '  font-weight: 700 !important; }'
            + '#knoll-retoure-print-btn:hover, button.knoll-retoure-print-knoll:hover {'
            + '  background: #125589 !important; background-color: #125589 !important; color: #fff !important; }'
            + '#knoll-retoure-print-btn:focus, button.knoll-retoure-print-knoll:focus {'
            + '  outline: 2px solid rgba(21, 101, 168, 0.45) !important; outline-offset: 1px; }'
            + '.knoll-retoure-hud-actions { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 14px; }'
            + '@media print {'
            + '  #knoll-retoure-popup, #knoll-retoure-qty-popup, #knoll-retoure-exclusion-popup, #knoll-retoure-print-btn'
            + '    { display: none !important; visibility: hidden !important; }'
            + '}';
        document.head.appendChild(style);
    }

    function bindHudActions(popup, showListButton) {
        var listBtn = popup.querySelector('#knoll-retoure-go-list');
        if (listBtn) {
            listBtn.addEventListener('click', function() {
                clickZurRueckgabeliste();
            });
        }
        popup.querySelector('#knoll-retoure-hud-close').addEventListener('click', function() {
            popup.remove();
            lastHudSignature = '';
        });
    }

    function showHud(title, message, isEnd, showListButton) {
        hudEnsureStyles();
        var lastHtml = lastAddedHudHtml();
        var signature = title + '|' + message + '|' + !!isEnd + '|' + !!showListButton + '|' + lastHtml;
        var existing = document.getElementById('knoll-retoure-popup');
        if (existing && signature === lastHudSignature) return;
        lastHudSignature = signature;
        var tone = hudTone(title, message, isEnd);
        var actionHtml = showListButton
            ? '<div class="knoll-retoure-hud-actions"><button type="button" class="knoll-retoure-hud-action" id="knoll-retoure-go-list">Zur Rückgabeliste</button></div>'
            : '';
        var html = ''
            + '<div class="arasaka-hud-head" id="knoll-retoure-hud-drag">'
            + '  <div>'
            + '    <div class="arasaka-hud-eyebrow">KNOLL Retoure · ' + hudEscape(hudToneLabel(tone)) + '</div>'
            + '    <div class="arasaka-hud-title">' + hudEscape(title) + '</div>'
            + '  </div>'
            + '  <div class="arasaka-hud-actions">'
            + '    <button type="button" class="arasaka-hud-btn" id="knoll-retoure-hud-close">X</button>'
            + '  </div>'
            + '</div>'
            + '<div class="arasaka-hud-body">'
            + '  <div class="arasaka-hud-main">'
            + '    <div class="arasaka-hud-badge">' + hudEscape(hudToneLabel(tone)) + '</div>'
            + '    <div class="arasaka-hud-message">' + hudMessageHtml(message) + '</div>'
            + lastHtml
            + '  </div>'
            + actionHtml
            + '  ' + hudFooterHtml()
            + '</div>';
        if (existing) {
            existing.setAttribute('data-tone', tone);
            existing.innerHTML = html;
            bindHudActions(existing, showListButton);
            hudEnableDrag(existing, existing.querySelector('#knoll-retoure-hud-drag'));
            return;
        }
        var popup = document.createElement('div');
        popup.id = 'knoll-retoure-popup';
        popup.setAttribute('data-tone', tone);
        popup.innerHTML = html;
        document.body.appendChild(popup);
        hudApplyPosition(popup);
        hudEnableDrag(popup, popup.querySelector('#knoll-retoure-hud-drag'));
        bindHudActions(popup, showListButton);
    }

    function bootAttempt() {
        restoreFlowState();
        pendingReload = sessionStorage.getItem(PENDING_RELOAD_KEY) === '1';
        if (!isRetoureUrl() && !isRetoureContext()) return false;
        var page = detectPage();
        if (page === 'returnList') {
            confirmCleanSearchPage();
        } else if (expectsCleanSearchPage() && isOrdersPage()) {
            forceCleanSearchReload('Seite wird neu geladen...');
            return true;
        } else if (!isOrdersPage() && isFlowActive() && loadFlowState().step === 'search') {
            confirmCleanSearchPage();
        }
        if (page === 'search' && isOrdersPage()) page = 'orders';
        if (page === 'articlePick' && isOrdersPage()) page = 'orders';
        if (orderHandledForSearch && isOrdersPage()) {
            orderHandledForSearch = false;
            reloadForNextProduct();
            return true;
        }
        if (page === 'orders' || page === 'returnList' || page === 'submitDone' || page === 'articlePick' || isFlowActive()) {
            var bootState = loadFlowState();
            if (!(page === 'orders' && bootState.orderHandled) && page !== 'submitDone') {
                saveFlowState({ active: true, step: page === 'unknown' ? bootState.step || 'search' : page });
            }
        }
        startPageWatcher();
        if (page === 'orders') {
            if (!expectsCleanSearchPage()) {
                showHud('KNOLL RETOURE LÄUFT', 'Wähle neuesten Auftrag...', false, true);
            }
        } else if (page === 'returnList') {
            hideHud();
        } else if (page === 'submitDone') {
            showHud('KNOLL RETOURE BEREIT', 'Absenden bestätigt — zurück zur Artikelsuche...', false, false);
        } else if ((page === 'articlePick' || searchSubmitted) && !isOrdersPage()) {
            showHud('KNOLL RETOURE LÄUFT', 'Seite neu geladen — Artikel wird ausgewählt...', false, true);
        } else if (page === 'search') {
            if (isFlowActive() && !isOrdersPage()) {
                showHud('KNOLL RETOURE BEREIT', 'EAN scannen oder Artikelnummer tippen — Enter bestätigt, Scanner startet automatisch nach Pause.', false, true);
            }
            scheduleSearchAutoFocus(true);
        }
        tickFlow(true);
        return true;
    }

    function boot() {
        if (bootAttempt()) return;
        if (!isRetoureUrl() && !isFlowActive()) return;
        var attempts = 0;
        var retry = setInterval(function() {
            attempts++;
            if (bootAttempt() || attempts >= 60) clearInterval(retry);
        }, 500);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', boot);
    } else {
        boot();
    }
})();
