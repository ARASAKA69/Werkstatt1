// ==UserScript==
// @name         AAG Retoure Bot
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Sucht AAG Retouren per Artikelnummer oder EAN, nimmt die letzte Lieferung und schickt den Warenkorb ab
// @author       ARASAKA
// @match        https://portal.allianceautomotive.de/*
// @run-at       document-start
// @updateURL    https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/Warenr%C3%BCckgabe/aag-retoure-automation.user.js
// @downloadURL  https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/Warenr%C3%BCckgabe/aag-retoure-automation.user.js
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    var BOT_VERSION = '1.0';
    var GRUND = 'Falscher Artikel';
    var HUD_POS_KEY = 'aag_retoure_hud_position';
    var QUEUE_KEY = 'aag_retoure_queue';
    var PENDING_KEY = 'aag_retoure_pending';
    var FINISH_KEY = 'aag_retoure_finish';
    var ADDED_KEY = 'aag_retoure_added';
    var EAN_MAP_KEY = 'aag_retoure_ean_map';
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
    var apiPaths = null;
    var lastPath = location.pathname;

    patchTraffic();
    hookHistory();

    function loadEanMap() {
        try {
            var data = JSON.parse(sessionStorage.getItem(EAN_MAP_KEY) || '{}');
            return data && typeof data === 'object' ? data : {};
        } catch (e) {
            return {};
        }
    }

    function rememberEan(ean, article) {
        var key = onlyDigits(ean);
        var art = usableArticle(article, key);
        if (!key || !art) return;
        eanMap[key] = art;
        var keys = Object.keys(eanMap);
        if (keys.length > 300) delete eanMap[keys[0]];
        try { sessionStorage.setItem(EAN_MAP_KEY, JSON.stringify(eanMap)); } catch (e) {}
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

    function deliveryCards() {
        var root = resultRoot();
        var nodes = root.querySelectorAll('tr, [role="row"], mat-expansion-panel, li, div');
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
        for (var i = 0; i < 8 && el; i++) {
            var raw = el.innerText || '';
            if (raw.length < 500 && normArt(raw).indexOf(want) !== -1 && lineHasArticle(el, article)) return el;
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
        var raw = button.innerText || '';
        if (fold(raw).indexOf('zurueckgeben') === -1) return false;
        if (/[−–]|remove/i.test(raw)) return false;
        if (looksFilled(button) && !/[+]/.test(raw)) return false;
        return true;
    }

    function returnButtonsNear(card, article) {
        var roots = [card, card.nextElementSibling, card.parentElement].filter(Boolean);
        var seen = [];
        roots.forEach(function (root) {
            Array.prototype.forEach.call(root.querySelectorAll('button'), function (button) {
                if (seen.indexOf(button) !== -1) return;
                if (fold(button.innerText || '').indexOf('zurueckgeben') === -1) return;
                if (!lineOf(button, article)) return;
                seen.push(button);
            });
        });
        return seen;
    }

    function cartBadge() {
        var tab = findButton(/warenkorb/);
        if (!tab) return 0;
        var root = tab.parentElement || tab;
        var match = normSpace(root.innerText || '').match(/(\d+)/);
        return match ? Number(match[1]) : 0;
    }

    function cartHasItems() {
        if (pageKind() !== 'cart' || !document.body) return false;
        if (/LS\d{6,}/i.test(document.body.innerText || '')) return true;
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
        var xhrProto = XMLHttpRequest.prototype;
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
        if (window.fetch && !window.fetch.__aagRetoureFetch) {
            var originalFetch = window.fetch;
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
            window.fetch = wrapped;
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

    async function resolveCode(code) {
        if (!looksLikeEan(code)) return { code: code, eanFailed: false };
        setStatus('EAN ' + code + ' wird aufgelöst...', 'active');
        var art = eanMap[code] || await lookupEan(code);
        if (art) {
            setStatus('EAN gehört zu ' + art, 'active');
            return { code: art, eanFailed: false };
        }
        return { code: code, eanFailed: true };
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

    async function addFromCard(card, article) {
        var buttons = returnButtonsNear(card, article);
        if (!buttons.length) {
            var open = expander(card);
            try { card.scrollIntoView({ block: 'center' }); } catch (e) {}
            if (open && open.getAttribute('aria-expanded') !== 'true') open.click();
            var appeared = await waitFor(function () {
                return returnButtonsNear(card, article).length ? true : null;
            }, 2500);
            if (!appeared) return 'miss';
            buttons = returnButtonsNear(card, article);
        }
        var addButton = null;
        var line = null;
        for (var i = 0; i < buttons.length; i++) {
            var row = lineOf(buttons[i], article);
            var qty = returnableQty(row);
            if (qty === 0) continue;
            if (!isAddButton(buttons[i])) continue;
            addButton = buttons[i];
            line = row;
            break;
        }
        if (!addButton) return buttons.length ? 'already' : 'miss';
        var before = cartBadge();
        addButton.click();
        var proof = await waitFor(function () {
            if (cartBadge() > before) return 'added';
            if (!isAddButton(addButton)) return 'flipped';
            return null;
        }, 4000);
        if (cartBadge() < before) {
            addButton.click();
            return 'fail';
        }
        if (!proof) return 'fail';
        return { ok: true, name: lineName(line, article), ls: lsOf(card), date: dateOf(card) };
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
        var suchen = findButton(/suchen$/, /rueckgaben/);
        if (!suchen) {
            setStatus('Button Suchen nicht gefunden.', 'error');
            return;
        }
        var beforeSig = deliveryCards().map(lsOf).join('|');
        var sawSpinner = false;
        suchen.click();
        focusScan();
        var result = await waitFor(function () {
            var spinning = !!document.querySelector('mat-progress-spinner, mat-spinner, .mat-mdc-progress-spinner');
            if (spinning) sawSpinner = true;
            if (spinning) return null;
            var cards = deliveryCards();
            var sig = cards.map(lsOf).join('|');
            if (!sig) return null;
            if (sig !== beforeSig || sawSpinner) return cards;
            return null;
        }, 8000);
        var cards = result || deliveryCards();
        if (!cards.length) {
            if (eanFailed) setStatus('EAN nicht aufgelöst. Artikelnummer eintippen, zum Beispiel J1340842, und Enter.', 'error');
            else setStatus(code + ' nicht gefunden. Letztes Jahr steht oben, sonst die Nummer prüfen.', 'error');
            focusScan();
            return;
        }
        var already = false;
        for (var i = 0; i < cards.length && i < 8; i++) {
            if (abortMission) return;
            setStatus('Lieferung ' + (lsOf(cards[i]) || '') + ' vom ' + (dateOf(cards[i]) || '?') + '...', 'active');
            var added = await addFromCard(cards[i], code);
            if (added && added.ok) {
                rememberAdded({ article: code, name: added.name || '', ls: added.ls || '' });
                setStatus(code + ' liegt im Warenkorb. Nächste Nummer, oder Drucken und absenden.', 'done');
                focusScan();
                return;
            }
            if (added === 'already') already = true;
        }
        if (already) setStatus(code + ' liegt schon im Warenkorb.', 'warn');
        else if (eanFailed) setStatus('EAN nicht aufgelöst. Artikelnummer eintippen und Enter.', 'error');
        else setStatus(code + ' ist in den Lieferungen nicht rückgabefähig.', 'error');
        focusScan();
    }

    async function addArticle(raw) {
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
            if (pageKind() !== 'search') {
                sessionStorage.setItem(PENDING_KEY, code);
                leaving = true;
                skipPump = true;
                goSearch();
                return;
            }
            var resolved = await resolveCode(code);
            if (abortMission) return;
            await runSearchAdd(resolved.code, resolved.eanFailed && resolved.code === code);
        } catch (e) {
            setStatus('Fehler beim Suchen.', 'error');
        } finally {
            busy = false;
            if (!skipPump && !abortMission && sessionStorage.getItem(FINISH_KEY) !== '1') pump();
        }
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
        addArticle(code);
    }

    function isGrundText(text) {
        if (!text || text.length > 48) return false;
        if (text.indexOf('falscher artikel') !== -1) return true;
        return text.indexOf('bitte grund') !== -1;
    }

    function grundIsSet(trigger) {
        if (!trigger) return false;
        if (trigger.tagName === 'SELECT') {
            var selected = trigger.options[trigger.selectedIndex];
            return !!(selected && fold(selected.text).indexOf('falscher artikel') !== -1);
        }
        return fold(trigger.innerText || '').indexOf('falscher artikel') !== -1 && fold(trigger.innerText || '').indexOf('bitte grund') === -1;
    }

    async function chooseGrund(trigger) {
        if (grundIsSet(trigger)) return true;
        if (trigger.tagName === 'SELECT') {
            for (var i = 0; i < trigger.options.length; i++) {
                if (fold(trigger.options[i].text).indexOf('falscher artikel') !== -1) {
                    trigger.selectedIndex = i;
                    trigger.dispatchEvent(new Event('input', { bubbles: true }));
                    trigger.dispatchEvent(new Event('change', { bubbles: true }));
                    return true;
                }
            }
            return true;
        }
        trigger.click();
        var option = await waitFor(findGrundOption, 2000);
        if (!option) return false;
        option.click();
        await sleep(250);
        return grundIsSet(trigger);
    }

    function findGrundOption() {
        var overlay = document.querySelector('.cdk-overlay-container');
        var scope = overlay || document;
        var nodes = scope.querySelectorAll('mat-option, [role="option"]');
        if (!nodes.length && overlay) nodes = overlay.querySelectorAll('li, button, span, div');
        var best = null;
        var bestLen = 999;
        for (var i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (!isVisible(el)) continue;
            var text = fold(el.innerText || '');
            if (!text || text.length > 40 || text.indexOf('falscher artikel') === -1) continue;
            if (text.length < bestLen) {
                best = el;
                bestLen = text.length;
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
            var ok = await chooseGrund(pending[0]);
            if (!ok) return false;
        }
        return grundReady();
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
            if (pageKind() !== 'cart') {
                sessionStorage.setItem(FINISH_KEY, '1');
                leaving = true;
                skipPump = true;
                setStatus('Öffne den Warenkorb...', 'active');
                goCart();
                return;
            }
            sessionStorage.removeItem(FINISH_KEY);
            setStatus('Setze Grund auf ' + GRUND + '...', 'active');
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
        setTimeout(function () {
            var dialog = document.querySelector('mat-dialog-container, [role="dialog"]');
            if (!dialog) {
                setStatus('Retoure abgeschickt.', 'done');
                return;
            }
            var confirm = findByText(dialog, /^(bestaetigen|ja|ok|absenden)$/, 24);
            if (confirm) confirm.click();
            setStatus('Retoure abgeschickt.', 'done');
        }, 400);
    }

    function onFinishClick() {
        if (busy) return;
        if (pageKind() === 'cart' && grundReady()) {
            sessionStorage.removeItem(FINISH_KEY);
            printAndSendSync();
            return;
        }
        finishCart();
    }

    function continueFlow() {
        ensureHud();
        if (!onRetours() || busy) return;
        if (sessionStorage.getItem(FINISH_KEY) === '1' && pageKind() === 'cart') {
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
        if (history.pushState && !history.pushState.__aagRetoure) {
            var original = history.pushState;
            var wrapped = function () {
                var result = original.apply(this, arguments);
                setTimeout(continueFlow, 40);
                return result;
            };
            wrapped.__aagRetoure = true;
            history.pushState = wrapped;
        }
        if (history.replaceState && !history.replaceState.__aagRetoure) {
            var originalReplace = history.replaceState;
            var wrappedReplace = function () {
                var result = originalReplace.apply(this, arguments);
                setTimeout(continueFlow, 40);
                return result;
            };
            wrappedReplace.__aagRetoure = true;
            history.replaceState = wrappedReplace;
        }
        window.addEventListener('popstate', function () { setTimeout(continueFlow, 40); });
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
        var input = document.getElementById('aag-retoure-scan');
        if (!input) return;
        setTimeout(function () {
            if (!onRetours()) return;
            input.focus();
        }, 60);
    }

    function resetHud() {
        abortMission = true;
        saveQueue([]);
        sessionStorage.removeItem(PENDING_KEY);
        sessionStorage.removeItem(FINISH_KEY);
        sessionStorage.removeItem(ADDED_KEY);
        var input = document.getElementById('aag-retoure-scan');
        if (input) input.value = '';
        renderAdded();
        setStatus('HUD leer. Der Warenkorb auf der Seite bleibt.', 'neutral');
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
            + '#aag-retoure-scan{width:100%;box-sizing:border-box;margin-top:10px;min-height:44px;border-radius:12px;border:1px solid #3a4452;background:#0d1117;color:#f0f6fc;font-size:16px;font-weight:700;padding:0 12px}'
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
            + '.aag-foot{display:flex;justify-content:space-between;gap:10px;margin-top:12px;padding-top:10px;border-top:1px solid rgba(255,255,255,.07);color:#8b949e;font-size:12px;font-weight:800}'
            + '.aag-credit{text-align:right;margin-top:4px;color:#6e7681;font-size:10px;font-weight:600}'
            + '@media print{#aag-retoure-hud,#aag-lieferschein-sync-status{display:none !important}}';
        document.documentElement.appendChild(style);
    }

    function ensureHud() {
        if (!document.body) return;
        var hud = document.getElementById('aag-retoure-hud');
        if (!onRetours()) {
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
            + '  <div class="aag-actions">'
            + '    <button type="button" id="aag-retoure-cart">Warenkorb</button>'
            + '    <button type="button" id="aag-retoure-finish">Drucken und absenden</button>'
            + '  </div>'
            + '  <div id="aag-retoure-badge">BEREIT</div>'
            + '  <div id="aag-retoure-message">Artikelnummer scannen oder eintippen. Enter nimmt die letzte Lieferung in den Warenkorb.</div>'
            + '  <div id="aag-retoure-last"></div>'
            + '  <div class="aag-foot"><span>Bot v' + BOT_VERSION + '</span><span>ESC Stop</span></div>'
            + '  <div class="aag-credit">by Arasaka</div>'
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
            setStatus('Warenkorb offen. Grund, dann Drucken und absenden.', 'neutral');
            goCart();
        });
        document.getElementById('aag-retoure-finish').addEventListener('click', onFinishClick);
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
        ensureHud();
        window.addEventListener('beforeprint', hideForPrint);
        window.addEventListener('afterprint', showAfterPrint);
        document.addEventListener('keydown', function (e) {
            if (e.key !== 'Escape' || !onRetours()) return;
            abortMission = true;
            saveQueue([]);
            sessionStorage.removeItem(PENDING_KEY);
            sessionStorage.removeItem(FINISH_KEY);
            setStatus('Stop. Nächster Scan geht wieder.', 'warn');
            setTimeout(function () { abortMission = false; }, 50);
        }, true);
        setInterval(function () {
            var pathChanged = location.pathname !== lastPath;
            lastPath = location.pathname;
            if (pathChanged || sessionStorage.getItem(FINISH_KEY) === '1' || sessionStorage.getItem(PENDING_KEY)) continueFlow();
            else ensureHud();
        }, 700);
        continueFlow();
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
    else boot();
})();
