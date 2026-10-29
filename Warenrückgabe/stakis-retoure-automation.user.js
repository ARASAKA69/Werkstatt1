// ==UserScript==
// @name         STAKIS Warenrückgabe Retoure Bot
// @namespace    http://tampermonkey.net/
// @version      1.0.12
// @description  Sucht STAKIS Lieferungen per EAN, nimmt den ältesten Beleg und schreibt ihn ins Blatt STA
// @author       ARASAKA
// @match        *://*.stakis.de/*
// @match        *://stakis.de/*
// @run-at       document-idle
// @updateURL    https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/Warenr%C3%BCckgabe/stakis-retoure-automation.user.js
// @downloadURL  https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/Warenr%C3%BCckgabe/stakis-retoure-automation.user.js
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_xmlhttpRequest
// @connect      script.google.com
// @connect      script.googleusercontent.com
// ==/UserScript==

(function () {
    'use strict';

    var BOT_VERSION = '1.0.12';
    var WORKER_KEY = 'sta_retoure_worker';
    var SHEET_URL_KEY = 'sta_retoure_sheet_url';
    var COUNT_KEY = 'sta_retoure_count';
    var LAST_KEY = 'sta_retoure_last';
    var GEAR_KEY = 'sta_retoure_gear';
    var MIN_SCAN = 8;
    var MAX_SCAN = 14;
    var SCAN_IDLE_MS = 700;

    var busy = false;
    var abortMission = false;
    var scanTimer = null;
    var boundInput = null;
    var waitPick = null;
    var sheetBusy = '';

    function sleep(ms) {
        return new Promise(function (resolve) { setTimeout(resolve, ms); });
    }

    function gmGet(key, fallback) {
        try { return GM_getValue(key, fallback); } catch (e) { return fallback; }
    }

    function gmSet(key, value) {
        try { GM_setValue(key, value); } catch (e) {}
    }

    function textOf(el) {
        return String((el && (el.innerText || el.textContent)) || '').replace(/\s+/g, ' ').trim();
    }

    function visible(el) {
        if (!el || !el.getBoundingClientRect) return false;
        var rect = el.getBoundingClientRect();
        if (rect.width < 2 || rect.height < 2) return false;
        var style = window.getComputedStyle(el);
        return style.display !== 'none' && style.visibility !== 'hidden';
    }

    function pageWindow() {
        try {
            if (typeof unsafeWindow !== 'undefined' && unsafeWindow) return unsafeWindow;
        } catch (e) {}
        return window;
    }

    function makeMouse(type, opts) {
        var Ctor = type.indexOf('pointer') === 0 && typeof PointerEvent === 'function' ? PointerEvent : MouseEvent;
        var full = Object.assign({ view: window }, opts);
        try { return new Ctor(type, full); } catch (e) {}
        try { return new Ctor(type, opts); } catch (e2) {}
        try { return new MouseEvent(type, opts); } catch (e3) {}
        return null;
    }

    function fireMouse(el, types) {
        if (!el) return;
        var rect = el.getBoundingClientRect ? el.getBoundingClientRect() : { left: 0, top: 0, width: 0, height: 0 };
        var opts = { bubbles: true, cancelable: true, composed: true, clientX: rect.left + rect.width / 2, clientY: rect.top + rect.height / 2, button: 0, buttons: 1, pointerType: 'mouse', isPrimary: true };
        types.forEach(function (type) {
            var ev = makeMouse(type, opts);
            if (!ev) return;
            try { el.dispatchEvent(ev); } catch (e) {}
        });
    }

    function fireClick(el) {
        if (!el) return;
        fireMouse(el, ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']);
        try { if (typeof el.click === 'function') el.click(); } catch (e) {}
    }

    function invokeReactClick(el) {
        var node = el;
        var guard = 0;
        while (node && guard < 4) {
            var props = reactProps(node);
            if (props && typeof props.onClick === 'function') {
                try {
                    var Ctor = pageWindow().MouseEvent || MouseEvent;
                    props.onClick({ currentTarget: node, target: node, button: 0, nativeEvent: new Ctor('click', { bubbles: true }), preventDefault: function () {}, stopPropagation: function () {} });
                    return true;
                } catch (e) {}
            }
            if (props && typeof props.onMouseDown === 'function') {
                try {
                    props.onMouseDown({ button: 0, currentTarget: node, target: node, preventDefault: function () {}, stopPropagation: function () {} });
                    return true;
                } catch (e2) {}
            }
            node = node.parentElement;
            guard++;
        }
        return false;
    }

    function pressControl(el) {
        if (!el) return;
        if (!invokeReactClick(el)) fireClick(el);
    }

    function setNativeValue(input, value) {
        var proto = input.tagName === 'SELECT' ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
        var desc = Object.getOwnPropertyDescriptor(proto, 'value');
        if (desc && desc.set) desc.set.call(input, value);
        else input.value = value;
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
    }

    function isLieferungen() {
        return /\/cis\/vouchers(\/|$)/i.test(location.pathname);
    }

    function isReturns() {
        return /\/cis\/returns(\/|$)/i.test(location.pathname);
    }

    function isBotPage() {
        return /(^|\.)stakis\.de$/i.test(location.hostname) && (isLieferungen() || isReturns());
    }

    function workRoot() {
        var match = location.pathname.match(/^(.*\/cis)(?:\/|$)/i);
        return match ? match[1] : '';
    }

    function clickTab(name) {
        var want = String(name || '').toUpperCase();
        var nodes = document.querySelectorAll('[role="tab"], a, button, div, span');
        var best = null;
        var bestArea = Infinity;
        var i;
        for (i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (!visible(el)) continue;
            if (el.closest && el.closest('#sta-retoure-bot')) continue;
            var label = textOf(el).replace(/\s+\d+\s*$/, '').trim().toUpperCase();
            if (label !== want) continue;
            var rect = el.getBoundingClientRect();
            var area = rect.width * rect.height;
            if (area > 0 && area < bestArea) {
                best = el;
                bestArea = area;
            }
        }
        if (!best) return false;
        var target = (best.closest && best.closest('[role="tab"], a, button')) || best;
        pressControl(target);
        fireClick(target);
        return true;
    }

    function goLieferungen() {
        var root = workRoot();
        var targetUrl = (root || '') + '/vouchers/3';
        if (clickTab('LIEFERUNGEN')) {
            setTimeout(function () {
                if (!isLieferungen() && root) location.assign(targetUrl);
            }, 400);
            return;
        }
        if (root) location.assign(targetUrl);
    }

    function goReturns() {
        var root = workRoot();
        var targetUrl = (root || '') + '/returns';
        if (clickTab('WARENRÜCKGABE')) {
            setTimeout(function () {
                if (!isReturns() && root) location.assign(targetUrl);
            }, 400);
            return;
        }
        if (root) location.assign(targetUrl);
    }

    function basketCount() {
        return parseInt(sessionStorage.getItem(COUNT_KEY) || '0', 10) || 0;
    }

    function setBasketCount(value) {
        sessionStorage.setItem(COUNT_KEY, String(value || 0));
    }

    function loadLast() {
        try { return JSON.parse(sessionStorage.getItem(LAST_KEY) || 'null'); } catch (e) { return null; }
    }

    function saveLast(info) {
        sessionStorage.setItem(LAST_KEY, JSON.stringify(info || null));
    }

    function workerName() {
        return String(gmGet(WORKER_KEY, '') || '').trim();
    }

    function sheetUrl() {
        var saved = String(gmGet(SHEET_URL_KEY, '') || '').trim();
        if (/^https:\/\/script\.google\.com\/.+\/exec\/?$/i.test(saved)) return saved;
        return '';
    }

    function hudEscape(value) {
        return String(value == null ? '' : value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function ensureStyles() {
        if (document.getElementById('sta-retoure-style')) return;
        var style = document.createElement('style');
        style.id = 'sta-retoure-style';
        style.textContent = ''
            + '#sta-retoure-bot{position:fixed;z-index:2147483646;top:84px;right:16px;width:min(380px,calc(100vw - 24px));color:#c9d1d9;font-family:"Segoe UI",Arial,sans-serif;border-radius:22px;overflow:hidden;'
            + 'background:radial-gradient(circle at top left,rgba(59,130,246,.18),transparent 34%),linear-gradient(180deg,rgba(21,27,35,.98),rgba(14,19,26,.98));border:1px solid #2d3642;box-shadow:0 24px 56px rgba(0,0,0,.52);user-select:none}'
            + '#sta-retoure-bot *{box-sizing:border-box}'
            + '#sta-retoure-bot[data-tone="active"]{border-color:rgba(59,130,246,.55)}'
            + '#sta-retoure-bot[data-tone="done"]{border-color:rgba(86,211,100,.55)}'
            + '#sta-retoure-bot[data-tone="error"]{border-color:rgba(248,81,73,.68)}'
            + '#sta-retoure-bot button{font-family:inherit;cursor:pointer}'
            + '#sta-retoure-bot .head{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:14px 14px 12px;cursor:move;border-bottom:1px solid #242d39}'
            + '#sta-retoure-bot .eyebrow{color:#8b949e;font-size:11px;font-weight:800;letter-spacing:1.4px;text-transform:uppercase}'
            + '#sta-retoure-bot .title{color:#f0f6fc;font-size:18px;font-weight:900;letter-spacing:.4px}'
            + '#sta-retoure-bot .title span{color:#3b82f6}'
            + '#sta-retoure-bot .tools{display:flex;gap:6px}'
            + '#sta-retoure-bot .icon{min-width:36px;height:34px;border-radius:12px;border:1px solid rgba(255,255,255,.12);background:rgba(13,17,23,.72);color:#c9d1d9;font-weight:900}'
            + '#sta-retoure-bot.min .body{display:none}'
            + '#sta-retoure-bot .body{padding:14px;max-height:calc(100vh - 160px);overflow-y:auto;overflow-x:hidden;overscroll-behavior:contain}'
            + '#sta-retoure-bot .badge{display:inline-flex;min-height:28px;align-items:center;padding:4px 10px;border-radius:999px;border:2px solid rgba(59,130,246,.45);color:#3b82f6;font-size:11px;font-weight:900;letter-spacing:1px}'
            + '#sta-retoure-bot[data-tone="done"] .badge{border-color:rgba(86,211,100,.5);color:#56d364}'
            + '#sta-retoure-bot[data-tone="error"] .badge{border-color:rgba(248,81,73,.55);color:#f85149}'
            + '#sta-retoure-bot .status{margin:8px 0 10px;color:#d7dee8;font-size:14px;font-weight:700;line-height:1.4;user-select:text}'
            + '#sta-retoure-bot .last{border:1px solid rgba(59,130,246,.28);border-radius:16px;padding:12px;margin:0 0 10px;background:rgba(13,30,60,.42);font-size:13px;line-height:1.4}'
            + '#sta-retoure-bot .last em{display:block;color:#8b949e;font-style:normal;font-size:10px;font-weight:800;letter-spacing:1px;margin-bottom:4px}'
            + '#sta-retoure-bot .last div:first-of-type{color:#f0f6fc;font-weight:800}'
            + '#sta-retoure-bot .stock{color:#60a5fa;font-weight:900}'
            + '#sta-retoure-bot button.act,#sta-retoure-bot button.ghost{width:100%;min-height:40px;border-radius:12px;padding:10px 12px;font-size:13px;font-weight:900;margin-top:8px}'
            + '#sta-retoure-bot button.act{border:none;background:#3b82f6;color:#06122a}'
            + '#sta-retoure-bot button.act:hover{background:#60a5fa}'
            + '#sta-retoure-bot button.act:disabled{opacity:.55;cursor:wait}'
            + '#sta-retoure-bot button.ghost{border:1px solid rgba(255,255,255,.12);background:rgba(13,17,23,.82);color:#d7dee8}'
            + '#sta-retoure-bot .row{display:flex;gap:8px}'
            + '#sta-retoure-bot .row > *{flex:1;min-width:0}'
            + '#sta-retoure-bot input{width:100%;min-height:38px;border:1px solid #3a4452;border-radius:12px;background:#0d1117;color:#f0f6fc;padding:0 12px;font:700 14px "Segoe UI",Arial,sans-serif;user-select:text}'
            + '#sta-retoure-bot input:focus{outline:none;border-color:#3b82f6;box-shadow:0 0 0 3px rgba(59,130,246,.18)}'
            + '#sta-retoure-bot label{display:block;margin:8px 0 4px;color:#8b949e;font-size:10px;font-weight:800;letter-spacing:.6px;text-transform:uppercase}'
            + '#sta-retoure-bot .picks{display:flex;flex-direction:column;gap:6px;margin-top:8px}'
            + '#sta-retoure-bot .picks button{text-align:left;min-height:36px;border:1px solid rgba(59,130,246,.45);border-radius:10px;background:rgba(59,130,246,.12);color:#93c5fd;padding:8px 10px;font-size:13px;font-weight:800}'
            + '#sta-retoure-bot .picks button.pick{display:flex;flex-direction:column;gap:3px}'
            + '#sta-retoure-bot .picks button.pick b{color:#f0f6fc;font-size:13px}'
            + '#sta-retoure-bot .picks button.pick span{color:#93c5fd;font-size:11px;font-weight:700}'
            + '#sta-retoure-bot .picks button.pick:hover{background:rgba(59,130,246,.24)}'
            + '#sta-retoure-bot .picks button.pick.grey{border-color:rgba(248,158,73,.55);background:rgba(248,158,73,.1)}'
            + '#sta-retoure-bot .picks button.pick.grey span{color:#f0b86e}'
            + '.sta-picked-row{outline:3px solid #3b82f6 !important;outline-offset:-3px}'
            + '#sta-retoure-bot .picks button.ghost{border-color:rgba(255,255,255,.12);background:rgba(13,17,23,.82);color:#d7dee8}'
            + '#sta-retoure-bot .gear{margin-top:10px;padding-top:8px;border-top:1px solid rgba(255,255,255,.07)}'
            + '#sta-retoure-bot .foot{display:flex;justify-content:space-between;gap:10px;margin-top:12px;padding-top:10px;border-top:1px solid rgba(255,255,255,.07);color:#8b949e;font-size:12px;font-weight:800}'
            + '#sta-retoure-bot .credit{text-align:right;margin-top:4px;color:#6e7681;font-size:10px;font-weight:600}'
            + '@media print{#sta-retoure-bot{display:none !important}}';
        document.documentElement.appendChild(style);
    }

    function hudEl() {
        return document.getElementById('sta-retoure-bot');
    }

    function paintHud(message, tone) {
        if (!isBotPage()) {
            var old = hudEl();
            if (old) old.remove();
            return;
        }
        ensureStyles();
        var panel = hudEl();
        if (!panel) {
            panel = document.createElement('section');
            panel.id = 'sta-retoure-bot';
            document.documentElement.appendChild(panel);
            var drag = null;
            panel.addEventListener('pointerdown', function (event) {
                if (!event.target.closest('.head') || event.target.closest('button')) return;
                var rect = panel.getBoundingClientRect();
                drag = { dx: event.clientX - rect.left, dy: event.clientY - rect.top };
                panel.style.left = rect.left + 'px';
                panel.style.top = rect.top + 'px';
                panel.style.right = 'auto';
                function move(e) {
                    if (!drag) return;
                    panel.style.left = (e.clientX - drag.dx) + 'px';
                    panel.style.top = (e.clientY - drag.dy) + 'px';
                }
                function up() {
                    drag = null;
                    document.removeEventListener('pointermove', move);
                }
                document.addEventListener('pointermove', move);
                document.addEventListener('pointerup', up, { once: true });
            });
        }
        var last = loadLast();
        var lastHtml = '';
        if (last && (last.name || last.article)) {
            lastHtml = '<div class="last"><em>ZULETZT</em><div>' + hudEscape(last.name || last.article) + '</div><div>'
                + hudEscape([last.article, last.beleg, last.stockId ? '' : 'ohne Stock-ID'].filter(Boolean).join(' · '))
                + (last.stockId ? ' · <span class="stock">' + hudEscape(last.stockId) + '</span>' : '')
                + '</div>'
                + (last.sheet ? '<div>' + hudEscape(last.sheet) + '</div>' : '')
                + '</div>';
        }
        var count = basketCount();
        var actions = '';
        if (isLieferungen() && count > 0) {
            actions += '<button type="button" class="act" data-act="returns">Zur Warenrückgabe (' + count + ')</button>';
        }
        if (isReturns()) {
            actions += '<button type="button" class="act" data-act="reasons">Falscher Artikel setzen</button>';
            actions += '<button type="button" class="act" data-act="print">Drucken und absenden</button>';
            actions += '<button type="button" class="ghost" data-act="back">Zurück zur Suche</button>';
        }
        var gearOpen = sessionStorage.getItem(GEAR_KEY) === '1';
        var gear = gearOpen
            ? '<div class="gear"><label>Dein Name</label><input id="sta-worker" value="' + hudEscape(workerName()) + '">'
                + '<label>Sheet Web-App</label><input id="sta-sheet" placeholder="https://script.google.com/.../exec" value="' + hudEscape(sheetUrl()) + '">'
                + '<button type="button" class="ghost" data-act="save">Speichern</button></div>'
            : '';
        var prevSlot = panel.querySelector('#sta-slot');
        var slotKids = prevSlot ? Array.prototype.slice.call(prevSlot.childNodes) : [];
        var toneKey = tone === 'bad' ? 'error' : tone === 'ok' ? 'done' : busy ? 'active' : 'neutral';
        var toneLabel = toneKey === 'error' ? 'FEHLER' : toneKey === 'done' ? 'OK' : toneKey === 'active' ? 'LÄUFT' : 'BEREIT';
        panel.setAttribute('data-tone', toneKey);
        var minimized = panel.classList.contains('min');
        panel.innerHTML = ''
            + '<div class="head"><div><div class="eyebrow">STAKIS Retoure · ' + (isReturns() ? 'Warenrückgabe' : 'Lieferungen') + '</div>'
            + '<div class="title">STAKIS <span>RETOURE</span> v' + BOT_VERSION + '</div></div>'
            + '<div class="tools"><button type="button" class="icon" data-act="gear">⚙</button><button type="button" class="icon" data-act="min">' + (minimized ? '+' : '–') + '</button></div></div>'
            + '<div class="body"><div class="badge">' + toneLabel + '</div><div class="status">' + hudEscape(message || '') + '</div>'
            + lastHtml
            + '<div id="sta-slot"></div>'
            + actions
            + gear
            + '<div class="foot"><span>Bot v' + BOT_VERSION + '</span><span>ESC Stop</span></div>'
            + '<div class="credit">by Arasaka</div></div>';
        if (slotKids.length) {
            var nextSlot = panel.querySelector('#sta-slot');
            slotKids.forEach(function (child) { nextSlot.appendChild(child); });
        }
        panel.querySelector('[data-act="min"]').addEventListener('click', function () {
            panel.classList.toggle('min');
            this.textContent = panel.classList.contains('min') ? '+' : '–';
        });
        panel.querySelector('[data-act="gear"]').addEventListener('click', function () {
            sessionStorage.setItem(GEAR_KEY, gearOpen ? '0' : '1');
            paintHud(message, tone);
        });
        var returns = panel.querySelector('[data-act="returns"]');
        if (returns) returns.addEventListener('click', function () { goReturns(); });
        var reasons = panel.querySelector('[data-act="reasons"]');
        if (reasons) reasons.addEventListener('click', function () { runReasons(); });
        var print = panel.querySelector('[data-act="print"]');
        if (print) print.addEventListener('click', function () { runPrint(); });
        var back = panel.querySelector('[data-act="back"]');
        if (back) back.addEventListener('click', function () { finishCycle('Zurück zur Suche.'); });
        var save = panel.querySelector('[data-act="save"]');
        if (save) save.addEventListener('click', function () {
            gmSet(WORKER_KEY, panel.querySelector('#sta-worker').value.trim());
            gmSet(SHEET_URL_KEY, panel.querySelector('#sta-sheet').value.trim());
            sessionStorage.setItem(GEAR_KEY, '0');
            paintHud('Gespeichert.', 'ok');
        });
    }

    function setSlot(html) {
        var slot = document.getElementById('sta-slot');
        if (slot) slot.innerHTML = html || '';
        return slot;
    }

    var ART_WORDS = { katalognummer: 1, artikelnummer: 1, belegnummer: 1, ean: 1, hersteller: 1 };

    function optionWord(opt) {
        return String((opt && (opt.textContent || opt.label || opt.text)) || '').replace(/\s+/g, ' ').trim();
    }

    function eanSelectOption(select) {
        return Array.prototype.filter.call(select.options || [], function (opt) {
            return /^ean$/i.test(optionWord(opt)) || /^europeanarticlenumber$/i.test(String(opt.value || ''));
        })[0] || null;
    }

    function findArtSelect() {
        var selects = document.querySelectorAll('select');
        var i;
        for (i = 0; i < selects.length; i++) {
            if (selects[i].closest && selects[i].closest('#sta-retoure-bot')) continue;
            if (eanSelectOption(selects[i])) return selects[i];
        }
        return null;
    }

    function findArtCombo() {
        var nodes = document.querySelectorAll('[role="combobox"], [aria-haspopup="listbox"], .MuiSelect-select');
        var label = findArtLabel();
        var labelRect = label ? label.getBoundingClientRect() : null;
        var best = null;
        var bestDist = Infinity;
        var i;
        for (i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (el.tagName === 'INPUT' || el.tagName === 'SELECT') continue;
            if (!visible(el)) continue;
            if (el.closest && el.closest('#sta-retoure-bot')) continue;
            if (!ART_WORDS[fieldValue(el).toLowerCase()]) continue;
            if (!labelRect) return el;
            var rect = el.getBoundingClientRect();
            var dist = Math.abs(rect.left - labelRect.left) + Math.abs(rect.top - labelRect.bottom);
            if (dist < bestDist) {
                best = el;
                bestDist = dist;
            }
        }
        return best;
    }

    function hiddenArtValue(control) {
        var box = control;
        var guard = 0;
        while (box && guard < 3) {
            var input = box.querySelector && box.querySelector('input');
            if (input && input !== control) return String(input.value || '');
            box = box.parentElement;
            guard++;
        }
        return '';
    }

    function fieldValue(el) {
        if (!el) return '';
        var copy = el.cloneNode(true);
        var menus = copy.querySelectorAll('[role="listbox"], [role="menu"], ul');
        var i;
        for (i = 0; i < menus.length; i++) menus[i].remove();
        return String(copy.textContent || '').replace(/\s+/g, ' ').trim();
    }

    function findArtLabel() {
        var nodes = document.querySelectorAll('label, span, div, p');
        var best = null;
        var bestArea = Infinity;
        var i;
        for (i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (!visible(el)) continue;
            if (el.closest && el.closest('#sta-retoure-bot')) continue;
            if (textOf(el).toLowerCase() !== 'art') continue;
            var rect = el.getBoundingClientRect();
            if (rect.width > 90) continue;
            var area = rect.width * rect.height;
            if (area < bestArea) {
                best = el;
                bestArea = area;
            }
        }
        return best;
    }

    function findArtControl() {
        var combo = findArtCombo();
        if (combo) return combo;
        var label = findArtLabel();
        if (!label) return null;
        var labelRect = label.getBoundingClientRect();
        var root = label.parentElement;
        var guard = 0;
        var best = null;
        var bestArea = Infinity;
        while (root && guard < 6) {
            var nodes = root.querySelectorAll('button, [role="button"], [role="combobox"], div, span');
            var i;
            for (i = 0; i < nodes.length; i++) {
                var el = nodes[i];
                if (el === label || (label.contains && label.contains(el))) continue;
                if (!visible(el)) continue;
                if (el.closest && (el.closest('#sta-retoure-bot') || el.closest('[role="listbox"], [role="menu"], [role="option"]'))) continue;
                if (!ART_WORDS[fieldValue(el).toLowerCase()]) continue;
                var rect = el.getBoundingClientRect();
                if (rect.width < 70 || rect.height < 16) continue;
                if (rect.top > labelRect.bottom + 90) continue;
                if (rect.bottom < labelRect.top - 30) continue;
                var area = rect.width * rect.height;
                if (area < bestArea) {
                    best = el;
                    bestArea = area;
                }
            }
            if (best) return best;
            root = root.parentElement;
            guard++;
        }
        return null;
    }

    function artOpener(el) {
        var node = el;
        var guard = 0;
        while (node && guard < 6) {
            var role = node.getAttribute ? (node.getAttribute('role') || '') : '';
            if (node.tagName === 'BUTTON' || role === 'button' || role === 'combobox' || (node.getAttribute && node.getAttribute('aria-haspopup'))) return node;
            node = node.parentElement;
            guard++;
        }
        return el;
    }

    function findEanOption() {
        var nodes = document.querySelectorAll('[role="option"], [role="menuitem"], li, div, span, button');
        var best = null;
        var bestArea = Infinity;
        var i;
        for (i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (!visible(el)) continue;
            if (el.closest && el.closest('#sta-retoure-bot')) continue;
            if (textOf(el).toLowerCase() !== 'ean') continue;
            var role = el.getAttribute ? (el.getAttribute('role') || '') : '';
            var parent = el.parentElement;
            var menuHits = 0;
            var guard = 0;
            while (parent && guard < 3 && menuHits < 3) {
                var kids = parent.children;
                var k;
                menuHits = 0;
                for (k = 0; k < kids.length; k++) {
                    if (ART_WORDS[textOf(kids[k]).toLowerCase()]) menuHits++;
                }
                if (menuHits >= 3) break;
                parent = parent.parentElement;
                guard++;
            }
            var inMenu = role === 'option' || role === 'menuitem' || el.tagName === 'LI' || menuHits >= 3 || (el.closest && el.closest('[role="listbox"], [role="menu"], ul'));
            if (!inMenu) continue;
            var rect = el.getBoundingClientRect();
            var area = rect.width * rect.height;
            if (area > 0 && area < bestArea) {
                best = el;
                bestArea = area;
            }
        }
        if (!best) return null;
        var row = best.closest && best.closest('[role="option"], [role="menuitem"], li');
        return row || best;
    }

    function pressKey(target, key) {
        if (!target) return;
        var code = key === 'ArrowDown' ? 'ArrowDown' : key === 'ArrowUp' ? 'ArrowUp' : key === 'Enter' ? 'Enter' : key;
        var keyCode = key === 'ArrowDown' ? 40 : key === 'ArrowUp' ? 38 : key === 'Enter' ? 13 : 0;
        var init = { key: key, code: code, bubbles: true, cancelable: true, view: window, keyCode: keyCode, which: keyCode };
        try { target.dispatchEvent(new KeyboardEvent('keydown', init)); } catch (e) {}
        try { target.dispatchEvent(new KeyboardEvent('keyup', init)); } catch (e2) {}
    }

    function artIsEan() {
        var select = findArtSelect();
        if (select && visible(select)) {
            var opt = eanSelectOption(select);
            return !!(opt && select.value === opt.value);
        }
        var control = findArtControl();
        if (!control) return false;
        if (fieldValue(control).toLowerCase() === 'ean') return true;
        return /^(2|ean|europeanarticlenumber)$/i.test(hiddenArtValue(control));
    }

    function findMenu(pattern) {
        var nodes = document.querySelectorAll('[role="option"], [role="menuitem"], li, button, div, span');
        var i;
        for (i = 0; i < nodes.length; i++) {
            if (!visible(nodes[i])) continue;
            if (nodes[i].closest && nodes[i].closest('#sta-retoure-bot')) continue;
            var text = textOf(nodes[i]);
            if (!pattern.test(text) || text.length > 80) continue;
            if (nodes[i].closest && nodes[i].closest('#sta-retoure-bot')) continue;
            return nodes[i];
        }
        return null;
    }

    function reactProps(el) {
        if (!el) return null;
        var target = el.wrappedJSObject || el;
        var keys;
        try { keys = Object.keys(target); } catch (e) { return null; }
        var i;
        for (i = 0; i < keys.length; i++) {
            if (keys[i].indexOf('__reactProps$') === 0) return target[keys[i]];
        }
        return null;
    }

    function openArt(el) {
        var props = reactProps(el);
        if (props && typeof props.onMouseDown === 'function') {
            try {
                props.onMouseDown({ button: 0, currentTarget: el, target: el, preventDefault: function () {}, stopPropagation: function () {} });
                return;
            } catch (e) {}
        }
        fireMouse(el, ['pointerdown', 'mousedown']);
    }

    function openArtByKey(el) {
        try { el.focus(); } catch (e) {}
        pressKey(el, 'ArrowDown');
    }

    function pickOption(el) {
        var row = (el.closest && el.closest('[role="option"], [role="menuitem"], li')) || el;
        var props = reactProps(row);
        if (props && typeof props.onClick === 'function') {
            try {
                var Ctor = pageWindow().MouseEvent || MouseEvent;
                props.onClick({ currentTarget: row, target: row, nativeEvent: new Ctor('click', { bubbles: true }), preventDefault: function () {}, stopPropagation: function () {} });
                return;
            } catch (e) {}
        }
        fireClick(row);
    }

    async function ensureEan() {
        var waited = 0;
        while (waited < 40 && !findArtSelect() && !findArtControl()) {
            await sleep(250);
            waited++;
        }
        var select = findArtSelect();
        if (select && visible(select)) {
            var picked = eanSelectOption(select);
            if (picked && select.value !== picked.value) setNativeValue(select, picked.value);
            await sleep(150);
            return artIsEan();
        }
        if (artIsEan()) return true;
        var control = findArtControl();
        if (!control) return false;
        var attempt;
        for (attempt = 0; attempt < 4; attempt++) {
            var option = findEanOption();
            if (!option) {
                if (attempt % 2 === 0) openArt(artOpener(control));
                else openArtByKey(artOpener(control));
                var tries = 0;
                while (tries < 12 && !option) {
                    await sleep(120);
                    option = findEanOption();
                    tries++;
                }
            }
            if (option) {
                pickOption(option);
                await sleep(250);
                if (artIsEan()) return true;
            }
            control = findArtControl() || control;
        }
        if (artIsEan()) return true;
        var options = document.querySelectorAll('[role="option"], [role="menuitem"], li');
        var list = [];
        var index = -1;
        var current = 0;
        var i;
        for (i = 0; i < options.length; i++) {
            if (!visible(options[i])) continue;
            if (options[i].closest && options[i].closest('#sta-retoure-bot')) continue;
            var word = textOf(options[i]).toLowerCase();
            if (!ART_WORDS[word]) continue;
            if (options[i].querySelector && options[i].querySelector('[role="option"], [role="menuitem"], li')) continue;
            if (word === 'ean') index = list.length;
            if (options[i].getAttribute('aria-selected') === 'true') current = list.length;
            list.push(options[i]);
        }
        if (index < 0) return false;
        var steps = index - current;
        var dir = steps < 0 ? 'ArrowUp' : 'ArrowDown';
        var n = Math.abs(steps);
        var focus = document.activeElement || list[current] || list[0];
        for (i = 0; i < n; i++) {
            pressKey(document.activeElement || focus, dir);
            await sleep(50);
        }
        pressKey(document.activeElement || list[index], 'Enter');
        await sleep(180);
        if (!artIsEan()) fireClick(list[index]);
        await sleep(120);
        return artIsEan();
    }

    function findSearchInput() {
        var inputs = document.querySelectorAll('input');
        var best = null;
        var scoreBest = 0;
        var i;
        for (i = 0; i < inputs.length; i++) {
            var input = inputs[i];
            if (!visible(input)) continue;
            if (input.closest && input.closest('#sta-retoure-bot')) continue;
            var type = (input.getAttribute('type') || 'text').toLowerCase();
            if (type === 'hidden' || type === 'checkbox' || type === 'radio' || type === 'date') continue;
            var blob = ((input.getAttribute('placeholder') || '') + ' ' + (input.getAttribute('aria-label') || '') + ' ' + (input.name || '') + ' ' + (input.id || '')).toLowerCase();
            var parent = input.parentElement;
            var guard = 0;
            while (parent && guard < 3) {
                blob += ' ' + textOf(parent).toLowerCase().slice(0, 40);
                parent = parent.parentElement;
                guard++;
            }
            var score = 0;
            if (/suche/.test(blob)) score += 6;
            if (/ean|search|artikel/.test(blob)) score += 2;
            if (/datum|date|von|bis/.test(blob)) score -= 5;
            if (score > scoreBest) {
                scoreBest = score;
                best = input;
            }
        }
        return scoreBest > 0 ? best : null;
    }

    function buttonHint(button) {
        var hint = (button.getAttribute('aria-label') || '') + ' ' + (button.getAttribute('title') || '') + ' ' + textOf(button);
        var icons = button.querySelectorAll('svg, [data-testid], i');
        var i;
        for (i = 0; i < icons.length; i++) {
            hint += ' ' + (icons[i].getAttribute('data-testid') || '') + ' ' + (icons[i].getAttribute('class') || '');
        }
        return hint.toLowerCase();
    }

    function insideField(button) {
        var node = button.parentElement;
        var guard = 0;
        while (node && guard < 3) {
            if (node.querySelector && node.querySelector('input')) return true;
            node = node.parentElement;
            guard++;
        }
        return false;
    }

    function searchButtonNear(input) {
        if (!input) return null;
        var inputRect = input.getBoundingClientRect();
        var root = input.parentElement;
        var guard = 0;
        while (root && guard < 7) {
            var buttons = Array.prototype.filter.call(root.querySelectorAll('button, [role="button"]'), function (button) {
                if (!visible(button)) return false;
                if (button.closest && button.closest('#sta-retoure-bot')) return false;
                if (insideField(button)) return false;
                var rect = button.getBoundingClientRect();
                if (rect.left < inputRect.right) return false;
                if (Math.abs((rect.top + rect.height / 2) - (inputRect.top + inputRect.height / 2)) > 30) return false;
                return !/refresh|reset|replay|restart|calendar|event|date|kalender|datum/.test(buttonHint(button));
            });
            var named = buttons.filter(function (button) { return /search|such|lupe|magnif/.test(buttonHint(button)); });
            if (named.length) return named[0];
            if (buttons.length) {
                buttons.sort(function (a, b) { return a.getBoundingClientRect().left - b.getBoundingClientRect().left; });
                return buttons[0];
            }
            root = root.parentElement;
            guard++;
        }
        return null;
    }

    function submitSearch(input) {
        input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
        input.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
        var button = searchButtonNear(input);
        if (button) fireClick(button);
    }

    function selectSearch() {
        var input = findSearchInput();
        if (!input) return;
        try {
            input.focus();
            input.select();
            if (typeof input.setSelectionRange === 'function') input.setSelectionRange(0, String(input.value || '').length);
        } catch (e) {}
    }

    function clearSearch() {
        var input = findSearchInput();
        if (!input) return;
        setNativeValue(input, '');
        selectSearch();
    }

    function parseDate(text) {
        var match = String(text || '').match(/(\d{2})\.(\d{2})\.(\d{4})/);
        if (!match) return 0;
        return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1])).getTime();
    }

    function smallestRows(list) {
        return list.filter(function (el) {
            return !list.some(function (other) { return other !== el && el.contains(other); });
        });
    }

    var ROW_SKIP = '#sta-retoure-bot, [role="tooltip"], [role="listbox"], [role="menu"], .MuiPopover-paper, .MuiMenu-paper, .MuiPickersPopper-root, .MuiDateRangeCalendar-root, [class*="Calendar"], [class*="DatePicker"], [class*="DateRange"]';

    function iconControls(row) {
        var icons = Array.prototype.slice.call(row.querySelectorAll('svg, i'));
        var found = [];
        var i;
        for (i = 0; i < icons.length; i++) {
            var node = icons[i];
            var picked = null;
            var guard = 0;
            while (node && node !== row && guard < 5) {
                if (visible(node)) {
                    var rect = node.getBoundingClientRect();
                    if (rect.width >= 18 && rect.width <= 72 && rect.height >= 18 && rect.height <= 56) picked = node;
                    if (node.tagName === 'BUTTON' || node.tagName === 'A' || node.getAttribute('role') === 'button' || node.getAttribute('role') === 'combobox') break;
                }
                node = node.parentElement;
                guard++;
            }
            if (picked && found.indexOf(picked) === -1) found.push(picked);
        }
        return found;
    }

    function rowButtons(row) {
        var raw = Array.prototype.slice.call(row.querySelectorAll('button, a, [role="button"], [role="combobox"], [aria-haspopup]')).concat(iconControls(row));
        var list = raw.filter(function (button, index) {
            if (raw.indexOf(button) !== index) return false;
            if (!visible(button)) return false;
            if (button.closest && button.closest('#sta-retoure-bot')) return false;
            var rect = button.getBoundingClientRect();
            return rect.width <= 90 && rect.height <= 60;
        });
        list = list.filter(function (button) {
            return !list.some(function (other) { return other !== button && button.contains(other); });
        });
        list.sort(function (a, b) { return a.getBoundingClientRect().left - b.getBoundingClientRect().left; });
        return list;
    }

    function isDisabled(button) {
        var node = button;
        var guard = 0;
        while (node && guard < 4) {
            if (node.disabled || node.getAttribute('aria-disabled') === 'true' || node.getAttribute('disabled') !== null) return true;
            var cls = String(node.getAttribute('class') || '');
            if (/(^|\s|-)disabled(\s|$)/i.test(cls)) return true;
            var style = window.getComputedStyle(node);
            if (style.pointerEvents === 'none' || style.cursor === 'not-allowed' || style.cursor === 'default') return true;
            if (Number(style.opacity || '1') < 0.7) return true;
            var bg = (style.backgroundColor || '').match(/(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
            if (bg && Number(bg[1]) > 140 && Number(bg[2]) > 160 && Number(bg[3]) > 210) return true;
            node = node.parentElement;
            guard++;
        }
        return false;
    }

    function returnableLine(row) {
        var button = returnButton(row);
        return !!(button && !isDisabled(button));
    }

    function findMenuItem(pattern) {
        var nodes = document.querySelectorAll('[role="option"], [role="menuitem"], li, button, a, div, span, p');
        var best = null;
        var bestArea = Infinity;
        var i;
        for (i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (!visible(el)) continue;
            if (el.closest && el.closest('#sta-retoure-bot')) continue;
            var text = textOf(el);
            if (text.length > 40 || !pattern.test(text)) continue;
            var rect = el.getBoundingClientRect();
            var area = rect.width * rect.height;
            if (area < bestArea) {
                best = el;
                bestArea = area;
            }
        }
        if (!best) return null;
        return (best.closest && best.closest('[role="option"], [role="menuitem"], li, button, a')) || best;
    }

    var POS_SKIP = ROW_SKIP;

    function numberCells(skip) {
        var nodes = document.querySelectorAll('td, div, span, p, a, b, strong');
        var cells = [];
        var i;
        for (i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            var text = String(el.textContent || '').replace(/\s+/g, ' ').trim();
            if (!/^\d{6,12}$/.test(text)) continue;
            if (!visible(el)) continue;
            if (el.closest && el.closest(skip || ROW_SKIP)) continue;
            if (el.closest('input, button, [role="button"], [role="combobox"]')) continue;
            cells.push(el);
        }
        return smallestRows(cells);
    }

    function rowOfCell(cell) {
        var node = cell.parentElement;
        var guard = 0;
        var picked = null;
        while (node && node !== document.body && guard < 10) {
            var text = textOf(node);
            var dates = (text.match(/\d{2}\.\d{2}\.\d{4}/g) || []).length;
            if (dates > 1 || text.length > 500) break;
            var count = rowButtons(node).length;
            if (dates === 1 && count >= 1) {
                if (!picked || count > rowButtons(picked).length) picked = node;
            }
            node = node.parentElement;
            guard++;
        }
        return picked;
    }

    function tableRows(test, skip) {
        var cells = numberCells(skip);
        var found = [];
        var i;
        for (i = 0; i < cells.length; i++) {
            var row = rowOfCell(cells[i]);
            if (!row || found.indexOf(row) !== -1) continue;
            var text = textOf(row);
            if (/\d{2}\.\d{2}\.\d{4}\s*[-–]\s*\d{2}\.\d{2}\.\d{4}/.test(text)) continue;
            if (!test(text, rowButtons(row))) continue;
            found.push(row);
        }
        return found;
    }

    function voucherRows() {
        return tableRows(function (text, buttons) {
            return buttons.length >= 1 && buttons.length <= 4 && /\d,\d{2}/.test(text) && !/mwst|belegnummer|anzahl der positionen/i.test(text);
        });
    }

    function voucherNumber(row) {
        var match = textOf(row).match(/(?:^|\s)(\d{6,12})(?=\s|$)/);
        return match ? match[1] : '';
    }

    function openButton(row) {
        var buttons = rowButtons(row);
        return buttons.length ? buttons[buttons.length - 1] : null;
    }

    function positionRows() {
        return tableRows(function (text, buttons) {
            return buttons.length >= 3 && /[a-zäöü]{3,}/i.test(text);
        }, POS_SKIP);
    }

    function returnButton(row) {
        var buttons = rowButtons(row);
        if (buttons.length < 3) return null;
        return buttons[buttons.length - 1];
    }

    function debugRows(reason) {
        try {
            var sample = numberCells().slice(0, 12).map(function (cell) {
                var chain = [];
                var node = cell.parentElement;
                var guard = 0;
                while (node && node !== document.body && guard < 8) {
                    chain.push(node.tagName + '.' + String(node.getAttribute('class') || '').split(/\s+/).slice(0, 2).join('.') + ' dates=' + (textOf(node).match(/\d{2}\.\d{2}\.\d{4}/g) || []).length + ' controls=' + rowButtons(node).length + ' len=' + textOf(node).length);
                    node = node.parentElement;
                    guard++;
                }
                var row = rowOfCell(cell);
                return { cell: textOf(cell), row: row ? textOf(row).slice(0, 160) : null, controls: row ? rowButtons(row).map(function (b) { return b.tagName + '[' + (b.getAttribute('role') || '') + '] ' + buttonHint(b).slice(0, 60); }) : [], chain: chain };
            });
            console.log('[STA ' + BOT_VERSION + '] ' + reason + ' ' + JSON.stringify({ vouchers: voucherRows().length, positions: positionRows().length, sample: sample }, null, 1));
        } catch (e) {}
    }

    function samePart(line, target) {
        if (!target) return false;
        var info = parsePosition(textOf(line));
        var norm = function (s) { return String(s || '').replace(/\s+/g, '').toUpperCase(); };
        if (target.dealer && info.dealer && norm(target.dealer) === norm(info.dealer)) return true;
        if (target.article && info.article && norm(target.article) === norm(info.article)) return true;
        return !!(target.name && info.name && info.name.toLowerCase() === target.name.toLowerCase());
    }

    function matchPosition(lines, ean, target) {
        var hits = lines.filter(function (row) { return String(row.textContent || '').indexOf(ean) !== -1; });
        if (hits.length === 1) return hits[0];
        if (target) {
            var byTarget = lines.filter(function (row) { return samePart(row, target); });
            if (byTarget.length) return byTarget.filter(returnableLine)[0] || byTarget[0];
        }
        if (lines.length === 1) return lines[0];
        return null;
    }

    function rowToggle(row) {
        var controls = rowButtons(row);
        if (controls.length < 2) return null;
        return controls[controls.length - 2];
    }

    async function showMatchedPositions(row, number) {
        var toggle = rowToggle(row);
        if (!toggle) return null;
        var before = positionRows();
        pressControl(toggle);
        var item = await waitFor(function () { return findMenuItem(/^positionen?\s+anzeigen$/i); }, 1800);
        if (!item) {
            openArt(toggle);
            item = await waitFor(function () { return findMenuItem(/^positionen?\s+anzeigen$/i); }, 1200);
        }
        if (item) pickOption(item);
        return waitFor(function () {
            var list = positionRows().filter(function (line) { return before.indexOf(line) === -1; });
            if (!list.length) return null;
            var owned = list.filter(function (line) {
                return row.contains(line) || (row.compareDocumentPosition(line) & Node.DOCUMENT_POSITION_FOLLOWING);
            });
            var next = voucherRows().filter(function (other) { return other !== row && (row.compareDocumentPosition(other) & Node.DOCUMENT_POSITION_FOLLOWING); })[0];
            if (next) owned = owned.filter(function (line) { return next.compareDocumentPosition(line) & Node.DOCUMENT_POSITION_PRECEDING; });
            return owned.length ? owned : list;
        }, 4000);
    }

    async function openDelivery(row, number) {
        var opener = openButton(row);
        if (!opener) return false;
        pressControl(opener);
        var opened = await waitFor(function () {
            return (number && location.pathname.indexOf(number) !== -1) || (number && new RegExp('Belegnummer\\s*' + number, 'i').test(textOf(document.body))) ? true : null;
        }, 3000);
        if (opened) return true;
        fireClick(opener);
        opened = await waitFor(function () {
            return (number && location.pathname.indexOf(number) !== -1) || (number && new RegExp('Belegnummer\\s*' + number, 'i').test(textOf(document.body))) ? true : null;
        }, 2500);
        if (opened) return true;
        if (number && location.pathname.indexOf(number) === -1) {
            var base = location.pathname.replace(/\/$/, '').replace(/\/\d{6,12}$/, '');
            try {
                history.pushState(null, '', base + '/' + number);
                window.dispatchEvent(new PopStateEvent('popstate'));
            } catch (e) {}
            opened = await waitFor(function () {
                return new RegExp('Belegnummer\\s*' + number, 'i').test(textOf(document.body)) || positionRows().length ? true : null;
            }, 3500);
        }
        return !!opened;
    }

    function parsePosition(text) {
        var raw = String(text || '').replace(/\s+/g, ' ').trim();
        var cols = raw.match(/(\d{6,12})\s+\d{2}\.\d{2}\.\d{4}\s+(\d{5,12})\s+(.+?)\s+([A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\s,./()-]*?)\s*\|\s*(.+?)(?:\s+MwSt\.?|\s+\d|$)/);
        if (cols) {
            var qtyCol = raw.match(/MwSt\.?\s+[\d.,]+\s+(\d{1,4})\b/i);
            return { article: cols[3].trim() || cols[2], dealer: cols[2], name: cols[4].trim(), brand: cols[5].trim(), qty: qtyCol ? qtyCol[1] : '' };
        }
        var brand = '';
        var nameSide = raw;
        var pipe = raw.indexOf('|');
        if (pipe >= 0) {
            brand = (raw.slice(pipe + 1).trim().split(/\s+/)[0] || '').replace(/[,.].*$/, '');
            nameSide = raw.slice(0, pipe);
        }
        var words = nameSide.split(/\s+/);
        var name = words.filter(function (word) { return /[a-zäöü]/.test(word) && !/^mwst$/i.test(word); }).join(' ');
        var article = '';
        words.forEach(function (word) {
            if (!article && /[A-Za-z]/.test(word) && /\d/.test(word) && word.length >= 4) article = word;
        });
        if (!article) {
            var digits = words.filter(function (word) { return /^\d{5,}$/.test(word); });
            article = digits.length > 1 ? digits[1] : (digits[0] || '');
        }
        var qtyMatch = raw.match(/\b(\d{1,3})\b\s*(?:stk|stück|stueck)?\s*$/i);
        return { article: article, name: name, brand: brand, qty: qtyMatch ? qtyMatch[1] : '' };
    }

    function belegNumber() {
        var body = textOf(document.body);
        var labeled = body.match(/Belegnummer\s*(\d{5,})/i) || body.match(/Lieferschein-Nr\.?\s*(\d{5,})/i);
        if (labeled) return labeled[1];
        var path = location.pathname.match(/\/vouchers\/\d+\/(\d{5,})/i);
        return path ? path[1] : '';
    }

    function angebotNumber(text) {
        var match = String(text || '').match(/Angebot(?:-Nr\.?)?\s*(\d{5,})/i);
        return match ? match[1] : '';
    }

    async function waitFor(fn, timeout) {
        var start = Date.now();
        while (Date.now() - start < timeout) {
            if (abortMission) return null;
            var found = fn();
            if (found) return found;
            await sleep(250);
        }
        return null;
    }

    function sheetParse(raw) {
        var text = String(raw || '').replace(/^\)\]\}'\n?/, '').trim();
        if (!text || text.charAt(0) === '<') return null;
        try { return JSON.parse(text); } catch (e) { return null; }
    }

    function sheetReq(url, hop) {
        return new Promise(function (resolve, reject) {
            GM_xmlhttpRequest({
                method: 'GET',
                url: url,
                anonymous: false,
                timeout: 30000,
                onload: function (res) {
                    var data = sheetParse(res && res.responseText);
                    if (data && data.ok === true) resolve(data);
                    else if (data && data.ok === false) reject(new Error(data.error || 'Sheet hat abgelehnt'));
                    else reject(new Error('Sheet Antwort ' + (res && res.status)));
                },
                onerror: function () { reject(new Error('Sheet nicht erreichbar')); },
                ontimeout: function () { reject(new Error('Sheet Timeout')); }
            });
        });
    }

    function postSheet(row) {
        var url = sheetUrl();
        if (!url) return Promise.resolve({ ok: false, error: 'Sheet-Link fehlt. Oben auf das Zahnrad.' });
        if (row.token && sheetBusy === row.token) return Promise.resolve({ ok: false, error: 'läuft' });
        sheetBusy = row.token || 'x';
        var payload = JSON.stringify({ action: 'staAppend', entries: [row] });
        var join = url.indexOf('?') === -1 ? '?' : '&';
        return sheetReq(url + join + 'payload=' + encodeURIComponent(payload), 0).then(function (data) {
            if (sheetBusy === row.token) sheetBusy = '';
            return data;
        }).catch(function (err) {
            if (sheetBusy === row.token) sheetBusy = '';
            return { ok: false, error: err && err.message ? err.message : 'Sheet fehlgeschlagen' };
        });
    }

    function askChoice(lines) {
        setSlot('');
        paintHud('Lieferung hat ' + lines.length + ' Positionen. Welches Teil geht zurück?', '');
        var slot = setSlot('<div class="picks" id="sta-picks"></div>');
        var box = slot.querySelector('#sta-picks');
        return new Promise(function (resolve) {
            waitPick = resolve;
            lines.forEach(function (line, index) {
                var info = parsePosition(textOf(line));
                var grey = !returnableLine(line);
                var button = document.createElement('button');
                button.type = 'button';
                button.className = grey ? 'pick grey' : 'pick';
                var title = document.createElement('b');
                title.textContent = (info.name || 'Position ' + (index + 1)) + (info.brand ? ' | ' + info.brand : '');
                var sub = document.createElement('span');
                sub.textContent = [info.article, info.dealer && info.dealer !== info.article ? 'Händler ' + info.dealer : '', info.qty ? info.qty + ' Stk' : '', grey ? 'hier ausgegraut → nächstes Datum' : ''].filter(Boolean).join(' · ');
                button.appendChild(title);
                button.appendChild(sub);
                button.addEventListener('click', function () { resolve(lines[index]); });
                box.appendChild(button);
            });
            var cancel = document.createElement('button');
            cancel.type = 'button';
            cancel.className = 'ghost';
            cancel.textContent = 'Abbrechen';
            cancel.addEventListener('click', function () { resolve(null); });
            box.appendChild(cancel);
        }).then(function (picked) {
            waitPick = null;
            setSlot('');
            return picked;
        });
    }

    function askStock(info) {
        setSlot('');
        paintHud('Gefunden in Lieferung ' + (info.beleg || '?') + '. Stock-ID ist optional.', '');
        var slot = setSlot('<div class="last"><em>TEIL</em><div>' + hudEscape((info.name || 'Teil') + (info.brand ? ' | ' + info.brand : '')) + '</div>'
            + '<div>' + hudEscape([info.article, info.qty ? info.qty + ' Stk' : ''].filter(Boolean).join(' · ')) + '</div></div>'
            + '<label>Stock-ID optional</label><input id="sta-stock" maxlength="16" placeholder="z.B. MA45327" autocomplete="off">'
            + '<div class="row"><button type="button" class="ghost" id="sta-skip">Ohne Stock-ID</button><button type="button" class="act" id="sta-add">Hinzufügen</button></div>');
        var input = slot.querySelector('#sta-stock');
        return new Promise(function (resolve) {
            function done(stock) { resolve(stock); }
            waitPick = function () { done(null); };
            slot.querySelector('#sta-add').addEventListener('click', function () { done(input.value.trim()); });
            slot.querySelector('#sta-skip').addEventListener('click', function () { done(''); });
            input.addEventListener('keydown', function (event) {
                if (event.key === 'Enter') done(input.value.trim());
            });
            setTimeout(function () { try { input.focus(); } catch (e) {} }, 30);
        }).then(function (stock) {
            waitPick = null;
            setSlot('');
            return stock;
        });
    }

    function cleanStock(value) {
        var text = String(value || '').toUpperCase().replace(/\s+/g, '');
        if (!text) return '';
        var match = text.match(/[A-Z]{1,3}\d{4,8}/);
        return match ? match[0] : null;
    }

    async function addEan(raw) {
        var ean = String(raw || '').replace(/\D/g, '');
        if (ean.length < MIN_SCAN || ean.length > MAX_SCAN) return;
        if (busy) return;
        busy = true;
        abortMission = false;
        paintHud('Suche ' + ean + ' …', '');
        try {
            await ensureEan();
            var input = findSearchInput();
            if (!input) {
                paintHud('Das Suche-Feld auf Lieferungen wurde nicht gefunden.', 'bad');
                return;
            }
            var before = voucherRows().map(textOf).join('|');
            setNativeValue(input, ean);
            await sleep(80);
            submitSearch(input);
            var started = Date.now();
            var rows = await waitFor(function () {
                var list = voucherRows();
                if (!list.length) return null;
                if (list.map(textOf).join('|') === before && Date.now() - started < 2500) return null;
                return list;
            }, 9000);
            if (!rows) {
                debugRows('keine Lieferung');
                paintHud('Keine Lieferung zu ' + ean + ' gefunden.', 'bad');
                selectSearch();
                return;
            }
            await sleep(300);
            rows = voucherRows().length ? voucherRows() : rows;
            rows.sort(function (a, b) { return parseDate(textOf(a)) - parseDate(textOf(b)); });
            var numbers = rows.map(voucherNumber).filter(Boolean);
            var dates = rows.map(function (r) { var m = textOf(r).match(/\d{2}\.\d{2}\.\d{4}/); return m ? m[0] : ''; });
            var targetPart = null;
            var line = null;
            var number = '';
            var hadGrey = false;
            var idx;
            for (idx = 0; idx < numbers.length; idx++) {
                if (abortMission) return;
                number = numbers[idx];
                var when = dates[idx];
                var rowEl = voucherRows().filter(function (r) { return voucherNumber(r) === number; })[0] || rows[idx];
                paintHud('Prüfe Lieferung ' + (idx + 1) + '/' + numbers.length + ': ' + number + ' vom ' + when + ' …', '');
                var beforeSigns = positionRows().map(textOf).join('|');
                var lines = await showMatchedPositions(rowEl, number);
                if (!lines) {
                    await openDelivery(rowEl, number);
                    await sleep(400);
                    var waitStart = Date.now();
                    lines = await waitFor(function () {
                        var list = positionRows();
                        if (!list.length) return null;
                        if (idx > 0 && list.map(textOf).join('|') === beforeSigns && Date.now() - waitStart < 2500) return null;
                        return list;
                    }, 6000);
                }
                if (!lines || !lines.length) continue;
                var candidate = lines.length === 1 ? lines[0] : matchPosition(lines, ean, targetPart);
                if (!candidate) {
                    if (targetPart) continue;
                    candidate = await askChoice(lines);
                    if (!candidate) {
                        paintHud('Abgebrochen. Nächste EAN scannen.', '');
                        clearSearch();
                        return;
                    }
                }
                var candInfo = parsePosition(textOf(candidate));
                if (!targetPart && (candInfo.article || candInfo.dealer || candInfo.name)) targetPart = candInfo;
                if (!returnableLine(candidate)) {
                    hadGrey = true;
                    paintHud('In ' + number + ' (' + when + ') ausgegraut. Prüfe nächstes Datum …', '');
                    await sleep(250);
                    continue;
                }
                line = candidate;
                break;
            }
            if (!line) {
                debugRows(hadGrey ? 'alle ausgegraut' : 'keine Position');
                paintHud(hadGrey
                    ? 'In allen ' + numbers.length + ' Lieferungen ist dieses Teil ausgegraut (nicht mehr rückgabefähig).'
                    : 'In keiner Lieferung eine passende Position gefunden.', 'bad');
                clearSearch();
                return;
            }
            var lineText = textOf(line);
            var parsed = parsePosition(lineText);
            var beleg = number || belegNumber();
            try {
                line.classList.add('sta-picked-row');
                line.scrollIntoView({ block: 'center' });
            } catch (e) {}
            var stockRaw = await askStock({ name: parsed.name, article: parsed.article, brand: parsed.brand, qty: parsed.qty, beleg: beleg });
            if (stockRaw == null) return;
            var stock = cleanStock(stockRaw);
            if (stock == null) {
                paintHud('Stock-ID passt nicht. Leer lassen oder z.B. MA45327.', 'bad');
                return;
            }
            if (!document.contains(line)) {
                line = positionRows().filter(function (row) { return textOf(row) === lineText; })[0] || line;
            }
            try { line.classList.remove('sta-picked-row'); } catch (e2) {}
            var button = returnButton(line);
            if (!button) {
                debugRows('kein Rückgabe-Icon');
                paintHud('Rückgabe-Icon an der Position nicht gefunden.', 'bad');
                selectSearch();
                return;
            }
            if (isDisabled(button)) {
                paintHud((parsed.name || 'Teil') + ' liegt schon in der Warenrückgabe oder ist nicht rückgabefähig.', 'bad');
                clearSearch();
                return;
            }
            pressControl(button);
            await sleep(400);
            var row = {
                orderId: beleg,
                stockId: stock,
                article: parsed.article || ean,
                name: parsed.name,
                supplier: 'STAKIS',
                brand: parsed.brand,
                who: workerName(),
                ean: ean,
                qty: parsed.qty,
                token: [ean, beleg, parsed.article || ean].join('|')
            };
            var saved = await postSheet(row);
            var note = saved && saved.ok ? ((saved.added || 0) > 0 ? 'Im Blatt STA' : 'Steht schon im Blatt STA') : ('Nicht im Blatt: ' + ((saved && saved.error) || 'Fehler'));
            saveLast({
                name: parsed.name,
                article: parsed.article || ean,
                beleg: beleg,
                stockId: stock,
                sheet: note
            });
            setBasketCount(basketCount() + 1);
            clearSearch();
            paintHud('Drin. Suche ist leer für das nächste Teil. ' + note, saved && saved.ok ? 'ok' : 'bad');
        } catch (err) {
            paintHud(err && err.message ? err.message : String(err), 'bad');
            selectSearch();
        } finally {
            busy = false;
        }
    }

    function bindSearch(input) {
        if (!input || input === boundInput) return;
        boundInput = input;
        input.addEventListener('input', function () {
            if (busy || !isLieferungen()) return;
            var digits = String(input.value || '').replace(/\D/g, '');
            if (scanTimer) clearTimeout(scanTimer);
            if (digits.length < MIN_SCAN || digits.length > MAX_SCAN) return;
            scanTimer = setTimeout(function () { addEan(digits); }, SCAN_IDLE_MS);
        });
        input.addEventListener('keydown', function (event) {
            if (event.key !== 'Enter' || !isLieferungen() || busy) return;
            var digits = String(input.value || '').replace(/\D/g, '');
            if (digits.length < MIN_SCAN) return;
            event.preventDefault();
            if (scanTimer) clearTimeout(scanTimer);
            addEan(digits);
        });
    }

    async function applyReasons() {
        var changed = 0;
        var selects = document.querySelectorAll('select');
        var i;
        for (i = 0; i < selects.length; i++) {
            var opt = Array.prototype.filter.call(selects[i].options, function (item) {
                return /falscher artikel/i.test(textOf(item));
            })[0];
            if (!opt) continue;
            if (selects[i].value !== opt.value) {
                setNativeValue(selects[i], opt.value);
                changed++;
            }
        }
        var buttons = document.querySelectorAll('button, [role="button"], [role="combobox"]');
        for (i = 0; i < buttons.length; i++) {
            if (!visible(buttons[i])) continue;
            if (buttons[i].closest && buttons[i].closest('#sta-retoure-bot')) continue;
            var label = textOf(buttons[i]);
            if (!/rückgabegrund|rueckgabegrund/i.test(label)) continue;
            if (/falscher artikel/i.test(label)) continue;
            fireClick(buttons[i]);
            await sleep(220);
            var item = findMenu(/falscher artikel/i);
            if (item) {
                fireClick(item);
                changed++;
                await sleep(180);
            }
        }
        return changed;
    }

    async function runReasons() {
        if (busy) return;
        busy = true;
        paintHud('Setze alle Gründe auf Falscher Artikel …', '');
        try {
            var changed = await applyReasons();
            paintHud(changed ? (changed + ' Gründe auf Falscher Artikel.') : 'Keine offenen Rückgabegründe gefunden.', changed ? 'ok' : 'bad');
        } finally {
            busy = false;
        }
    }

    function printButton() {
        var nodes = document.querySelectorAll('button, [role="button"], a');
        var send = null;
        var i;
        for (i = 0; i < nodes.length; i++) {
            if (/WARENRÜCKGABE BEAUFTRAGEN/i.test(textOf(nodes[i]))) {
                send = nodes[i];
                break;
            }
        }
        if (!send) return null;
        var parent = send.parentElement;
        var previous = send.previousElementSibling;
        while (previous) {
            if (previous.matches && previous.matches('button, [role="button"], a')) return previous;
            var inner = previous.querySelector && previous.querySelector('button, [role="button"], a');
            if (inner) return inner;
            previous = previous.previousElementSibling;
        }
        if (!parent) return null;
        var buttons = parent.querySelectorAll('button, [role="button"]');
        var prev = null;
        for (i = 0; i < buttons.length; i++) {
            if (buttons[i] === send || buttons[i].contains(send)) return prev;
            prev = buttons[i];
        }
        return null;
    }

    function finishCycle(message) {
        setBasketCount(0);
        saveLast(null);
        goLieferungen();
        setTimeout(function () {
            clearSearch();
            paintHud(message || 'Fertig. Nächste EAN scannen.', 'ok');
        }, 600);
    }

    async function runPrint() {
        if (busy) return;
        busy = true;
        paintHud('Gründe setzen, dann drucken und absenden …', '');
        try {
            await applyReasons();
            var button = printButton();
            if (!button) {
                paintHud('Das Druck-Icon neben Warenrückgabe beauftragen fehlt.', 'bad');
                return;
            }
            fireClick(button);
            paintHud('Druck ist angestoßen. Danach geht es zurück zur Suche.', 'ok');
            setTimeout(function () { finishCycle('Nächster Vorgang. EAN scannen.'); }, 1600);
        } finally {
            busy = false;
        }
    }

    function pageSection() {
        if (isReturns()) return 'returns';
        if (isLieferungen()) return 'vouchers';
        return '';
    }

    function boot() {
        if (!isBotPage()) {
            var old = hudEl();
            if (old) old.remove();
            return;
        }
        if (busy) return;
        if (isReturns()) {
            setSlot('');
            paintHud('Warenrückgabe. Gründe setzen, dann Drucken und absenden.', '');
            return;
        }
        if (artIsEan()) {
            showEanState(true, true);
            var readyInput = findSearchInput();
            if (readyInput) {
                bindSearch(readyInput);
                if (!hudFocused()) selectSearch();
            }
            return;
        }
        paintHud('Stelle Art auf EAN …', '');
        runEanGuard(true);
    }

    var eanRunning = false;
    var eanShown = null;
    var eanLastTry = 0;

    function hudFocused() {
        var panel = hudEl();
        return !!(panel && document.activeElement && panel.contains(document.activeElement));
    }

    function showEanState(ok, force) {
        if (!isLieferungen() || busy) return;
        if (!force && eanShown === ok) return;
        if (!force && hudFocused()) return;
        eanShown = ok;
        paintHud(ok
            ? 'EAN scannen. Art steht auf EAN. Der älteste Beleg wird genommen.'
            : 'Art konnte nicht auf EAN gestellt werden. Versuche es weiter …', ok ? '' : 'bad');
    }

    function runEanGuard(force) {
        if (eanRunning || busy || !isLieferungen()) return;
        eanRunning = true;
        eanLastTry = Date.now();
        ensureEan().then(function (ok) {
            eanRunning = false;
            showEanState(ok, force);
            var input = findSearchInput();
            if (input) {
                bindSearch(input);
                if (ok && force && !hudFocused()) {
                    try { input.focus(); } catch (e) {}
                }
            }
        }, function () {
            eanRunning = false;
        });
    }

    setInterval(function () {
        if (!isLieferungen()) return;
        var input = findSearchInput();
        if (input) bindSearch(input);
        if (busy || eanRunning) return;
        var ok = artIsEan();
        if (ok) {
            showEanState(true, false);
            return;
        }
        if (Date.now() - eanLastTry > 2500) runEanGuard(false);
    }, 1000);

    document.addEventListener('keydown', function (event) {
        if (event.key !== 'Escape') return;
        abortMission = true;
        busy = false;
        if (scanTimer) clearTimeout(scanTimer);
        if (waitPick) waitPick(null);
        setSlot('');
        paintHud('Gestoppt. ESC nochmal schließt nichts weiter. Nächste EAN scannen.', 'bad');
        selectSearch();
    });

    var lastSection = pageSection();
    setInterval(function () {
        var section = pageSection();
        if (section !== lastSection) {
            lastSection = section;
            boot();
            return;
        }
        if (!hudEl() && isBotPage()) boot();
    }, 700);
    boot();
})();
