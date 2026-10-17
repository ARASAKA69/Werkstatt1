// ==UserScript==
// @name         Carol Tagesliste Herausgegeben Bridge
// @namespace    http://tampermonkey.net/
// @version      1.1
// @description  Carol Herausgegeben Automatisierung // Tagesliste
// @author       ARASAKA
// @match        *://carol.autohero.com/*
// @match        *://docs.google.com/spreadsheets/*
// @match        *://script.google.com/*
// @match        *://*.googleusercontent.com/*
// @updateURL    https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/DailyPlanning/carol-tagesliste-bridge.user.js
// @downloadURL  https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/DailyPlanning/carol-tagesliste-bridge.user.js
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addValueChangeListener
// ==/UserScript==

(function () {
    'use strict';

    const KEY_QUEUE = 'dpl_tl_carol_queue_v4';
    const KEY_STATE = 'dpl_tl_carol_state_v4';
    const KEY_EVENTS = 'dpl_tl_carol_events_v4';
    const KEY_ABORT = 'dpl_tl_carol_abort_v4';
    const KEY_HEARTBEAT = 'dpl_tl_carol_hb_v4';

    try {
        if (typeof GM_setValue === 'function') {
            GM_setValue('dpl_tl_carol_queue_v1', '[]');
            GM_setValue('dpl_tl_carol_queue_v2', '[]');
            GM_setValue('dpl_tl_carol_queue_v3', '[]');
        }
    } catch (e0) {}

    const isCarolHost = location.hostname.includes('carol.autohero.com');
    const isWmsSingleAuto = isCarolHost && location.href.includes('arasaka_auto=1');

    function readJson(key, fallback) {
        try {
            if (typeof GM_getValue === 'function') {
                const rawGm = GM_getValue(key, '');
                if (rawGm) return JSON.parse(rawGm);
            }
        } catch (e1) {}
        return fallback;
    }

    function writeJson(key, val) {
        try {
            if (typeof GM_setValue === 'function') {
                GM_setValue(key, JSON.stringify(val));
            }
        } catch (e1) {}
    }

    if (!isCarolHost) {
        let lastEventCursor = 0;
        let lastStateTs = 0;

        function postToFrameTree(win, payload, depth) {
            if (!win || depth > 4) return;
            try { win.postMessage(payload, '*'); } catch (e0) {}
            try {
                for (let i = 0; i < win.frames.length; i++) {
                    postToFrameTree(win.frames[i], payload, depth + 1);
                }
            } catch (e1) {}
        }

        function forwardToHud() {
            const hb = readJson(KEY_HEARTBEAT, null);
            if (hb && (Date.now() - Number(hb.ts || 0)) < 25000) {
                postToFrameTree(window, {
                    source: 'dpl-tagesliste-bridge',
                    type: 'heartbeat',
                    idle: !!hb.idle,
                    ts: hb.ts
                }, 0);
            }

            const st = readJson(KEY_STATE, null);
            if (st && Number(st.ts || 0) > lastStateTs && (Date.now() - Number(st.ts || 0)) < 25000) {
                lastStateTs = Number(st.ts || 0);
                postToFrameTree(window, {
                    source: 'dpl-tagesliste-bridge',
                    type: 'state',
                    busy: !!st.busy,
                    currentSid: st.currentSid || '',
                    stepText: st.stepText || '',
                    ts: st.ts
                }, 0);
            }

            const evs = readJson(KEY_EVENTS, []);
            if (Array.isArray(evs) && evs.length) {
                for (let i = 0; i < evs.length; i++) {
                    const ev = evs[i];
                    if (!ev || !ev.stockId) continue;
                    if (Number(ev.ts || 0) <= lastEventCursor) continue;
                    postToFrameTree(window, {
                        source: 'dpl-tagesliste-bridge',
                        type: 'item_result',
                        eventId: ev.eventId,
                        stockId: ev.stockId,
                        dateStr: ev.dateStr || '',
                        rows: ev.rows || [],
                        status: ev.status || 'ok',
                        message: ev.message || '',
                        ts: ev.ts
                    }, 0);
                }
                const maxTs = evs.reduce((m, x) => Math.max(m, Number((x && x.ts) || 0)), lastEventCursor);
                lastEventCursor = maxTs;
            }
        }

        window.addEventListener('message', (ev) => {
            const d = ev && ev.data;
            if (!d || d.source !== 'dpl-tagesliste-hud') return;

            if (d.cmd === 'ping') {
                forwardToHud();
                return;
            }

            if (d.cmd === 'stop_all') {
                writeJson(KEY_ABORT, { ts: Date.now() });
                writeJson(KEY_QUEUE, []);
                writeJson(KEY_STATE, { busy: false, currentSid: '', stepText: '', ts: Date.now() });
                return;
            }

            if (d.cmd === 'enqueue_jobs') {
                const incoming = (d.payload && d.payload.jobs) || [];
                if (!incoming.length) return;
                const bySid = {};
                const freshQueue = [];
                for (let j = 0; j < incoming.length; j++) {
                    const inc = incoming[j];
                    if (!inc || !inc.stockId) continue;
                    const sid = String(inc.stockId).trim().toUpperCase();
                    if (!bySid[sid]) {
                        const job = {
                            stockId: sid,
                            dateStr: inc.dateStr || '',
                            carolUrl: inc.carolUrl || ('https://carol.autohero.com/en-GB/refurbishment?rsv=' + encodeURIComponent(sid)),
                            model: inc.model || '',
                            rows: inc.rows || [],
                            row: inc.row || 0,
                            status: 'pending',
                            phase: 'init',
                            addedAt: Date.now()
                        };
                        bySid[sid] = job;
                        freshQueue.push(job);
                    }
                }
                writeJson(KEY_ABORT, { ts: 0 });
                writeJson(KEY_QUEUE, freshQueue);
            }
        });

        if (typeof GM_addValueChangeListener === 'function') {
            GM_addValueChangeListener(KEY_STATE, () => forwardToHud());
            GM_addValueChangeListener(KEY_EVENTS, () => forwardToHud());
            GM_addValueChangeListener(KEY_HEARTBEAT, () => forwardToHud());
        }
        setInterval(forwardToHud, 350);
        return;
    }

    if (isWmsSingleAuto) return;

    if (location.href.includes('tl_standby=1') || location.href.includes('tl_job=1')) {
        try { sessionStorage.setItem('dpl_tl_worker_tab', '1'); } catch (e) {}
    }

    let workerBusy = false;
    let isNavigating = false;
    let abortFlag = false;
    let currentSid = '';
    let currentStepText = '';
    let hudEl = null;

    const initialQ = readJson(KEY_QUEUE, []);
    if (Array.isArray(initialQ) && initialQ.length > 0 && initialQ[0] && initialQ[0].stockId) {
        currentSid = initialQ[0].stockId;
        currentStepText = 'Lade Auftrag in Carol…';
        try { sessionStorage.setItem('dpl_tl_worker_tab', '1'); } catch (e) {}
    }

    function isWorkerTab() {
        try {
            if (sessionStorage.getItem('dpl_tl_worker_tab') === '1') return true;
        } catch (e) {}
        const q = readJson(KEY_QUEUE, []);
        return Array.isArray(q) && q.length > 0;
    }

    function postToOpenerTree(win, payload, depth) {
        if (!win || depth > 4) return;
        try { win.postMessage(payload, '*'); } catch (e0) {}
        try {
            for (let i = 0; i < win.frames.length; i++) {
                postToOpenerTree(win.frames[i], payload, depth + 1);
            }
        } catch (e1) {}
    }

    function sendToOpenerAllFrames(payload) {
        try {
            if (window.opener && !window.opener.closed) {
                let root = window.opener;
                try { if (root.top) root = root.top; } catch (e) {}
                postToOpenerTree(root, payload, 0);
            }
        } catch (e) {}
    }

    function sendHeartbeat() {
        const ts = Date.now();
        const q = readJson(KEY_QUEUE, []);
        const isIdle = !workerBusy && !isNavigating && (!Array.isArray(q) || q.length === 0);
        writeJson(KEY_HEARTBEAT, { ts: ts, idle: isIdle });
        sendToOpenerAllFrames({
            source: 'dpl-tagesliste-bridge',
            type: 'heartbeat',
            idle: isIdle,
            ts: ts
        });
    }

    function publishActiveState(busy, sid, stepText) {
        const payload = {
            busy: !!busy,
            currentSid: sid || '',
            stepText: stepText || '',
            ts: Date.now()
        };
        writeJson(KEY_STATE, payload);
        writeJson(KEY_HEARTBEAT, { ts: payload.ts, idle: !payload.busy });
        sendToOpenerAllFrames({
            source: 'dpl-tagesliste-bridge',
            type: 'state',
            busy: payload.busy,
            currentSid: payload.currentSid,
            stepText: payload.stepText,
            ts: payload.ts
        });
    }

    function pushResultEvent(stockId, status, message, jobMeta) {
        const ev = {
            eventId: stockId + '_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
            stockId: stockId,
            dateStr: (jobMeta && jobMeta.dateStr) || '',
            rows: (jobMeta && jobMeta.rows) || [],
            status: status,
            message: message || '',
            ts: Date.now()
        };
        const list = readJson(KEY_EVENTS, []);
        const trimmed = Array.isArray(list) ? list.slice(-40) : [];
        trimmed.push(ev);
        writeJson(KEY_EVENTS, trimmed);

        const msgPayload = {
            source: 'dpl-tagesliste-bridge',
            type: 'item_result',
            eventId: ev.eventId,
            stockId: ev.stockId,
            dateStr: ev.dateStr,
            rows: ev.rows,
            status: ev.status,
            message: ev.message,
            ts: ev.ts
        };
        sendToOpenerAllFrames(msgPayload);
    }

    function ensureHud() {
        if (!document.body || !isWorkerTab()) return;
        if (!document.getElementById('dpl-tl-bridge-style')) {
            const st = document.createElement('style');
            st.id = 'dpl-tl-bridge-style';
            st.textContent = `
                #dpl-tl-bridge-hud {
                    position: fixed;
                    bottom: 20px;
                    left: 20px;
                    width: 350px;
                    background: rgba(8, 12, 18, 0.96);
                    border: 1.5px solid #00ffcc;
                    border-radius: 10px;
                    color: #e6edf3;
                    font-family: "Segoe UI", Consolas, monospace;
                    font-size: 13px;
                    z-index: 999998;
                    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.55), 0 0 18px rgba(0, 255, 204, 0.2);
                    overflow: hidden;
                    user-select: none;
                }
                .dpl-tl-head {
                    display: flex;
                    align-items: center;
                    justify-content: space-between;
                    padding: 9px 12px;
                    background: rgba(0, 255, 204, 0.1);
                    border-bottom: 1px solid rgba(0, 255, 204, 0.25);
                    font-weight: 800;
                    font-size: 12px;
                    letter-spacing: 0.8px;
                    color: #00ffcc;
                }
                .dpl-tl-body {
                    padding: 10px 12px 6px;
                }
                .dpl-tl-sid {
                    font-size: 15px;
                    font-weight: 800;
                    color: #ffffff;
                    margin-bottom: 4px;
                }
                .dpl-tl-step {
                    font-size: 12.5px;
                    color: #8bcef7;
                    white-space: nowrap;
                    overflow: hidden;
                    text-overflow: ellipsis;
                }
                .dpl-tl-foot {
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    padding: 4px 12px 8px;
                    font-size: 10.5px;
                    color: #6e7681;
                    font-weight: 600;
                }
                .dpl-tl-stop {
                    background: rgba(248, 81, 73, 0.18);
                    border: 1px solid #f85149;
                    color: #ff7b72;
                    border-radius: 5px;
                    padding: 2px 8px;
                    font-size: 11px;
                    font-weight: 700;
                    cursor: pointer;
                }
            `;
            document.head.appendChild(st);
        }
        if (!hudEl || !document.body.contains(hudEl)) {
            hudEl = document.createElement('div');
            hudEl.id = 'dpl-tl-bridge-hud';
            hudEl.innerHTML = `
                <div class="dpl-tl-head">
                    <span>TAGESLISTE // CAROL BRIDGE</span>
                    <button type="button" class="dpl-tl-stop" id="dpl-tl-stop-btn">STOP</button>
                </div>
                <div class="dpl-tl-body">
                    <div class="dpl-tl-sid" id="dpl-tl-sid-txt">${currentSid || 'BEREIT'}</div>
                    <div class="dpl-tl-step" id="dpl-tl-step-txt">${currentStepText || 'Bereit für Stock-IDs vom HUD'}</div>
                </div>
                <div class="dpl-tl-foot">
                    <span>Carol Bridge v1.1 · [ESC] Stop</span>
                    <span>By Arasaka</span>
                </div>
            `;
            document.body.appendChild(hudEl);
            const stopBtn = document.getElementById('dpl-tl-stop-btn');
            if (stopBtn) {
                stopBtn.addEventListener('click', () => {
                    abortFlag = true;
                    workerBusy = false;
                    isNavigating = false;
                    writeJson(KEY_QUEUE, []);
                    setHudStatus('GESTOPPT', 'Warteschlange geleert', '#f85149', false);
                });
            }
        }
    }

    function setHudStatus(sidLabel, stepText, borderColor, isBusyState = true) {
        currentStepText = stepText || '';
        ensureHud();
        if (hudEl) {
            if (borderColor) hudEl.style.borderColor = borderColor;
            const sidEl = document.getElementById('dpl-tl-sid-txt');
            const stEl = document.getElementById('dpl-tl-step-txt');
            if (sidEl) sidEl.textContent = sidLabel || 'BEREIT';
            if (stEl) stEl.textContent = stepText || '';
        }
        if (isBusyState && currentSid) {
            publishActiveState(true, currentSid, stepText);
        } else if (!isBusyState) {
            publishActiveState(false, '', '');
        }
    }

    function playTone(ok) {
        try {
            const ctx = new (window.AudioContext || window.webkitAudioContext)();
            const osc = ctx.createOscillator();
            const gain = ctx.createGain();
            osc.connect(gain);
            gain.connect(ctx.destination);
            osc.type = ok ? 'sine' : 'sawtooth';
            osc.frequency.setValueAtTime(ok ? 880 : 180, ctx.currentTime);
            if (ok) osc.frequency.exponentialRampToValueAtTime(1560, ctx.currentTime + 0.12);
            gain.gain.setValueAtTime(0.18, ctx.currentTime);
            gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.35);
            osc.start();
            osc.stop(ctx.currentTime + 0.35);
        } catch (e) {}
    }

    async function sleep(ms) {
        let elapsed = 0;
        const step = Math.min(50, ms);
        while (elapsed < ms) {
            if (abortFlag) throw new Error('Abort');
            await new Promise(r => setTimeout(r, step));
            elapsed += step;
        }
    }

    function getDeepestElementsByText(text) {
        const results = [];
        for (const el of document.querySelectorAll('*')) {
            if (['SCRIPT', 'STYLE', 'HTML', 'HEAD', 'BODY'].includes(el.tagName)) continue;
            if (el.textContent?.includes(text)) {
                const childHasText = Array.from(el.children).some(c => c.textContent?.includes(text));
                if (!childHasText) results.push(el);
            }
        }
        return results;
    }

    async function waitForElement(text, timeout = 8000, pickLast = false) {
        const start = Date.now();
        while ((Date.now() - start) < timeout) {
            if (abortFlag) throw new Error('Abort');
            const els = getDeepestElementsByText(text);
            if (els.length > 0) return pickLast ? els[els.length - 1] : els[0];
            await sleep(80);
        }
        return null;
    }

    function forceClick(el) {
        if (!el) return;
        try {
            const orig = { outline: el.style.outline, shadow: el.style.boxShadow, trans: el.style.transition };
            el.style.transition = 'all 0.1s';
            el.style.outline = '2px solid #ff0000';
            el.style.boxShadow = '0 0 15px #ff0000';
            setTimeout(() => {
                if (el) { el.style.outline = orig.outline; el.style.boxShadow = orig.shadow; el.style.transition = orig.trans; }
            }, 300);
        } catch (e) {}
        el.click();
        if (el.parentElement) setTimeout(() => el.parentElement.click(), 30);
    }

    function getSparePartSelects() {
        const out = [];
        for (const sel of document.querySelectorAll('select')) {
            const hasOpt = Array.from(sel.options).some(o => o.text.includes('Handed out'));
            if (hasOpt) out.push(sel);
        }
        return out;
    }

    function checkAllHandedOut() {
        const selects = getSparePartSelects();
        if (!selects.length) return false;
        for (const sel of selects) {
            if (!sel.options[sel.selectedIndex]?.text.includes('Handed out')) return false;
        }
        return true;
    }

    function forceAllHandedOut() {
        for (const sel of getSparePartSelects()) {
            const opt = Array.from(sel.options).find(o => o.text.includes('Handed out'));
            if (opt && !sel.options[sel.selectedIndex]?.text.includes('Handed out')) {
                sel.value = opt.value;
                sel.dispatchEvent(new Event('change', { bubbles: true }));
            }
        }
    }

    function findExactCloseButton() {
        const candidates = [];
        for (const el of document.querySelectorAll('button, [role="button"], a')) {
            const txt = (el.textContent || '').trim();
            if (txt === 'Close') candidates.push(el);
        }
        return candidates.length ? candidates[candidates.length - 1] : null;
    }

    function findExactConfirmButton() {
        const candidates = [];
        for (const el of document.querySelectorAll('button, [role="button"]')) {
            const txt = (el.textContent || '').trim();
            if (txt === 'Confirm' || txt === 'Bestätigen') candidates.push(el);
        }
        return candidates.length ? candidates[candidates.length - 1] : null;
    }

    async function waitAndClickConfirmButton(timeout = 6000) {
        const start = Date.now();
        while ((Date.now() - start) < timeout) {
            if (abortFlag) throw new Error('Abort');
            const btn = findExactConfirmButton();
            if (btn) {
                forceClick(btn);
                await sleep(300);
                const stillThere = findExactConfirmButton();
                if (stillThere) {
                    forceClick(stillThere);
                }
                return true;
            }
            await sleep(80);
        }
        return false;
    }

    async function clickSavedDialogCloseFast(timeout = 2200) {
        const start = Date.now();
        while ((Date.now() - start) < timeout) {
            if (abortFlag) throw new Error('Abort');
            const btn = findExactCloseButton();
            if (btn) {
                try {
                    btn.style.outline = '2px solid #00ff00';
                    btn.style.boxShadow = '0 0 15px #00ff00';
                } catch (e) {}
                btn.click();
                return true;
            }
            await sleep(80);
        }
        return false;
    }

    function isRefurbDetailUrl(url) {
        return /\/refurbishment\/[0-9a-f-]{8,}/i.test(String(url || ''));
    }

    function extractRefurbUuid(url) {
        const m = String(url || '').match(/\/refurbishment\/([0-9a-f-]{8,})/i);
        return m ? m[1].toLowerCase() : '';
    }

    function isAlreadyOnPostAuthorizationPage() {
        const txt = String((document.body && document.body.innerText) || '');
        return txt.includes('Back to refurbishment detail') || txt.includes('Post Authorization');
    }

    async function waitForPostAuthorizationPage(timeout = 10000) {
        const start = Date.now();
        while ((Date.now() - start) < timeout) {
            if (abortFlag) throw new Error('Abort');
            if (isAlreadyOnPostAuthorizationPage()) return true;
            await sleep(80);
        }
        return isAlreadyOnPostAuthorizationPage();
    }

    function findSparePartsSidebarTab() {
        const all = document.querySelectorAll('*');
        const exactMatches = [];
        const fallbackMatches = [];
        for (const el of all) {
            if (['SCRIPT', 'STYLE', 'HTML', 'HEAD', 'BODY'].includes(el.tagName)) continue;
            const raw = (el.textContent || '').trim();
            if (!raw.includes('Spare parts')) continue;
            if (raw.includes('management') || raw.includes('Management')) continue;
            const childHas = Array.from(el.children).some(c => (c.textContent || '').includes('Spare parts'));
            if (childHas) continue;
            if (/^Spare parts(\s*\d+)?$/.test(raw)) {
                exactMatches.push(el);
            } else {
                fallbackMatches.push(el);
            }
        }
        return exactMatches[0] || fallbackMatches[0] || null;
    }

    async function openSparePartsTabAndWaitSelectAll(timeout = 9000) {
        const start = Date.now();
        let lastClickAt = 0;
        while ((Date.now() - start) < timeout) {
            if (abortFlag) throw new Error('Abort');
            const selectAllEls = getDeepestElementsByText('Select all');
            if (selectAllEls.length > 0) {
                return { ok: true, btnSelectAll: selectAllEls[0] };
            }
            if ((Date.now() - lastClickAt) >= 1100) {
                const tabEl = findSparePartsSidebarTab();
                if (tabEl) {
                    forceClick(tabEl);
                    const clickableParent = tabEl.closest('li, a, button, [role="tab"], [role="button"]') || tabEl.parentElement;
                    if (clickableParent && clickableParent !== tabEl) {
                        try { clickableParent.click(); } catch (e) {}
                    }
                    lastClickAt = Date.now();
                }
            }
            await sleep(100);
        }
        const finalCheck = getDeepestElementsByText('Select all');
        if (finalCheck.length > 0) {
            return { ok: true, btnSelectAll: finalCheck[0] };
        }
        return { ok: false, btnSelectAll: null };
    }

    function isOrderAlreadyCompletedOnPage() {
        const txt = String((document.body && document.body.innerText) || '');
        if (/\bCompleted on\s+\d/i.test(txt) || /\bFertiggestellt am\s+\d/i.test(txt)) return true;
        if (txt.includes('Car position: Retail Ready') && txt.includes('Overview vehicle details') && !txt.includes('Edit damages and services')) return true;
        return false;
    }

    async function waitForEditButtonOrCompleted(timeout = 10000) {
        const start = Date.now();
        while ((Date.now() - start) < timeout) {
            if (abortFlag) throw new Error('Abort');
            const els = getDeepestElementsByText('Edit damages and services');
            if (els.length > 0) return { kind: 'edit', el: els[0] };
            if (isOrderAlreadyCompletedOnPage()) return { kind: 'completed', el: null };
            await sleep(80);
        }
        if (isOrderAlreadyCompletedOnPage()) return { kind: 'completed', el: null };
        return { kind: 'missing', el: null };
    }

    function findStockLinkOnSearchPage(stockId) {
        const sid = String(stockId || '').trim().toUpperCase();
        const links = Array.from(document.querySelectorAll('a[href*="/refurbishment/"]'))
            .filter(a => isRefurbDetailUrl(a.href || a.getAttribute('href') || ''));
        if (!links.length) return null;
        for (let i = 0; i < links.length; i++) {
            const row = links[i].closest('tr, [role="row"], article, li, [class*="row"], [class*="card"]') || links[i].parentElement;
            const txt = String((row && row.textContent) || links[i].textContent || '').toUpperCase().replace(/\s+/g, '');
            if (txt.includes(sid)) return links[i];
        }
        return null;
    }

    function updateJobInQueue(stockId, patch) {
        const q = readJson(KEY_QUEUE, []);
        let changed = false;
        for (let i = 0; i < q.length; i++) {
            if (q[i] && q[i].stockId === stockId) {
                Object.assign(q[i], patch);
                changed = true;
            }
        }
        if (changed) writeJson(KEY_QUEUE, q);
    }

    function removeJobFromQueue(stockId) {
        const q = readJson(KEY_QUEUE, []);
        const next = q.filter(j => j && j.stockId !== stockId);
        writeJson(KEY_QUEUE, next);
        return next;
    }

    function safeNavigateTo(url) {
        isNavigating = true;
        setTimeout(() => { isNavigating = false; }, 5000);
        location.assign(url);
    }

    async function finishJobAndAdvance(job, status, message, indexLabel) {
        const sid = job.stockId;
        const remaining = removeJobFromQueue(sid);
        pushResultEvent(sid, status, message, job);
        playTone(status === 'ok');

        if (remaining.length > 0 && !abortFlag) {
            const nextJob = remaining[0];
            currentSid = nextJob.stockId;
            const nextUrl = nextJob.carolUrl || ('https://carol.autohero.com/en-GB/refurbishment?rsv=' + encodeURIComponent(nextJob.stockId));
            updateJobInQueue(nextJob.stockId, { status: 'running', phase: 'opened' });
            setHudStatus(`offen: ${remaining.length} · ${nextJob.stockId}`, 'Öffne nächsten Auftrag…', '#00ffcc', true);
            safeNavigateTo(nextUrl);
            return 'navigating';
        }

        currentSid = '';
        if (status === 'ok') {
            setHudStatus(`FERTIG · ${sid}`, `✓ ${message}`, '#00ff00', false);
        } else {
            setHudStatus(`FEHLER · ${sid}`, `✗ ${message}`, '#f85149', false);
        }
        return 'all_done';
    }

    async function processCurrentJob(job, indexLabel) {
        const sid = job.stockId;
        currentSid = sid;

        const targetUrl = (job.carolUrl && /^https?:\/\//i.test(job.carolUrl))
            ? job.carolUrl
            : ('https://carol.autohero.com/en-GB/refurbishment?rsv=' + encodeURIComponent(sid));
        const targetUuid = extractRefurbUuid(targetUrl);
        const currentUuid = extractRefurbUuid(location.href);

        if (targetUuid) {
            if (currentUuid !== targetUuid) {
                updateJobInQueue(sid, { status: 'running', phase: 'opened' });
                setHudStatus(`${indexLabel} · ${sid}`, 'Öffne Auftrag in Carol…', '#00ffcc', true);
                safeNavigateTo(targetUrl);
                return 'navigating';
            }
        } else if (!isRefurbDetailUrl(location.href)) {
            if (!location.href.toUpperCase().includes(sid)) {
                const searchUrl = 'https://carol.autohero.com/en-GB/refurbishment?rsv=' + encodeURIComponent(sid);
                updateJobInQueue(sid, { status: 'running', phase: 'opened' });
                setHudStatus(`${indexLabel} · ${sid}`, 'Suche Auftrag in Carol…', '#00ffcc', true);
                safeNavigateTo(searchUrl);
                return 'navigating';
            }

            setHudStatus(`${indexLabel} · ${sid}`, 'Warte auf Suchergebnis in Carol…', '#00ffcc', true);
            let foundLink = null;
            for (let t = 0; t < 40; t++) {
                await sleep(150);
                if (isRefurbDetailUrl(location.href)) break;
                foundLink = findStockLinkOnSearchPage(sid);
                if (foundLink) break;
            }

            if (!isRefurbDetailUrl(location.href)) {
                if (!foundLink) {
                    return await finishJobAndAdvance(job, 'error', 'Auftrag in Carol nicht gefunden', indexLabel);
                }
                const href = foundLink.href || foundLink.getAttribute('href');
                if (href && isRefurbDetailUrl(href)) {
                    updateJobInQueue(sid, { phase: 'detail_loaded' });
                    safeNavigateTo(href);
                    return 'navigating';
                }
                forceClick(foundLink);
                await sleep(800);
            }
        }

        if (!isAlreadyOnPostAuthorizationPage()) {
            setHudStatus(`${indexLabel} · ${sid}`, 'Prüfe Auftrag / Öffne Edit damages...', '#00ffcc', true);
            const mainCheck = await waitForEditButtonOrCompleted(10000);
            if (mainCheck.kind === 'completed') {
                setHudStatus(`${indexLabel} · ${sid}`, '⏭ Bereits Completed — überspringe…', '#00ff00', true);
                return await finishJobAndAdvance(job, 'ok', 'Bereits Completed (übersprungen)', indexLabel);
            }
            if (!mainCheck.el) {
                return await finishJobAndAdvance(job, 'error', 'Edit-Button nicht gefunden', indexLabel);
            }
            forceClick(mainCheck.el);
            setHudStatus(`${indexLabel} · ${sid}`, 'Warte auf Post Authorization...', '#00ffcc', true);
            await waitForPostAuthorizationPage(8000);
            await sleep(300);
        }

        setHudStatus(`${indexLabel} · ${sid}`, 'Öffne Spare Parts...', '#00ffcc', true);
        const spOpen = await openSparePartsTabAndWaitSelectAll(9000);
        if (!spOpen.ok || !spOpen.btnSelectAll) {
            return await finishJobAndAdvance(job, 'error', 'Spare Parts / Select All nicht gefunden', indexLabel);
        }
        const btnSelectAll = spOpen.btnSelectAll;

        await sleep(150);

        if (checkAllHandedOut()) {
            return await finishJobAndAdvance(job, 'ok', 'Bereits Herausgegeben', indexLabel);
        }

        setHudStatus(`${indexLabel} · ${sid}`, 'Setze alle Teile auf Handed out...', '#ffff00', true);
        forceClick(btnSelectAll);
        await sleep(250);
        forceAllHandedOut();
        await sleep(150);

        const btnApply = await waitForElement('Apply', 2500);
        if (btnApply) forceClick(btnApply);

        let verified = false;
        for (let i = 0; i < 20; i++) {
            await sleep(100);
            if (checkAllHandedOut()) { verified = true; break; }
            if (i === 8) {
                forceAllHandedOut();
                const retryApply = getDeepestElementsByText('Apply')[0];
                if (retryApply) forceClick(retryApply);
            }
        }

        if (!verified) {
            return await finishJobAndAdvance(job, 'error', 'Handed out konnte nicht gesetzt werden', indexLabel);
        }

        setHudStatus(`${indexLabel} · ${sid}`, 'Speichere (Submit)...', '#00ffcc', true);
        const btnSubmit = await waitForElement('Submit', 5000);
        if (!btnSubmit) {
            return await finishJobAndAdvance(job, 'error', 'Submit-Button nicht gefunden', indexLabel);
        }
        forceClick(btnSubmit);
        await sleep(200);

        setHudStatus(`${indexLabel} · ${sid}`, 'Bestätige Freigabe (Confirm)...', '#00ffcc', true);
        const confirmed = await waitAndClickConfirmButton(6000);
        if (confirmed) {
            setHudStatus(`${indexLabel} · ${sid}`, 'Schließe Dialog...', '#00ffcc', true);
            await clickSavedDialogCloseFast(3500);
        }

        return await finishJobAndAdvance(job, 'ok', 'Herausgegeben', indexLabel);
    }

    async function checkAndRunQueue() {
        if (workerBusy || isNavigating) return;
        const abortObj = readJson(KEY_ABORT, null);
        if (abortObj && Number(abortObj.ts || 0) > (Date.now() - 4000)) {
            writeJson(KEY_QUEUE, []);
            return;
        }

        const currentQ = readJson(KEY_QUEUE, []);
        if (!Array.isArray(currentQ) || !currentQ.length) return;

        try { sessionStorage.setItem('dpl_tl_worker_tab', '1'); } catch (e) {}
        workerBusy = true;
        abortFlag = false;

        try {
            const nextJob = currentQ[0];
            if (!nextJob || !nextJob.stockId) {
                writeJson(KEY_QUEUE, currentQ.slice(1));
                workerBusy = false;
                return;
            }

            const result = await processCurrentJob(nextJob, `offen: ${currentQ.length}`);
            if (result === 'navigating') {
                return;
            }
        } catch (err) {
            if (err && err.message === 'Abort') {
                currentSid = '';
                setHudStatus('GESTOPPT', 'Durchlauf abgebrochen', '#f85149', false);
            } else {
                setHudStatus('FEHLER', String((err && err.message) || err), '#f85149', false);
            }
        } finally {
            if (!isNavigating) {
                workerBusy = false;
            }
        }
    }

    window.addEventListener('message', (ev) => {
        const d = ev && ev.data;
        if (!d || d.source !== 'dpl-tagesliste-hud') return;
        try { sessionStorage.setItem('dpl_tl_worker_tab', '1'); } catch (e) {}

        if (d.cmd === 'enqueue_jobs' && d.payload && Array.isArray(d.payload.jobs)) {
            const incoming = d.payload.jobs;
            const bySid = {};
            const freshQueue = [];
            for (let j = 0; j < incoming.length; j++) {
                const inc = incoming[j];
                if (!inc || !inc.stockId) continue;
                const sid = String(inc.stockId).trim().toUpperCase();
                if (!bySid[sid]) {
                    const job = {
                        stockId: sid,
                        dateStr: inc.dateStr || '',
                        carolUrl: inc.carolUrl || ('https://carol.autohero.com/en-GB/refurbishment?rsv=' + encodeURIComponent(sid)),
                        model: inc.model || '',
                        rows: inc.rows || [],
                        row: inc.row || 0,
                        status: 'pending',
                        phase: 'init',
                        addedAt: Date.now()
                    };
                    bySid[sid] = job;
                    freshQueue.push(job);
                }
            }
            writeJson(KEY_ABORT, { ts: 0 });
            writeJson(KEY_QUEUE, freshQueue);
            checkAndRunQueue();
        }

        if (d.cmd === 'stop_all') {
            abortFlag = true;
            workerBusy = false;
            isNavigating = false;
            writeJson(KEY_QUEUE, []);
            setHudStatus('GESTOPPT', 'Durchlauf durch HUD gestoppt', '#f85149', false);
        }
    });

    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && isWorkerTab()) {
            abortFlag = true;
            workerBusy = false;
            isNavigating = false;
            writeJson(KEY_QUEUE, []);
            setHudStatus('GESTOPPT', 'Durchlauf mit ESC abgebrochen', '#f85149', false);
        }
    });

    if (typeof GM_addValueChangeListener === 'function') {
        GM_addValueChangeListener(KEY_QUEUE, () => {
            if (!workerBusy && !isNavigating) checkAndRunQueue();
        });
    }

    setInterval(() => {
        if (isWorkerTab()) {
            ensureHud();
            sendHeartbeat();
        }
        if (!workerBusy && !isNavigating) {
            const q = readJson(KEY_QUEUE, []);
            if (Array.isArray(q) && q.length > 0) {
                checkAndRunQueue();
            }
        }
    }, 1000);

    setTimeout(() => {
        if (isWorkerTab()) {
            ensureHud();
            sendHeartbeat();
        }
        checkAndRunQueue();
    }, 300);
})();
