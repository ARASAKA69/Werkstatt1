// ==UserScript==
// @name         Carol Tagesliste Herausgegeben Bridge
// @namespace    http://tampermonkey.net/
// @version      1.3
// @description  Empfängt Stock-IDs vom Tagesliste-HUD, stellt alle Spare Parts in Carol auf Handed out (Herausgegeben) und meldet Erfolg für grüne Spalte F zurück
// @author       ARASAKA
// @match        *://carol.autohero.com/*
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

    try {
        if (typeof GM_setValue === 'function') {
            GM_setValue('dpl_tl_carol_queue_v1', '[]');
        }
        localStorage.removeItem('dpl_tl_carol_queue_v1');
    } catch (e0) {}

    const KEY_QUEUE = 'dpl_tl_carol_queue_v2';
    const KEY_STATE = 'dpl_tl_carol_state_v2';
    const KEY_EVENTS = 'dpl_tl_carol_events_v2';
    const KEY_ABORT = 'dpl_tl_carol_abort_v2';

    const isCarolHost = location.hostname.includes('carol.autohero.com');
    const isWmsSingleAuto = isCarolHost && location.href.includes('arasaka_auto=1');

    function readJson(key, fallback) {
        try {
            if (typeof GM_getValue === 'function') {
                const rawGm = GM_getValue(key, '');
                if (rawGm) return JSON.parse(rawGm);
            }
        } catch (e1) {}
        try {
            const rawLs = localStorage.getItem(key);
            if (rawLs) return JSON.parse(rawLs);
        } catch (e2) {}
        return fallback;
    }

    function writeJson(key, val) {
        const str = JSON.stringify(val);
        try {
            if (typeof GM_setValue === 'function') {
                GM_setValue(key, str);
            }
        } catch (e1) {}
        try {
            localStorage.setItem(key, str);
        } catch (e2) {}
    }

    if (!isCarolHost) {
        let lastEventCursor = 0;

        function forwardStateToHud() {
            const st = readJson(KEY_STATE, null);
            if (st && (Date.now() - Number(st.ts || 0)) < 20000) {
                window.postMessage({
                    source: 'dpl-tagesliste-bridge',
                    type: 'state',
                    busy: !!st.busy,
                    currentSid: st.currentSid || '',
                    stepText: st.stepText || '',
                    ts: st.ts
                }, '*');
            }
            const evs = readJson(KEY_EVENTS, []);
            if (Array.isArray(evs) && evs.length) {
                for (let i = 0; i < evs.length; i++) {
                    const ev = evs[i];
                    if (!ev || !ev.stockId) continue;
                    if (Number(ev.ts || 0) <= lastEventCursor) continue;
                    window.postMessage({
                        source: 'dpl-tagesliste-bridge',
                        type: 'item_result',
                        eventId: ev.eventId,
                        stockId: ev.stockId,
                        dateStr: ev.dateStr || '',
                        rows: ev.rows || [],
                        status: ev.status || 'ok',
                        message: ev.message || '',
                        ts: ev.ts
                    }, '*');
                }
                const maxTs = evs.reduce((m, x) => Math.max(m, Number((x && x.ts) || 0)), lastEventCursor);
                lastEventCursor = maxTs;
            }
        }

        window.addEventListener('message', (ev) => {
            const d = ev && ev.data;
            if (!d || d.source !== 'dpl-tagesliste-hud') return;

            if (d.cmd === 'ping') {
                forwardStateToHud();
                return;
            }

            if (d.cmd === 'stop_all') {
                writeJson(KEY_ABORT, { ts: Date.now() });
                writeJson(KEY_QUEUE, []);
                return;
            }

            if (d.cmd === 'enqueue_jobs') {
                const incoming = (d.payload && d.payload.jobs) || [];
                if (!incoming.length) return;
                const currentQ = readJson(KEY_QUEUE, []);
                const bySid = {};
                const merged = [];
                for (let i = 0; i < currentQ.length; i++) {
                    const item = currentQ[i];
                    if (!item || !item.stockId) continue;
                    if (item.status === 'done' || item.status === 'error') continue;
                    bySid[item.stockId] = item;
                    merged.push(item);
                }
                for (let j = 0; j < incoming.length; j++) {
                    const inc = incoming[j];
                    if (!inc || !inc.stockId) continue;
                    const sid = String(inc.stockId).trim().toUpperCase();
                    if (bySid[sid]) {
                        bySid[sid].carolUrl = inc.carolUrl || bySid[sid].carolUrl;
                        bySid[sid].dateStr = inc.dateStr || bySid[sid].dateStr || '';
                        bySid[sid].rows = inc.rows || bySid[sid].rows;
                        bySid[sid].row = inc.row || bySid[sid].row;
                    } else {
                        const job = {
                            stockId: sid,
                            dateStr: inc.dateStr || '',
                            carolUrl: inc.carolUrl || ('https://carol.autohero.com/en-GB/refurbishment?rsv=' + encodeURIComponent(sid)),
                            model: inc.model || '',
                            rows: inc.rows || [],
                            row: inc.row || 0,
                            status: 'pending',
                            phase: 'init',
                            navCount: 0,
                            addedAt: Date.now()
                        };
                        bySid[sid] = job;
                        merged.push(job);
                    }
                }
                writeJson(KEY_ABORT, { ts: 0 });
                writeJson(KEY_QUEUE, merged);
            }
        });

        if (typeof GM_addValueChangeListener === 'function') {
            GM_addValueChangeListener(KEY_STATE, () => forwardStateToHud());
            GM_addValueChangeListener(KEY_EVENTS, () => forwardStateToHud());
        }
        setInterval(forwardStateToHud, 500);
        return;
    }

    if (isWmsSingleAuto) return;

    let workerBusy = false;
    let isNavigating = false;
    let abortFlag = false;
    let currentSid = '';
    let currentStepText = 'Warte auf Stock-IDs vom Tagesliste-HUD…';
    let hudEl = null;

    function publishState() {
        const statePayload = {
            busy: workerBusy,
            currentSid: currentSid,
            stepText: currentStepText,
            ts: Date.now()
        };
        writeJson(KEY_STATE, statePayload);
        try {
            if (window.opener && !window.opener.closed) {
                window.opener.postMessage({
                    source: 'dpl-tagesliste-bridge',
                    type: 'state',
                    busy: workerBusy,
                    currentSid: currentSid,
                    stepText: currentStepText,
                    ts: statePayload.ts
                }, '*');
            }
        } catch (e) {}
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
        try {
            if (window.opener && !window.opener.closed) {
                window.opener.postMessage({
                    source: 'dpl-tagesliste-bridge',
                    type: 'item_result',
                    eventId: ev.eventId,
                    stockId: ev.stockId,
                    dateStr: ev.dateStr,
                    rows: ev.rows,
                    status: ev.status,
                    message: ev.message,
                    ts: ev.ts
                }, '*');
            }
        } catch (e) {}
    }

    function ensureHud() {
        if (!document.body) return;
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
                    <div class="dpl-tl-sid" id="dpl-tl-sid-txt">BEREIT</div>
                    <div class="dpl-tl-step" id="dpl-tl-step-txt">Warte auf Stock-IDs vom HUD…</div>
                </div>
                <div class="dpl-tl-foot">
                    <span>Bridge v1.3 · [ESC] Stop</span>
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
                    setHudStatus('GESTOPPT', 'Warteschlange geleert', '#f85149');
                });
            }
        }
    }

    function setHudStatus(sidLabel, stepText, borderColor) {
        currentStepText = stepText || '';
        ensureHud();
        if (hudEl) {
            if (borderColor) hudEl.style.borderColor = borderColor;
            const sidEl = document.getElementById('dpl-tl-sid-txt');
            const stEl = document.getElementById('dpl-tl-step-txt');
            if (sidEl) sidEl.textContent = sidLabel || 'BEREIT';
            if (stEl) stEl.textContent = stepText || '';
        }
        publishState();
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
            gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.4);
            osc.start();
            osc.stop(ctx.currentTime + 0.4);
        } catch (e) {}
    }

    async function sleep(ms) {
        let elapsed = 0;
        while (elapsed < ms) {
            if (abortFlag) throw new Error('Abort');
            await new Promise(r => setTimeout(r, 100));
            elapsed += 100;
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

    async function waitForElement(text, timeout = 10000, pickLast = false) {
        let passed = 0;
        while (passed < timeout) {
            if (abortFlag) throw new Error('Abort');
            const els = getDeepestElementsByText(text);
            if (els.length > 0) return pickLast ? els[els.length - 1] : els[0];
            await sleep(500);
            passed += 500;
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
            }, 400);
        } catch (e) {}
        el.click();
        if (el.parentElement) setTimeout(() => el.parentElement.click(), 50);
    }

    function checkAllHandedOut() {
        const selects = document.querySelectorAll('select');
        let relevant = 0;
        for (const sel of selects) {
            const hasOpt = Array.from(sel.options).some(o => o.text.includes('Handed out'));
            if (!hasOpt) continue;
            relevant++;
            if (!sel.options[sel.selectedIndex]?.text.includes('Handed out')) return false;
        }
        return relevant > 0;
    }

    function forceAllHandedOut() {
        for (const sel of document.querySelectorAll('select')) {
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

    async function clickSavedDialogClose(timeout = 10000) {
        let passed = 0;
        while (passed < timeout) {
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
            await sleep(250);
            passed += 250;
        }
        return false;
    }

    function isRefurbDetailUrl(url) {
        return /\/refurbishment\/[0-9a-f-]{8,}/i.test(String(url || ''));
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
        return links[0];
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
        setTimeout(() => { isNavigating = false; }, 6000);
        location.assign(url);
    }

    async function processCurrentJob(job, indexLabel) {
        const sid = job.stockId;
        currentSid = sid;

        if (job.phase === 'confirming') {
            setHudStatus(`${indexLabel} · ${sid}`, 'Schließe Dialog…', '#00ffcc');
            await sleep(1000);
            await clickSavedDialogClose(3500);
            removeJobFromQueue(sid);
            pushResultEvent(sid, 'ok', 'Herausgegeben', job);
            playTone(true);
            setHudStatus(`${indexLabel} · ${sid}`, '✓ Herausgegeben (Spalte F grün)', '#00ff00');
            await sleep(700);
            return 'done_item';
        }

        if (job.phase === 'init') {
            const targetUrl = (job.carolUrl && /^https?:\/\//i.test(job.carolUrl))
                ? job.carolUrl
                : ('https://carol.autohero.com/en-GB/refurbishment?rsv=' + encodeURIComponent(sid));
            updateJobInQueue(sid, { status: 'running', phase: 'opened', navCount: (job.navCount || 0) + 1 });
            setHudStatus(`${indexLabel} · ${sid}`, 'Öffne Auftrag in Carol…', '#00ffcc');
            safeNavigateTo(targetUrl);
            return 'navigating';
        }

        if (!isRefurbDetailUrl(location.href)) {
            setHudStatus(`${indexLabel} · ${sid}`, 'Warte auf Suchergebnis in Carol…', '#00ffcc');
            let foundLink = null;
            for (let t = 0; t < 25; t++) {
                await sleep(400);
                if (isRefurbDetailUrl(location.href)) break;
                foundLink = findStockLinkOnSearchPage(sid);
                if (foundLink) break;
            }

            if (!isRefurbDetailUrl(location.href)) {
                if (!foundLink) {
                    removeJobFromQueue(sid);
                    pushResultEvent(sid, 'error', 'Auftrag in Carol nicht gefunden', job);
                    playTone(false);
                    setHudStatus(`${indexLabel} · ${sid}`, '✗ Nicht in Carol gefunden', '#f85149');
                    await sleep(1000);
                    return 'done_item';
                }
                const href = foundLink.href || foundLink.getAttribute('href');
                if (href && isRefurbDetailUrl(href)) {
                    updateJobInQueue(sid, { phase: 'detail_loaded' });
                    safeNavigateTo(href);
                    return 'navigating';
                }
                forceClick(foundLink);
                await sleep(2500);
            }
        }

        setHudStatus(`${indexLabel} · ${sid}`, 'Öffne Edit damages...', '#00ffcc');
        const btnEdit = await waitForElement('Edit damages and services', 15000);
        if (!btnEdit) {
            removeJobFromQueue(sid);
            pushResultEvent(sid, 'error', 'Edit-Button nicht gefunden', job);
            playTone(false);
            setHudStatus(`${indexLabel} · ${sid}`, '✗ Edit-Button nicht gefunden!', '#f85149');
            await sleep(1000);
            return 'done_item';
        }
        forceClick(btnEdit);

        setHudStatus(`${indexLabel} · ${sid}`, 'Navigiere zu Spare Parts...', '#00ffcc');
        await sleep(2000);
        const btnSpareParts = await waitForElement('Spare parts', 10000);
        if (!btnSpareParts) {
            removeJobFromQueue(sid);
            pushResultEvent(sid, 'error', 'Spare Parts nicht gefunden', job);
            playTone(false);
            setHudStatus(`${indexLabel} · ${sid}`, '✗ Spare Parts nicht gefunden!', '#f85149');
            await sleep(1000);
            return 'done_item';
        }
        forceClick(btnSpareParts);

        setHudStatus(`${indexLabel} · ${sid}`, 'Wähle alle Positionen aus...', '#00ffcc');
        await sleep(1000);
        const btnSelectAll = await waitForElement('Select all', 10000);
        if (!btnSelectAll) {
            removeJobFromQueue(sid);
            pushResultEvent(sid, 'error', 'Select All nicht gefunden', job);
            playTone(false);
            setHudStatus(`${indexLabel} · ${sid}`, '✗ Select All nicht gefunden!', '#f85149');
            await sleep(1000);
            return 'done_item';
        }
        forceClick(btnSelectAll);

        setHudStatus(`${indexLabel} · ${sid}`, 'Setze Dropdowns auf Handed out...', '#ffff00');
        await sleep(1000);
        forceAllHandedOut();

        setHudStatus(`${indexLabel} · ${sid}`, 'Übernehme Änderungen...', '#00ffcc');
        await sleep(1000);
        const btnApply = await waitForElement('Apply', 10000);
        if (btnApply) forceClick(btnApply);

        setHudStatus(`${indexLabel} · ${sid}`, 'Verifiziere Handed-Out Status...', '#ffff00');
        let verified = false;
        for (let i = 0; i < 6; i++) {
            await sleep(1000);
            if (checkAllHandedOut()) { verified = true; break; }
        }

        if (!verified) {
            setHudStatus(`${indexLabel} · ${sid}`, 'Rettungsprotokoll aktiv...', '#ffaa00');
            forceAllHandedOut();
            await sleep(1000);
            const btnApplyRetry = await waitForElement('Apply', 3000);
            if (btnApplyRetry) forceClick(btnApplyRetry);
            await sleep(2500);
            verified = checkAllHandedOut();
        }

        if (!verified) {
            removeJobFromQueue(sid);
            pushResultEvent(sid, 'error', 'Nicht alle Positionen konnten auf Handed out gestellt werden', job);
            playTone(false);
            setHudStatus(`${indexLabel} · ${sid}`, '✗ Handed out nicht bestätigt', '#f85149');
            await sleep(1200);
            return 'done_item';
        }

        updateJobInQueue(sid, { phase: 'confirming' });
        setHudStatus(`${indexLabel} · ${sid}`, 'Übermittle Daten...', '#00ffcc');
        const btnSubmit = await waitForElement('Submit', 10000);
        if (!btnSubmit) {
            removeJobFromQueue(sid);
            pushResultEvent(sid, 'error', 'Submit-Button nicht gefunden', job);
            playTone(false);
            return 'done_item';
        }
        forceClick(btnSubmit);

        setHudStatus(`${indexLabel} · ${sid}`, 'Bestätige Freigabe...', '#00ffcc');
        await sleep(1000);
        const btnConfirm = await waitForElement('Confirm', 10000);
        if (btnConfirm) forceClick(btnConfirm);

        setHudStatus(`${indexLabel} · ${sid}`, 'Schließe Dialog...', '#00ffcc');
        await sleep(2000);
        const dialogClosed = await clickSavedDialogClose(10000);
        if (!dialogClosed) {
            const btnClose = await waitForElement('Close', 5000, true);
            if (btnClose) btnClose.click();
        }
        await sleep(800);

        removeJobFromQueue(sid);
        pushResultEvent(sid, 'ok', 'Herausgegeben', job);
        playTone(true);
        setHudStatus(`${indexLabel} · ${sid}`, '✓ Herausgegeben · Spalte F grün markiert!', '#00ff00');
        await sleep(600);
        return 'done_item';
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

            const rem = readJson(KEY_QUEUE, []);
            if (!rem.length && !abortFlag) {
                currentSid = '';
                setHudStatus('FERTIG', 'Alle Stock-IDs auf Herausgegeben gestellt · Warte auf HUD…', '#00ff00');
            } else if (rem.length > 0 && !abortFlag) {
                workerBusy = false;
                setTimeout(checkAndRunQueue, 500);
                return;
            }
        } catch (err) {
            if (err && err.message === 'Abort') {
                currentSid = '';
                setHudStatus('GESTOPPT', 'Durchlauf abgebrochen', '#f85149');
            } else {
                setHudStatus('FEHLER', String((err && err.message) || err), '#f85149');
            }
        } finally {
            if (!isNavigating) {
                workerBusy = false;
            }
            publishState();
        }
    }

    window.addEventListener('message', (ev) => {
        const d = ev && ev.data;
        if (!d || d.source !== 'dpl-tagesliste-hud') return;
        if (d.cmd === 'enqueue_jobs' && d.payload && Array.isArray(d.payload.jobs)) {
            const incoming = d.payload.jobs;
            const currentQ = readJson(KEY_QUEUE, []);
            const bySid = {};
            const merged = currentQ.slice();
            for (let i = 0; i < merged.length; i++) {
                if (merged[i] && merged[i].stockId) bySid[merged[i].stockId] = merged[i];
            }
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
                        navCount: 0,
                        addedAt: Date.now()
                    };
                    bySid[sid] = job;
                    merged.push(job);
                }
            }
            writeJson(KEY_ABORT, { ts: 0 });
            writeJson(KEY_QUEUE, merged);
            checkAndRunQueue();
        }
        if (d.cmd === 'stop_all') {
            abortFlag = true;
            workerBusy = false;
            isNavigating = false;
            writeJson(KEY_QUEUE, []);
            setHudStatus('GESTOPPT', 'Durchlauf durch HUD gestoppt', '#f85149');
        }
    });

    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            abortFlag = true;
            workerBusy = false;
            isNavigating = false;
            writeJson(KEY_QUEUE, []);
            setHudStatus('GESTOPPT', 'Durchlauf mit ESC abgebrochen', '#f85149');
        }
    });

    if (typeof GM_addValueChangeListener === 'function') {
        GM_addValueChangeListener(KEY_QUEUE, () => {
            if (!workerBusy && !isNavigating) checkAndRunQueue();
        });
    }

    setInterval(() => {
        ensureHud();
        publishState();
        if (!workerBusy && !isNavigating) {
            const q = readJson(KEY_QUEUE, []);
            if (Array.isArray(q) && q.length > 0) {
                checkAndRunQueue();
            }
        }
    }, 1500);

    setTimeout(() => {
        ensureHud();
        publishState();
        checkAndRunQueue();
    }, 1200);
})();
