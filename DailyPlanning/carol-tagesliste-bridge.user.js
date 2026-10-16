// ==UserScript==
// @name         Carol Tagesliste Herausgegeben Bridge
// @namespace    http://tampermonkey.net/
// @version      1.0
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

    const KEY_QUEUE = 'dpl_tl_carol_queue_v1';
    const KEY_STATE = 'dpl_tl_carol_state_v1';
    const KEY_EVENTS = 'dpl_tl_carol_events_v1';
    const KEY_ABORT = 'dpl_tl_carol_abort_v1';

    const isCarolHost = location.hostname.includes('carol.autohero.com');
    const isWmsSingleAuto = isCarolHost && location.href.includes('arasaka_auto=1');

    function readJson(key, fallback) {
        try {
            const raw = GM_getValue(key, '');
            if (!raw) return fallback;
            return JSON.parse(raw);
        } catch (e) {
            return fallback;
        }
    }

    function writeJson(key, val) {
        try {
            GM_setValue(key, JSON.stringify(val));
        } catch (e) {}
    }

    if (!isCarolHost) {
        let lastEventCursor = 0;

        function forwardStateToHud() {
            const st = readJson(KEY_STATE, null);
            if (st && (Date.now() - Number(st.ts || 0)) < 6000) {
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
                        bySid[sid].rows = inc.rows || bySid[sid].rows;
                        bySid[sid].row = inc.row || bySid[sid].row;
                    } else {
                        const job = {
                            stockId: sid,
                            carolUrl: inc.carolUrl || ('https://carol.autohero.com/en-GB/refurbishment?rsv=' + encodeURIComponent(sid)),
                            model: inc.model || '',
                            rows: inc.rows || [],
                            row: inc.row || 0,
                            status: 'pending',
                            phase: 'init',
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

    function pushResultEvent(stockId, status, message) {
        const ev = {
            eventId: stockId + '_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
            stockId: stockId,
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
                    padding: 10px 12px;
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
            `;
            document.body.appendChild(hudEl);
            const stopBtn = document.getElementById('dpl-tl-stop-btn');
            if (stopBtn) {
                stopBtn.addEventListener('click', () => {
                    abortFlag = true;
                    workerBusy = false;
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
        const needle = String(text || '').toLowerCase();
        for (const el of document.querySelectorAll('*')) {
            if (['SCRIPT', 'STYLE', 'HTML', 'HEAD', 'BODY'].includes(el.tagName)) continue;
            if ((el.textContent || '').toLowerCase().includes(needle)) {
                const childHas = Array.from(el.children).some(c => (c.textContent || '').toLowerCase().includes(needle));
                if (!childHas) results.push(el);
            }
        }
        return results;
    }

    async function waitForAnyElement(texts, timeout = 10000, pickLast = false) {
        const list = Array.isArray(texts) ? texts : [texts];
        let passed = 0;
        while (passed < timeout) {
            if (abortFlag) throw new Error('Abort');
            for (let i = 0; i < list.length; i++) {
                const els = getDeepestElementsByText(list[i]);
                if (els.length > 0) return pickLast ? els[els.length - 1] : els[0];
            }
            await sleep(350);
            passed += 350;
        }
        return null;
    }

    function forceClick(el) {
        if (!el) return;
        try { el.scrollIntoView({ block: 'center' }); } catch (e) {}
        try {
            el.style.outline = '2px solid #00ffcc';
            setTimeout(() => { try { el.style.outline = ''; } catch (e2) {} }, 350);
        } catch (e) {}
        try { el.click(); } catch (e) {}
        if (el.parentElement) {
            setTimeout(() => { try { el.parentElement.click(); } catch (e) {} }, 40);
        }
    }

    function isHandedOutOption(text) {
        const t = String(text || '').trim().toLowerCase();
        return t.includes('handed out') || t.includes('herausgegeben');
    }

    function checkAllHandedOut() {
        const selects = document.querySelectorAll('select');
        let relevant = 0;
        for (const sel of selects) {
            const hasOpt = Array.from(sel.options).some(o => isHandedOutOption(o.text));
            if (!hasOpt) continue;
            relevant++;
            if (!isHandedOutOption(sel.options[sel.selectedIndex]?.text)) return false;
        }
        return relevant > 0;
    }

    function forceAllHandedOut() {
        let count = 0;
        for (const sel of document.querySelectorAll('select')) {
            const opt = Array.from(sel.options).find(o => isHandedOutOption(o.text));
            if (opt) {
                count++;
                if (!isHandedOutOption(sel.options[sel.selectedIndex]?.text)) {
                    sel.value = opt.value;
                    sel.dispatchEvent(new Event('change', { bubbles: true }));
                }
            }
        }
        return count;
    }

    function findExactCloseButton() {
        const candidates = [];
        for (const el of document.querySelectorAll('button, [role="button"], a')) {
            const txt = (el.textContent || '').trim().toLowerCase();
            if (txt === 'close' || txt === 'schließen' || txt === 'schliessen') candidates.push(el);
        }
        return candidates.length ? candidates[candidates.length - 1] : null;
    }

    async function clickSavedDialogClose(timeout = 10000) {
        let passed = 0;
        while (passed < timeout) {
            if (abortFlag) throw new Error('Abort');
            const btn = findExactCloseButton();
            if (btn) {
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

    function pageContainsStockId(stockId) {
        const sid = String(stockId || '').trim().toUpperCase();
        if (!sid) return false;
        const bodyTxt = String((document.body && document.body.innerText) || '').toUpperCase().replace(/\s+/g, ' ');
        return bodyTxt.includes(sid);
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
        return links.length === 1 ? links[0] : null;
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

    async function processJob(job, indexLabel) {
        const sid = job.stockId;
        currentSid = sid;

        if (job.phase === 'confirming') {
            setHudStatus(`${indexLabel} · ${sid}`, 'Schließe Dialog nach Reload…', '#00ffcc');
            await sleep(1000);
            await clickSavedDialogClose(3500);
            removeJobFromQueue(sid);
            pushResultEvent(sid, 'ok', 'Herausgegeben');
            playTone(true);
            setHudStatus(`${indexLabel} · ${sid}`, '✓ Herausgegeben (Spalte F grün)', '#00ff00');
            await sleep(600);
            return;
        }

        if (!isRefurbDetailUrl(location.href) || !pageContainsStockId(sid)) {
            if (job.carolUrl && isRefurbDetailUrl(job.carolUrl) && job.navigatedUrl !== job.carolUrl) {
                updateJobInQueue(sid, { status: 'running', navigatedUrl: job.carolUrl });
                setHudStatus(`${indexLabel} · ${sid}`, 'Öffne Carol-Link aus Spalte F…', '#00ffcc');
                location.assign(job.carolUrl);
                return;
            }

            if (!isRefurbDetailUrl(location.href)) {
                const wantSearchUrl = 'https://carol.autohero.com/en-GB/refurbishment?rsv=' + encodeURIComponent(sid);
                if (!location.href.toUpperCase().includes(sid)) {
                    updateJobInQueue(sid, { status: 'running', navigatedUrl: wantSearchUrl });
                    setHudStatus(`${indexLabel} · ${sid}`, 'Suche Stock-ID in Carol…', '#00ffcc');
                    location.assign(wantSearchUrl);
                    return;
                }

                setHudStatus(`${indexLabel} · ${sid}`, 'Warte auf Suchergebnis in Carol…', '#00ffcc');
                let foundLink = null;
                for (let t = 0; t < 30; t++) {
                    await sleep(400);
                    if (isRefurbDetailUrl(location.href)) break;
                    foundLink = findStockLinkOnSearchPage(sid);
                    if (foundLink) break;
                }

                if (!isRefurbDetailUrl(location.href)) {
                    if (!foundLink) {
                        removeJobFromQueue(sid);
                        pushResultEvent(sid, 'error', 'Auftrag in Carol nicht gefunden');
                        playTone(false);
                        setHudStatus(`${indexLabel} · ${sid}`, '✗ Nicht in Carol gefunden', '#f85149');
                        await sleep(1000);
                        return;
                    }
                    const href = foundLink.href || foundLink.getAttribute('href');
                    if (href && isRefurbDetailUrl(href)) {
                        location.assign(href);
                        return;
                    }
                    forceClick(foundLink);
                    await sleep(2000);
                }
            }
        }

        let sidVerified = false;
        for (let v = 0; v < 20; v++) {
            if (pageContainsStockId(sid)) { sidVerified = true; break; }
            await sleep(400);
        }
        if (!sidVerified) {
            const fallbackSearch = 'https://carol.autohero.com/en-GB/refurbishment?rsv=' + encodeURIComponent(sid);
            if (job.navigatedUrl !== fallbackSearch) {
                updateJobInQueue(sid, { status: 'running', carolUrl: fallbackSearch, navigatedUrl: fallbackSearch });
                location.assign(fallbackSearch);
                return;
            }
            removeJobFromQueue(sid);
            pushResultEvent(sid, 'error', 'Stock-ID stimmt nicht mit geladenem Auftrag überein');
            playTone(false);
            return;
        }

        setHudStatus(`${indexLabel} · ${sid}`, 'Öffne Edit damages and services…', '#00ffcc');
        const btnEdit = await waitForAnyElement(['Edit damages and services', 'Schäden und Services bearbeiten'], 12000);
        if (!btnEdit) {
            removeJobFromQueue(sid);
            pushResultEvent(sid, 'error', 'Edit-Button nicht gefunden');
            playTone(false);
            setHudStatus(`${indexLabel} · ${sid}`, '✗ Edit-Button fehlt', '#f85149');
            await sleep(1000);
            return;
        }
        forceClick(btnEdit);

        setHudStatus(`${indexLabel} · ${sid}`, 'Navigiere zu Spare Parts…', '#00ffcc');
        await sleep(2000);
        const btnSpareParts = await waitForAnyElement(['Spare parts', 'Ersatzteile'], 10000);
        if (!btnSpareParts) {
            removeJobFromQueue(sid);
            pushResultEvent(sid, 'error', 'Tab Spare Parts nicht gefunden');
            playTone(false);
            setHudStatus(`${indexLabel} · ${sid}`, '✗ Spare Parts fehlt', '#f85149');
            await sleep(1000);
            return;
        }
        forceClick(btnSpareParts);

        setHudStatus(`${indexLabel} · ${sid}`, 'Wähle alle Positionen aus (Select all)…', '#00ffcc');
        await sleep(1000);
        const btnSelectAll = await waitForAnyElement(['Select all', 'Alle auswählen'], 8000);
        if (!btnSelectAll) {
            removeJobFromQueue(sid);
            pushResultEvent(sid, 'error', 'Keine Spare Parts / Select All nicht gefunden');
            playTone(false);
            setHudStatus(`${indexLabel} · ${sid}`, '✗ Select All fehlt', '#f85149');
            await sleep(1000);
            return;
        }
        forceClick(btnSelectAll);

        setHudStatus(`${indexLabel} · ${sid}`, 'Setze alle Spare Parts auf Handed out…', '#ffff00');
        await sleep(1000);
        forceAllHandedOut();

        setHudStatus(`${indexLabel} · ${sid}`, 'Übernehme Änderungen (Apply)…', '#00ffcc');
        await sleep(1000);
        const btnApply = await waitForAnyElement(['Apply', 'Übernehmen', 'Anwenden'], 6000);
        if (btnApply) forceClick(btnApply);

        setHudStatus(`${indexLabel} · ${sid}`, 'Verifiziere Handed out Status…', '#ffff00');
        let verified = false;
        for (let i = 0; i < 6; i++) {
            await sleep(900);
            if (checkAllHandedOut()) { verified = true; break; }
        }

        if (!verified) {
            setHudStatus(`${indexLabel} · ${sid}`, 'Rettungsversuch: Setze Handed out erneut…', '#ffaa00');
            forceAllHandedOut();
            await sleep(800);
            const btnApply2 = await waitForAnyElement(['Apply', 'Übernehmen', 'Anwenden'], 3000);
            if (btnApply2) forceClick(btnApply2);
            await sleep(2000);
            verified = checkAllHandedOut();
        }

        if (!verified) {
            removeJobFromQueue(sid);
            pushResultEvent(sid, 'error', 'Nicht alle Positionen konnten auf Handed out gestellt werden');
            playTone(false);
            setHudStatus(`${indexLabel} · ${sid}`, '✗ Handed out nicht bestätigt', '#f85149');
            await sleep(1200);
            return;
        }

        updateJobInQueue(sid, { phase: 'confirming' });
        setHudStatus(`${indexLabel} · ${sid}`, 'Übermittle Daten (Submit)…', '#00ffcc');
        const btnSubmit = await waitForAnyElement(['Submit', 'Absenden', 'Speichern'], 10000);
        if (!btnSubmit) {
            removeJobFromQueue(sid);
            pushResultEvent(sid, 'error', 'Submit-Button nicht gefunden');
            playTone(false);
            return;
        }
        forceClick(btnSubmit);

        setHudStatus(`${indexLabel} · ${sid}`, 'Bestätige Freigabe (Confirm)…', '#00ffcc');
        await sleep(1000);
        const btnConfirm = await waitForAnyElement(['Confirm', 'Bestätigen'], 10000);
        if (btnConfirm) forceClick(btnConfirm);

        setHudStatus(`${indexLabel} · ${sid}`, 'Schließe Dialog (Close)…', '#00ffcc');
        await sleep(2000);
        const closed = await clickSavedDialogClose(10000);
        if (!closed) {
            const btnClose = await waitForAnyElement(['Close', 'Schließen'], 4000, true);
            if (btnClose) btnClose.click();
        }
        await sleep(700);

        removeJobFromQueue(sid);
        pushResultEvent(sid, 'ok', 'Herausgegeben');
        playTone(true);
        setHudStatus(`${indexLabel} · ${sid}`, '✓ Herausgegeben · Spalte F grün markiert!', '#00ff00');
        await sleep(600);
    }

    async function runQueueLoop() {
        if (workerBusy) return;
        const q = readJson(KEY_QUEUE, []);
        if (!Array.isArray(q) || !q.length) return;

        workerBusy = true;
        abortFlag = false;

        try {
            while (!abortFlag) {
                const abortObj = readJson(KEY_ABORT, null);
                if (abortObj && Number(abortObj.ts || 0) > (Date.now() - 4000)) {
                    abortFlag = true;
                    break;
                }
                const currentQ = readJson(KEY_QUEUE, []);
                if (!Array.isArray(currentQ) || !currentQ.length) break;

                const nextJob = currentQ[0];
                if (!nextJob || !nextJob.stockId) {
                    writeJson(KEY_QUEUE, currentQ.slice(1));
                    continue;
                }

                await processJob(nextJob, `offen: ${currentQ.length}`);
            }

            const rem = readJson(KEY_QUEUE, []);
            if (!rem.length && !abortFlag) {
                currentSid = '';
                setHudStatus('FERTIG', 'Alle Stock-IDs auf Herausgegeben gestellt · Warte auf HUD…', '#00ff00');
            }
        } catch (err) {
            if (err && err.message === 'Abort') {
                currentSid = '';
                setHudStatus('GESTOPPT', 'Durchlauf abgebrochen', '#f85149');
            } else {
                setHudStatus('FEHLER', String((err && err.message) || err), '#f85149');
            }
        } finally {
            workerBusy = false;
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
                        carolUrl: inc.carolUrl || ('https://carol.autohero.com/en-GB/refurbishment?rsv=' + encodeURIComponent(sid)),
                        model: inc.model || '',
                        rows: inc.rows || [],
                        row: inc.row || 0,
                        status: 'pending',
                        phase: 'init',
                        addedAt: Date.now()
                    };
                    bySid[sid] = job;
                    merged.push(job);
                }
            }
            writeJson(KEY_ABORT, { ts: 0 });
            writeJson(KEY_QUEUE, merged);
            runQueueLoop();
        }
        if (d.cmd === 'stop_all') {
            abortFlag = true;
            workerBusy = false;
            writeJson(KEY_QUEUE, []);
            setHudStatus('GESTOPPT', 'Durchlauf durch HUD gestoppt', '#f85149');
        }
    });

    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && workerBusy) {
            abortFlag = true;
            workerBusy = false;
            writeJson(KEY_QUEUE, []);
            setHudStatus('GESTOPPT', 'Durchlauf mit ESC abgebrochen', '#f85149');
        }
    });

    if (typeof GM_addValueChangeListener === 'function') {
        GM_addValueChangeListener(KEY_QUEUE, () => {
            if (!workerBusy) runQueueLoop();
        });
    }

    setInterval(() => {
        ensureHud();
        publishState();
        if (!workerBusy) {
            const q = readJson(KEY_QUEUE, []);
            if (Array.isArray(q) && q.length > 0) {
                runQueueLoop();
            }
        }
    }, 1200);

    setTimeout(() => {
        ensureHud();
        publishState();
        runQueueLoop();
    }, 800);
})();
