// ==UserScript==
// @name Carol-Automation
// @namespace http://tampermonkey.net/
// @version 3.0
// @description ARASAKA v3.0
// @author ARASAKA
// @match *://*/*
// @updateURL https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/arasaka-automation.user.js
// @downloadURL https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/arasaka-automation.user.js
// @grant GM_setValue
// @grant GM_getValue
// @grant GM_registerMenuCommand
// @grant GM_xmlhttpRequest
// ==/UserScript==

(function () {
    'use strict';

    const googleWebAppUrl = 'https://script.google.com/a/macros/auto1.com/s/AKfycbzFCxeD2nL8-km1izhQr1mLU8mhd1hCymE5jwt3B_eYP7Iqe17DllFAUuRgNseWH7Ya/exec';

    let abortMission = false;
    let hudElement = null;
    let isLocked = false;
    const speedMultiplier = GM_getValue('arasaka_speed', 1);
    const TOTAL_STEPS = 14;
    let hudLog = [];
    let currentStep = 0;
    let startTime = null;
    let timerInterval = null;

    GM_registerMenuCommand(
        `⚙️ ARASAKA Speed: ${speedMultiplier === 1 ? 'NORMAL' : speedMultiplier === 1.5 ? 'LANGSAM' : 'SEHR LANGSAM'}`,
        () => {
            const next = speedMultiplier === 1 ? 1.5 : speedMultiplier === 1.5 ? 2 : 1;
            const name = next === 1 ? 'NORMAL' : next === 1.5 ? 'LANGSAM' : 'SEHR LANGSAM';
            GM_setValue('arasaka_speed', next);
            alert(`ARASAKA Geschwindigkeit → ${name}\nSeite neu laden, damit die Änderung aktiv wird.`);
        }
    );

    function injectStyles() {
        if (document.getElementById('arasaka-styles')) return;
        const style = document.createElement('style');
        style.id = 'arasaka-styles';
        style.innerHTML = `
            @keyframes arasaka-fadein {
                from { opacity: 0; transform: translateY(16px); }
                to { opacity: 1; transform: translateY(0); }
            }
            @keyframes arasaka-pulse {
                0%, 100% { box-shadow: 0 0 14px var(--a-col, #00ffcc88); }
                50% { box-shadow: 0 0 28px var(--a-col, #00ffcccc); }
            }
            @keyframes arasaka-blink {
                0%, 100% { opacity: 1; } 50% { opacity: 0; }
            }
            @keyframes arasaka-scanline {
                from { background-position: 0 0; }
                to { background-position: 0 100%; }
            }
            #arasaka-hud {
                position: fixed;
                bottom: 28px;
                right: 28px;
                width: 320px;
                background: rgba(5, 8, 12, 0.96);
                color: #00ffcc;
                border: 1.5px solid #00ffcc;
                border-radius: 6px;
                font-family: 'Courier New', monospace;
                font-size: 13px;
                z-index: 999999;
                pointer-events: none;
                animation: arasaka-fadein 0.3s ease, arasaka-pulse 2.5s ease-in-out infinite;
                --a-col: #00ffcc88;
                overflow: hidden;
                user-select: none;
            }
            #arasaka-hud::after {
                content: '';
                position: absolute;
                inset: 0;
                background: repeating-linear-gradient(
                    0deg,
                    transparent,
                    transparent 3px,
                    rgba(0,255,204,0.02) 3px,
                    rgba(0,255,204,0.02) 4px
                );
                pointer-events: none;
            }
            .a-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 10px 14px 8px;
                border-bottom: 1px solid rgba(0,255,204,0.25);
                font-size: 12px;
                letter-spacing: 2px;
                font-weight: bold;
            }
            .a-timer { font-size: 11px; opacity: 0.7; letter-spacing: 1px; }
            .a-progress-wrap {
                padding: 8px 14px 4px;
            }
            .a-step-label {
                display: flex;
                justify-content: space-between;
                font-size: 10px;
                opacity: 0.6;
                margin-bottom: 4px;
                letter-spacing: 1px;
            }
            .a-bar-bg {
                height: 5px;
                background: rgba(0,255,204,0.1);
                border-radius: 3px;
                overflow: hidden;
            }
            .a-bar-fill {
                height: 100%;
                background: #00ffcc;
                border-radius: 3px;
                transition: width 0.4s ease;
            }
            .a-log {
                padding: 6px 14px 0;
                min-height: 44px;
            }
            .a-log-entry {
                font-size: 11px;
                opacity: 0.55;
                margin-bottom: 2px;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }
            .a-log-entry::before { content: '✓ '; }
            .a-current {
                padding: 8px 14px 6px;
                font-size: 13px;
                border-top: 1px solid rgba(0,255,204,0.15);
                margin-top: 4px;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }
            .a-cursor { animation: arasaka-blink 0.9s infinite; }
            .a-footer {
                padding: 4px 14px 8px;
                font-size: 10px;
                opacity: 0.35;
                letter-spacing: 1px;
            }

            #arasaka-confirm {
                position: fixed;
                inset: 0;
                background: rgba(0,0,0,0.7);
                z-index: 1000000;
                display: flex;
                align-items: center;
                justify-content: center;
                animation: arasaka-fadein 0.2s ease;
                pointer-events: all;
            }
            .a-dialog {
                background: rgba(5, 8, 12, 0.98);
                border: 1.5px solid #ffaa00;
                border-radius: 6px;
                padding: 24px 28px;
                font-family: 'Courier New', monospace;
                color: #ffaa00;
                max-width: 420px;
                width: 90%;
                box-shadow: 0 0 30px rgba(255,170,0,0.4);
            }
            .a-dialog-title {
                font-size: 13px;
                font-weight: bold;
                letter-spacing: 2px;
                margin-bottom: 12px;
            }
            .a-dialog-msg {
                font-size: 12px;
                line-height: 1.6;
                opacity: 0.85;
                margin-bottom: 20px;
                white-space: pre-wrap;
            }
            .a-dialog-btns { display: flex; gap: 10px; justify-content: flex-end; }
            .a-btn {
                padding: 7px 18px;
                font-family: 'Courier New', monospace;
                font-size: 12px;
                font-weight: bold;
                letter-spacing: 1px;
                border-radius: 3px;
                cursor: pointer;
                border: 1.5px solid currentColor;
                background: transparent;
                transition: background 0.15s;
            }
            .a-btn-ok { color: #00ffcc; }
            .a-btn-ok:hover { background: rgba(0,255,204,0.15); }
            .a-btn-cancel { color: #ff4444; }
            .a-btn-cancel:hover { background: rgba(255,68,68,0.15); }
        `;
        document.head.appendChild(style);
    }

    function createHUD() {
        if (hudElement) return;
        injectStyles();
        hudElement = document.createElement('div');
        hudElement.id = 'arasaka-hud';
        hudElement.innerHTML = buildHUDHTML('Initialisiere...', '#00ffcc', currentStep);
        document.body.appendChild(hudElement);

        startTime = Date.now();
        timerInterval = setInterval(() => {
            const elapsed = Math.floor((Date.now() - startTime) / 1000);
            const m = String(Math.floor(elapsed / 60)).padStart(2, '0');
            const s = String(elapsed % 60).padStart(2, '0');
            const el = document.querySelector('.a-timer');
            if (el) el.textContent = `${m}:${s}`;
        }, 1000);
    }

    function buildHUDHTML(action, color, step) {
        const displayStep = Math.min(step, TOTAL_STEPS);
        const pct = Math.min(100, Math.round((displayStep / TOTAL_STEPS) * 100));
        const logHTML = hudLog.slice(-3).map(t =>
            `<div class="a-log-entry">${t}</div>`
        ).join('');

        return `
            <div class="a-header" style="color:${color}; border-color: ${color}40;">
                <span>ARASAKA SYS //</span>
                <span class="a-timer">00:00</span>
            </div>
            <div class="a-progress-wrap">
                <div class="a-step-label" style="color:${color}">
                    <span>FORTSCHRITT</span>
                    <span>SCHRITT ${displayStep}/${TOTAL_STEPS}</span>
                </div>
                <div class="a-bar-bg">
                    <div class="a-bar-fill" style="width:${pct}%; background:${color};"></div>
                </div>
            </div>
            <div class="a-log" style="color:${color}">${logHTML}</div>
            <div class="a-current" style="color:${color}">
                ► ${action}<span class="a-cursor">_</span>
            </div>
            <div class="a-footer">[ALT+Y] START  ·  [ESC] ABBRUCH</div>
        `;
    }

    function updateHUD(text, color = '#00ffcc', advanceStep = false) {
        if (advanceStep && hudLog[hudLog.length - 1] !== text) {
            const prev = document.querySelector('.a-current');
            if (prev) {
                const prevText = prev.textContent.replace('► ', '').replace('_', '').trim();
                if (prevText && prevText !== text) hudLog.push(prevText);
            }
            currentStep++;
        }

        if (!hudElement) createHUD();

        hudElement.style.borderColor = color;
        hudElement.style.setProperty('--a-col', color + '88');

        const elapsed = document.querySelector('.a-timer')?.textContent || '00:00';

        hudElement.innerHTML = buildHUDHTML(text, color, currentStep);
        hudElement.style.borderColor = color;

        const timerEl = hudElement.querySelector('.a-timer');
        if (timerEl) timerEl.textContent = elapsed;
    }

    function removeHUD() {
        clearInterval(timerInterval);
        timerInterval = null;
        if (hudElement) {
            hudElement.style.animation = 'none';
            hudElement.style.transition = 'opacity 0.4s ease';
            hudElement.style.opacity = '0';
            setTimeout(() => { hudElement?.remove(); hudElement = null; }, 400);
        }
        hudLog = [];
        currentStep = 0;
        startTime = null;
    }

    function arasakaConfirm(message) {
        return new Promise((resolve) => {
            injectStyles();
            const overlay = document.createElement('div');
            overlay.id = 'arasaka-confirm';
            overlay.innerHTML = `
                <div class="a-dialog">
                    <div class="a-dialog-title">⚠ ARASAKA WARNUNG</div>
                    <div class="a-dialog-msg">${message}</div>
                    <div class="a-dialog-btns">
                        <button class="a-btn a-btn-cancel" id="a-btn-no">ABBRECHEN</button>
                        <button class="a-btn a-btn-ok"    id="a-btn-yes">RETTUNG STARTEN</button>
                    </div>
                </div>
            `;
            document.body.appendChild(overlay);

            document.getElementById('a-btn-yes').onclick = () => { overlay.remove(); resolve(true); };
            document.getElementById('a-btn-no').onclick  = () => { overlay.remove(); resolve(false); };
        });
    }

    function playDing(type = 'success') {
        try {
            const ctx = new (window.AudioContext || window.webkitAudioContext)();
            const osc  = ctx.createOscillator();
            const gain = ctx.createGain();
            osc.connect(gain);
            gain.connect(ctx.destination);
            if (type === 'success') {
                osc.type = 'sine';
                osc.frequency.setValueAtTime(880, ctx.currentTime);
                osc.frequency.exponentialRampToValueAtTime(1760, ctx.currentTime + 0.1);
                gain.gain.setValueAtTime(0.2, ctx.currentTime);
                gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.5);
                osc.start(); osc.stop(ctx.currentTime + 0.5);
            } else {
                osc.type = 'sawtooth';
                osc.frequency.setValueAtTime(150, ctx.currentTime);
                gain.gain.setValueAtTime(0.3, ctx.currentTime);
                gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.4);
                osc.start(); osc.stop(ctx.currentTime + 0.4);
            }
        } catch (e) {}
    }

    async function sleep(ms) {
        let t = 0;
        const target = ms * speedMultiplier;
        while (t < target) {
            if (abortMission) throw new Error('Abort');
            await new Promise(r => setTimeout(r, 100));
            t += 100;
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
        const target = timeout * speedMultiplier;
        while (passed < target) {
            if (abortMission) throw new Error('Abort');
            const els = getDeepestElementsByText(text);
            if (els.length > 0) return pickLast ? els[els.length - 1] : els[0];
            await sleep(500);
            passed += 500;
        }
        return null;
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

    function showArasakaRegalHUD(regalText) {
        let existing = document.getElementById('arasaka-regal-hud');
        if (existing) existing.remove();

        const hud = document.createElement('div');
        hud.id = 'arasaka-regal-hud';
        hud.style.position = 'fixed';
        hud.style.top = '50%';
        hud.style.left = '50%';
        hud.style.transform = 'translate(-50%, -50%)';
        hud.style.backgroundColor = 'rgba(10, 10, 10, 0.95)';
        hud.style.border = '4px solid #00FF00';
        hud.style.boxShadow = '0 0 40px #00FF00, inset 0 0 20px #00FF00';
        hud.style.padding = '60px 100px';
        hud.style.borderRadius = '20px';
        hud.style.zIndex = '9999999';
        hud.style.color = '#00FF00';
        hud.style.fontFamily = 'monospace';
        hud.style.textAlign = 'center';
        hud.style.cursor = 'pointer';

        hud.innerHTML = `
            <div style="font-size: 30px; margin-bottom: 20px; text-shadow: 0 0 10px #00FF00;">TEILE STANDEN IN:</div>
            <div style="font-size: 80px; font-weight: bold; color: #FFF; text-shadow: 0 0 20px #FFF, 0 0 40px #00FF00;">${regalText}</div>
            <div style="font-size: 16px; margin-top: 30px; opacity: 0.7;">(Klick oder ESC zum Schließen)</div>
        `;

        document.body.appendChild(hud);

        const closeHUD = (e) => {
            if (e.type === 'keydown' && e.key !== 'Escape') return;
            hud.remove();
            document.removeEventListener('keydown', closeHUD);
            document.removeEventListener('click', closeHUD);
        };

        setTimeout(() => {
            document.addEventListener('click', closeHUD);
            document.addEventListener('keydown', closeHUD);
        }, 200);
    }

    function sendGhostPing(stockId) {
        GM_xmlhttpRequest({
            method: 'GET',
            url: `${googleWebAppUrl}?stock=${stockId}`,
            onload:  r => {
                let match = r.responseText.match(/OLD_REGAL:(.*?)(?: \||$)/);
                if (match && match[1]) {
                    sessionStorage.setItem('arasaka_old_regal', match[1].trim());
                }
            },
            onerror: e => {}
        });
    }

    async function sucheUndOeffnePdf() {
        updateHUD('Warte auf Tabellen-Aufbau...', '#00ffcc', true);
        await sleep(4000);

        const textWerkstatt = await waitForElement('Werkstattauftrag', 15000);
        if (!textWerkstatt) return;

        let parent = textWerkstatt.parentElement;
        let pdfClicked = false;

        updateHUD('Scanne nach PDF...', '#ffff00', true);

        for (let i = 0; i < 8; i++) {
            if (!parent) break;
            for (const el of parent.querySelectorAll('*')) {
                if (!el.textContent?.includes('.pdf')) continue;
                const childHas = Array.from(el.children).some(c => c.textContent?.includes('.pdf'));
                if (childHas) continue;
                el.scrollIntoView({ behavior: 'smooth', block: 'center' });
                await sleep(800);
                forceClick(el);
                pdfClicked = true;
                break;
            }
            if (pdfClicked) break;
            parent = parent.parentElement;
        }

        if (pdfClicked) {
            updateHUD('PDF GEÖFFNET ─ Bereit für STRG+P', '#00ff00', true);
            playDing('success');
           
            const oldRegal = sessionStorage.getItem('arasaka_old_regal');
            if (oldRegal) {
                showArasakaRegalHUD(oldRegal);
                sessionStorage.removeItem('arasaka_old_regal');
            }

            setTimeout(removeHUD, 6000);
        }
    }

    async function startMacro() {
        try {
            hudLog = []; currentStep = 0;
            createHUD();
            updateHUD('Initialisiere...');

            const titleMatch = document.title.match(/[A-Z]{2}\d{5}/i);
            if (titleMatch) {
                const stockId = titleMatch[0].toUpperCase();
                updateHUD(`Stock-ID ${stockId} erkannt`, '#ffff00', true);
                sendGhostPing(stockId);
                await sleep(1500);
            }

            updateHUD('Öffne Edit damages...', '#00ffcc', true);
            const btnEdit = await waitForElement('Edit damages and services');
            if (!btnEdit) { updateHUD('Edit-Button nicht gefunden!', '#ff4444'); return; }
            forceClick(btnEdit);

            updateHUD('Navigiere zu Spare Parts...', '#00ffcc', true);
            await sleep(2000);
            const btnSpareParts = await waitForElement('Spare parts');
            if (!btnSpareParts) { updateHUD('Spare Parts nicht gefunden!', '#ff4444'); return; }
            forceClick(btnSpareParts);

            updateHUD('Wähle alle Positionen aus...', '#00ffcc', true);
            await sleep(1000);
            const btnSelectAll = await waitForElement('Select all');
            if (!btnSelectAll) { updateHUD('Select All nicht gefunden!', '#ff4444'); return; }
            forceClick(btnSelectAll);

            updateHUD('Setze Dropdowns auf Handed out...', '#ffff00', true);
            await sleep(1000);
            forceAllHandedOut();

            updateHUD('Übernehme Änderungen...', '#00ffcc', true);
            await sleep(1000);
            const btnApply = await waitForElement('Apply');
            if (btnApply) forceClick(btnApply);

            updateHUD('Verifiziere Handed-Out Status...', '#ffff00', true);
            let verified = false;
            for (let i = 0; i < 6; i++) {
                await sleep(1000);
                if (checkAllHandedOut()) { verified = true; break; }
            }

            if (!verified) {
                playDing('error');
                const retry = await arasakaConfirm(
                    'Nicht alle Positionen konnten auf "Handed out"\numgestellt werden.\n\nSoll ARASAKA einen Rettungsversuch starten?'
                );

                if (retry) {
                    updateHUD('Rettungsprotokoll aktiv...', '#ffaa00', true);
                    forceAllHandedOut();
                    await sleep(3000);
                    if (!checkAllHandedOut()) {
                        updateHUD('RETTUNG FEHLGESCHLAGEN', '#ff4444', true);
                        playDing('error');
                        setTimeout(removeHUD, 5000);
                        return;
                    }
                    updateHUD('Rettung erfolgreich!', '#00ff00', true);
                    await sleep(1000);
                } else {
                    updateHUD('MANUELLE ÜBERNAHME', '#ff4444');
                    setTimeout(removeHUD, 5000);
                    return;
                }
            }

            updateHUD('Übermittle Daten...', '#00ffcc', true);
            const btnSubmit = await waitForElement('Submit');
            if (btnSubmit) forceClick(btnSubmit);

            updateHUD('Bestätige Freigabe...', '#00ffcc', true);
            await sleep(1000);
            const btnConfirm = await waitForElement('Confirm');
            if (btnConfirm) forceClick(btnConfirm);

            updateHUD('Schließe Dialog...', '#00ffcc', true);
            await sleep(2000);
            const btnClose = await waitForElement('Close', 10000, true);
            if (btnClose) forceClick(btnClose);

            updateHUD('Rückkehr zum Auftrag...', '#00ffcc', true);
            await sleep(1500);
            const btnBack = await waitForElement('Back to refurbishment detail', 15000);
            if (btnBack) {
                sessionStorage.setItem('hole_pdf_nach_reload', 'true');
                forceClick(btnBack);
            }

        } catch (e) {
            if (e.message === 'Abort') {
                updateHUD('ABBRUCH DURCH USER', '#ff4444');
            } else {
                updateHUD(`SYSTEMFEHLER: ${e.message}`, '#ff4444');
            }
            playDing('error');
            setTimeout(removeHUD, 4000);
        }
    }

    if (sessionStorage.getItem('hole_pdf_nach_reload') === 'true') {
        sessionStorage.removeItem('hole_pdf_nach_reload');
        isLocked = true;
        currentStep = 11;
        (async () => {
            try {
                await sucheUndOeffnePdf();
            } catch (e) {
                const msg = e.message === 'Abort' ? 'ABBRUCH DURCH USER' : 'SYSTEMFEHLER';
                updateHUD(msg, '#ff4444');
                playDing('error');
                setTimeout(removeHUD, 3000);
            } finally {
                isLocked = false;
            }
        })();
    }

    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') {
            abortMission = true;
            isLocked = false;
        }
        if (event.altKey && event.key.toLowerCase() === 'y') {
            event.preventDefault();
            if (isLocked) return;
            isLocked = true;
            abortMission = false;
            startMacro().finally(() => { isLocked = false; });
        }
    });

})();
