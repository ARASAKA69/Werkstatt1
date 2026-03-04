// ==UserScript==
// @name         Carol-Automation
// @namespace    http://tampermonkey.net/
// @version      1.8
// @description  ARASAKA Premium (HUD, Audio-Ping, Not-Aus, Laser-Targeting, Admin-Menu)
// @author       ARASAKA
// @match        *://*/*
// @updateURL    https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/arasaka-automation.user.js
// @downloadURL  https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/arasaka-automation.user.js
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_registerMenuCommand
// ==/UserScript==

(function() {
    'use strict';

    let abortMission = false;
    let hudElement = null;
    const speedMultiplier = GM_getValue('arasaka_speed', 1);

    GM_registerMenuCommand(`⚙️ ARASAKA Speed: ${speedMultiplier === 1 ? 'NORMAL' : (speedMultiplier === 1.5 ? 'LANGSAM' : 'SEHR LANGSAM')}`, () => {
        let next = speedMultiplier === 1 ? 1.5 : (speedMultiplier === 1.5 ? 2 : 1);
        let modeName = next === 1 ? 'NORMAL' : (next === 1.5 ? 'LANGSAM' : 'SEHR LANGSAM');
        GM_setValue('arasaka_speed', next);
        alert(`ARASAKA Geschwindigkeit geändert auf: ${modeName}\nBitte die Seite neu laden, damit die Änderung aktiv wird.`);
    });

    function createHUD() {
        if (!hudElement) {
            hudElement = document.createElement('div');
            hudElement.style.position = 'fixed';
            hudElement.style.bottom = '20px';
            hudElement.style.right = '20px';
            hudElement.style.backgroundColor = 'rgba(10, 10, 10, 0.9)';
            hudElement.style.color = '#00ffcc';
            hudElement.style.border = '1px solid #00ffcc';
            hudElement.style.padding = '12px 20px';
            hudElement.style.fontFamily = 'monospace';
            hudElement.style.fontSize = '14px';
            hudElement.style.zIndex = '999999';
            hudElement.style.pointerEvents = 'none';
            hudElement.style.boxShadow = '0 0 15px rgba(0, 255, 204, 0.4)';
            hudElement.style.borderRadius = '3px';
            document.body.appendChild(hudElement);
        }
    }

    function updateHUD(text, color = '#00ffcc') {
        createHUD();
        hudElement.style.color = color;
        hudElement.style.border = `1px solid ${color}`;
        hudElement.style.boxShadow = `0 0 15px ${color}80`;
        hudElement.innerHTML = `<strong style="letter-spacing: 2px;">ARASAKA SYS //</strong><br><br>${text}`;
    }

    function removeHUD() {
        if (hudElement) {
            hudElement.remove();
            hudElement = null;
        }
    }

    function playDing(type = 'success') {
        try {
            const ctx = new (window.AudioContext || window.webkitAudioContext)();
            const osc = ctx.createOscillator();
            const gain = ctx.createGain();
            osc.connect(gain);
            gain.connect(ctx.destination);
            
            if (type === 'success') {
                osc.type = 'sine';
                osc.frequency.setValueAtTime(880, ctx.currentTime);
                osc.frequency.exponentialRampToValueAtTime(1760, ctx.currentTime + 0.1);
                gain.gain.setValueAtTime(0.2, ctx.currentTime);
                gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.5);
                osc.start();
                osc.stop(ctx.currentTime + 0.5);
            } else {
                osc.type = 'sawtooth';
                osc.frequency.setValueAtTime(150, ctx.currentTime);
                gain.gain.setValueAtTime(0.3, ctx.currentTime);
                gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.4);
                osc.start();
                osc.stop(ctx.currentTime + 0.4);
            }
        } catch(e) {}
    }

    async function sleep(ms) {
        let t = 0;
        let targetMs = ms * speedMultiplier;
        while (t < targetMs) {
            if (abortMission) throw new Error("Abort");
            await new Promise(r => setTimeout(r, 100));
            t += 100;
        }
    }

    function getDeepestElementsByText(text) {
        const elements = document.querySelectorAll('*');
        const results = [];
        for (let el of elements) {
            if (['SCRIPT', 'STYLE', 'HTML', 'HEAD', 'BODY'].includes(el.tagName)) continue;
            if (el.textContent && el.textContent.includes(text)) {
                let childHasText = Array.from(el.children).some(child =>
                    child.textContent && child.textContent.includes(text)
                );
                if (!childHasText) {
                    results.push(el);
                }
            }
        }
        return results;
    }

    async function waitForElement(text, timeout = 10000, pickLast = false) {
        let timePassed = 0;
        let targetTimeout = timeout * speedMultiplier;
        while (timePassed < targetTimeout) {
            if (abortMission) throw new Error("Abort");
            let els = getDeepestElementsByText(text);
            if (els.length > 0) {
                return pickLast ? els[els.length - 1] : els[0];
            }
            await sleep(500);
            timePassed += 500;
        }
        return null;
    }

    function forceClick(el) {
        if (!el) return;
        try {
            const oOutline = el.style.outline;
            const oShadow = el.style.boxShadow;
            const oTrans = el.style.transition;
            el.style.transition = 'all 0.1s';
            el.style.outline = '2px solid #ff0000';
            el.style.boxShadow = '0 0 15px #ff0000';
            setTimeout(() => {
                if(el) {
                    el.style.outline = oOutline;
                    el.style.boxShadow = oShadow;
                    el.style.transition = oTrans;
                }
            }, 400);
        } catch(e) {}

        el.click();
        if (el.parentElement) {
            setTimeout(() => el.parentElement.click(), 50);
        }
    }

    async function sucheUndOeffnePdf() {
        updateHUD("Warte auf Tabellen-Aufbau...");
        await sleep(4000);

        updateHUD("Suche Werkstattauftrag...");
        let textWerkstatt = await waitForElement('Werkstattauftrag', 15000);
        if (textWerkstatt) {
            let parent = textWerkstatt.parentElement;
            let pdfClicked = false;

            updateHUD("Scanne nach PDF...", "#ffff00");
            for (let i = 0; i < 8; i++) {
                if (!parent) break;

                const allElements = parent.querySelectorAll('*');
                for (let el of allElements) {
                    if (el.textContent && el.textContent.includes('.pdf')) {
                        let childHasText = Array.from(el.children).some(child =>
                            child.textContent && child.textContent.includes('.pdf')
                        );
                        if (!childHasText) {
                            forceClick(el);
                            pdfClicked = true;
                            break;
                        }
                    }
                }
                if (pdfClicked) break;
                parent = parent.parentElement;
            }

            if (pdfClicked) {
                updateHUD("PDF GEÖFFNET.<br>Bereit für STRG+P", "#00ff00");
                playDing('success');
                setTimeout(removeHUD, 6000);
            }
        }
    }

    async function startMacro() {
        try {
            updateHUD("Initialisiere...");
            
            let btnEdit = await waitForElement('Edit damages and services');
            if (btnEdit) { forceClick(btnEdit); } else return;

            updateHUD("Navigiere zu Spare Parts...");
            await sleep(2000);
            let btnSpareParts = await waitForElement('Spare parts');
            if (btnSpareParts) { forceClick(btnSpareParts); } else return;

            updateHUD("Wähle alles aus...");
            await sleep(1000);
            let btnSelectAll = await waitForElement('Select all');
            if (btnSelectAll) { forceClick(btnSelectAll); } else return;

            updateHUD("Passe Dropdowns an...");
            await sleep(1000);
            let selects = document.querySelectorAll('select');
            let foundDropdown = false;
            for (let select of selects) {
                for (let option of select.options) {
                    if (option.text.includes('Handed out')) {
                        select.value = option.value;
                        select.dispatchEvent(new Event('change', { bubbles: true }));
                        foundDropdown = true;
                        break;
                    }
                }
                if (foundDropdown) break;
            }

            updateHUD("Wende Änderungen an...");
            await sleep(1000);
            let btnApply = await waitForElement('Apply');
            if (btnApply) { forceClick(btnApply); }

            updateHUD("Prüfe Handed-Out...", "#ffff00");
            let maxChecks = 6;
            let checksDone = 0;
            let allUpdated = false;

            while (checksDone < maxChecks) {
                await sleep(1000);
                checksDone++;

                let allSelects = document.querySelectorAll('select');
                let unfertigeGefunden = false;
                let relevanteDropdownsGefunden = 0;

                for (let select of allSelects) {
                    let hasHandedOutOption = Array.from(select.options).some(opt => opt.text.includes('Handed out'));

                    if (hasHandedOutOption) {
                        relevanteDropdownsGefunden++;
                        let currentText = select.options[select.selectedIndex].text;

                        if (!currentText.includes('Handed out')) {
                            unfertigeGefunden = true;
                            break;
                        }
                    }
                }

                if (relevanteDropdownsGefunden > 0 && !unfertigeGefunden) {
                    allUpdated = true;
                    break;
                }
            }

            if (!allUpdated) {
                updateHUD("FEHLER: SYNC FEHLGESCHLAGEN!", "#ff0000");
                playDing('error');
                setTimeout(removeHUD, 5000);
                alert("⚠️ HEE DU HONK ! Achtung: Es konnten nicht alle Positionen auf 'Handed out' umgestellt werden. Das Skript wurde gestoppt. Bitte prüfe die Liste und mach manuell weiter (Und Machs gescheit!).");
                return;
            }

            updateHUD("Übermittle Daten...");
            let btnSubmit = await waitForElement('Submit');
            if (btnSubmit) { forceClick(btnSubmit); }

            updateHUD("Bestätige Freigabe...");
            await sleep(1000);
            let btnConfirm = await waitForElement('Confirm');
            if (btnConfirm) { forceClick(btnConfirm); }

            updateHUD("Schließe Fenster...");
            await sleep(2000);
            let btnClose = await waitForElement('Close', 10000, true);
            if (btnClose) { forceClick(btnClose); }

            updateHUD("Rückkehr zum Auftrag...");
            await sleep(1500);
            let btnBack = await waitForElement('Back to refurbishment detail', 15000);
            if (btnBack) {
                sessionStorage.setItem('hole_pdf_nach_reload', 'true');
                forceClick(btnBack);
            }
            
        } catch (e) {
            if (e.message === "Abort") {
                updateHUD("ABBRUCH DURCH USER", "#ff0000");
                playDing('error');
                setTimeout(removeHUD, 3000);
            } else {
                updateHUD("SYSTEMFEHLER", "#ff0000");
                playDing('error');
                setTimeout(removeHUD, 3000);
                console.error("ARASAKA ERROR:", e);
            }
        }
    }

    if (sessionStorage.getItem('hole_pdf_nach_reload') === 'true') {
        sessionStorage.removeItem('hole_pdf_nach_reload');
        abortMission = false;
        (async () => {
            try {
                await sucheUndOeffnePdf();
            } catch (e) {
                if (e.message === "Abort") {
                    updateHUD("ABBRUCH DURCH USER", "#ff0000");
                    playDing('error');
                    setTimeout(removeHUD, 3000);
                } else {
                    updateHUD("SYSTEMFEHLER", "#ff0000");
                    playDing('error');
                    setTimeout(removeHUD, 3000);
                    console.error("ARASAKA ERROR:", e);
                }
            }
        })();
    }

    document.addEventListener('keydown', function(event) {
        if (event.key === 'Escape') {
            abortMission = true;
        }
        if (event.altKey && event.key.toLowerCase() === 'y') {
            event.preventDefault();
            abortMission = false;
            startMacro();
        }
    });

})();
