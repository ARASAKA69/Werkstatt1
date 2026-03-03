// ==UserScript==
// @name         Auftrags-Automatisierung
// @namespace    http://tampermonkey.net/
// @version      1.1
// @description  Automatisiert den Freigabe- und Druckprozess mit dynamischem Dropdown-Check
// @author       ARASAKA
// @match        *://*/*
// @updateURL    https://raw.githubusercontent.com/ARASAKA69/Werkstatt1/refs/heads/main/arasaka-automatisierung.user.js
// @downloadURL  https://raw.githubusercontent.com/ARASAKA69/Werkstatt1/refs/heads/main/arasaka-automatisierung.user.js
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

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
        while (timePassed < timeout) {
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
        el.click();
        if (el.parentElement) {
            setTimeout(() => el.parentElement.click(), 50);
        }
    }

    async function sucheUndOeffnePdf() {
        await sleep(4000);

        let textWerkstatt = await waitForElement('Werkstattauftrag', 15000);
        if (textWerkstatt) {
            let parent = textWerkstatt.parentElement;
            let pdfClicked = false;

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
        }
    }

    async function startMacro() {
        let btnEdit = await waitForElement('Edit damages and services');
        if (btnEdit) { forceClick(btnEdit); } else return;

        await sleep(2000);
        let btnSpareParts = await waitForElement('Spare parts');
        if (btnSpareParts) { forceClick(btnSpareParts); } else return;

        await sleep(1000);
        let btnSelectAll = await waitForElement('Select all');
        if (btnSelectAll) { forceClick(btnSelectAll); } else return;

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

        await sleep(1000);
        let btnApply = await waitForElement('Apply');
        if (btnApply) { forceClick(btnApply); }

        let maxChecks = 3;
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
            alert("⚠️ HEE DU HONK ! Achtung: Es konnten nicht alle Positionen auf 'Handed out' umgestellt werden. Das Skript wurde gestoppt. Bitte prüfe die Liste und mach manuell weiter (Und Machs gescheit!).");
            return;
        }

        let btnSubmit = await waitForElement('Submit');
        if (btnSubmit) { forceClick(btnSubmit); }

        await sleep(1000);
        let btnConfirm = await waitForElement('Confirm');
        if (btnConfirm) { forceClick(btnConfirm); }

        await sleep(2000);
        let btnClose = await waitForElement('Close', 10000, true);
        if (btnClose) { forceClick(btnClose); }

        await sleep(1500);
        let btnBack = await waitForElement('Back to refurbishment detail', 15000);
        if (btnBack) {
            sessionStorage.setItem('hole_pdf_nach_reload', 'true');
            forceClick(btnBack);
        }
    }

    if (sessionStorage.getItem('hole_pdf_nach_reload') === 'true') {
        sessionStorage.removeItem('hole_pdf_nach_reload');
        sucheUndOeffnePdf();
    }

    document.addEventListener('keydown', function(event) {
        if (event.altKey && event.key.toLowerCase() === 'y') {
            event.preventDefault();
            startMacro();
        }
    });

})();
