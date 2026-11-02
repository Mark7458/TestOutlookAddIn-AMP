/*
 * AMP - SNOW Assistant: area AMP dell'add-in Outlook — VERSIONE DI PRODUZIONE.
 *
 * Contiene solo la scheda In coda (mail della casella condivisa):
 *   - Attive / Ignorate;
 *   - scheda aperta: testo della mail e selettore Prendi in carico · Inoltra · Ignora;
 *     Prendi in carico c'è ma non è attiva (fino a nuovo ordine); Inoltra (gruppi ES, PEL, PO) e Ignora sì.
 * Niente Ticket, vista Mail, Impostazioni ed etichette della casella: sono solo nel pannello di sviluppo (amp.js).
 * Le parti in comune sono copiate da amp.js: una modifica fatta lì va riportata qui.
 *
 * Collegamento con l'add-in di jak (nessun suo file viene modificato):
 *   - espone window.AmpModule { initialize, show }; "Apri area AMP" chiama show();
 *   - la freccia in alto torna alla schermata principale (TestPluginAPI.showMainView);
 *   - l'utente è quello di Outlook (TestPluginAPI.getCurrentMessageContext, altrimenti Office).
 *
 * Servizi (window.AmpBackend, da collegare alla Function App; vedi amp-servizi.md):
 *   listQueue(), listIgnored(), discardMail(conversationId, utente)       → necessari
 *   getMailBody(messageId), forwardMail(conversationId, gruppo, utente), subscribe(listener) → facoltativi
 */
(function () {
    'use strict';

    // ======================================================================
    // 0. ASPETTO: TEMA E TESTI
    // ======================================================================

    // Tutti i valori di aspetto: il CSS li usa come variabili --ao-* (vedi themeCssVariables()).
    const THEME = {
        colors: {
            accent: '#0078d4', // pulsanti principali, scheda attiva, voci selezionate
            accentDark: '#106ebe', // colore principale al passaggio del mouse
            textAccent: '#106ebe', // tutte le scritte blu (numeri dei ticket, link, schede attive…)
            accentBorder: '#c7e0f4',
            accentBg: '#eff6fc',
            white: '#fff', // testo su sfondo blu (pulsanti, notifiche)
            surface: '#fff', // sfondo di card, campi, schede, barra di sistema
            text: '#242424',
            textMuted: '#605e5c',
            textSubtle: '#8a8886',
            bg: '#f5f5f5', // sfondo dell'area (come il riquadro dell'add-in)
            bgSoft: '#faf9f8',
            hover: '#faf9f8', // sfondo delle schede al passaggio del mouse
            lineLight: '#edebe9',
            line: '#e1dfdd',
            lineStrong: '#c8c6c4',
            success: '#107c10',
            successBg: '#dff6dd',
            warning: '#8a5a00',
            warningBg: '#fff4ce',
            danger: '#a4262c',
            dangerBg: '#fde7e9',
            forward: '#7c5bbf', // Inoltra (In coda): pulsante ed esito
            forwardText: '#6b47b0',
            forwardBg: '#f1ecf9'
        },
        // Tema scuro: segue Outlook (prefers-color-scheme) oppure la scelta in Impostazioni → Profilo.
        // Sfondo, riquadri, bordi e testi come l'add-in di jak in tema scuro.
        darkColors: {
            accentDark: '#2b88d8',
            textAccent: '#6cb8f6',
            accentBorder: '#3a6a8f',
            accentBg: '#19344a',
            surface: '#2d2c2b',
            text: '#f3f2f1',
            textMuted: '#c8c6c4',
            textSubtle: '#a19f9d',
            bg: '#201f1e',
            bgSoft: '#292827',
            hover: '#3b3a39', // al passaggio si schiarisce, come i pulsanti dell'add-in di jak
            lineLight: '#3b3a39',
            line: '#484644',
            lineStrong: '#605e5c',
            success: '#6ccb5f',
            successBg: '#1f3a1d',
            warning: '#f2c661',
            warningBg: '#3d3214',
            danger: '#ffb3b8',
            dangerBg: '#442726',
            forwardText: '#c3b2ff',
            forwardBg: '#2f2a45'
        },
        font: { family: '"Segoe UI", system-ui, -apple-system, BlinkMacSystemFont, sans-serif' },
        fontSize: { xs: '10px', sm: '11px', md: '12px', base: '13px', lg: '14px', xl: '16px' },
        fontWeight: { normal: '400', medium: '600', bold: '700' },
        radius: { base: '6px', pill: '999px', badge: '5px' }, // badge: notifiche, quadrato con gli angoli smussati
        shadow: {
            card: '0 1px 2px rgba(0,0,0,.06)',
            focus: '0 0 0 2px rgba(0,120,212,.25)'
        },
        state: { disabledOpacity: '.45' },
        layout: { controlHeight: '32px' }
    };

    const TEXTS = {
        back: 'Torna all’add-in',
        system: {
            servicesOk: 'Servizi collegati',
            servicesMissing: 'Servizi non ancora collegati',
            serviceError: message => `Servizi non raggiungibili: ${message}`
        },
        tabs: { queue: 'In coda' },
        common: {
            loading: 'Caricamento…',
            confirm: 'Confermi?',
            close: 'Chiudi',
            none: '—',
            servicesMissing: 'I servizi non sono ancora collegati: qui compariranno i dati della casella condivisa e di ServiceNow.'
        },
        queue: {
            take: 'Prendi in carico',
            forward: 'Inoltra',
            ignore: 'Ignora',
            ignoredDone: 'Ignorata',
            forwardWorking: 'Inoltro…',
            forwardDone: group => `Inoltrata a ${group}`,
            empty: 'Nessuna richiesta in coda.',
            active: 'Attive',
            ignoredTab: 'Ignorate',
            ignoredEmpty: 'Nessun thread ignorato.',
            ignoredBy: (who, when) => `Ignorata da ${who} · ${when}`
        }
    };

    // Versione del pannello (attributo data-build dell'area): serve a capire se il browser usa l'ultima.
    const BUILD = '2026-11-02 15:25';

    // Gruppi a cui si inoltra (il tag va davanti all'oggetto della mail: [ES], [PEL], [PO]).
    const FORWARD_GROUPS = [
        { id: 'ES', name: 'Email Security' },
        { id: 'PEL', name: 'Amministratori PEL' },
        { id: 'PO', name: 'Post Office' }
    ];

    // ======================================================================
    // 1. SERVIZI E UTENTE
    // ======================================================================

    const REQUIRED_SERVICES = ['listQueue', 'listIgnored', 'discardMail'];
    class ServicesMissingError extends Error {}
    const backend = () => window.AmpBackend;
    const call = (name, ...args) => {
        if (typeof backend()?.[name] !== 'function') return Promise.reject(new ServicesMissingError(name));
        return Promise.resolve(backend()[name](...args));
    };
    const Services = {
        available: () => REQUIRED_SERVICES.every(name => typeof backend()?.[name] === 'function'),
        listQueue: () => call('listQueue'),
        listIgnored: () => call('listIgnored'),
        discardMail: (conversationId, user) => call('discardMail', conversationId, user),
        forwardMail: (conversationId, group, user) => call('forwardMail', conversationId, group, user),
        getMailBody: async messageId => (typeof backend()?.getMailBody === 'function' ? backend().getMailBody(messageId) : null),
        subscribe: listener => backend()?.subscribe?.(listener) || (() => {})
    };
    // Utente del pannello: il nome dell'utente di Outlook.
    function currentUser() {
        const name = window.TestPluginAPI?.getCurrentMessageContext?.()?.mailbox?.userDisplayName
            || window.Office?.context?.mailbox?.userProfile?.displayName || '';
        return { displayValue: name, value: '' };
    }

    const pad = value => String(value).padStart(2, '0');
    function parseDate(value) {
        if (!value) return null;
        if (value instanceof Date) return value;
        const text = String(value);
        // Formato di ServiceNow "AAAA-MM-GG hh:mm:ss" (ora locale) oppure ISO.
        const date = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}(:\d{2})?$/.test(text) ? new Date(text.replace(' ', 'T')) : new Date(text);
        return Number.isNaN(date.getTime()) ? null : date;
    }
    // "14:05" oggi, "ieri 14:05", altrimenti "27/10 14:05".
    function formatWhen(value) {
        const date = parseDate(value);
        if (!date) return TEXTS.common.none;
        const time = `${pad(date.getHours())}:${pad(date.getMinutes())}`;
        const today = new Date();
        const dayDiff = Math.round((new Date(today.toDateString()) - new Date(date.toDateString())) / 86400000);
        if (dayDiff === 0) return time;
        if (dayDiff === 1) return `ieri ${time}`;
        if (dayDiff === -1) return `domani ${time}`;
        return `${pad(date.getDate())}/${pad(date.getMonth() + 1)} ${time}`;
    }

    // ======================================================================
    // 2. INTERFACCIA
    // ======================================================================

    const CHEVRON_UP_ICON = '<svg viewBox="0 0 20 20" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l5-5 5 5"/></svg>';
    function svgIcon(markup) {
        const icon = h('span', { class: 'ao-svg-icon', 'aria-hidden': 'true' });
        icon.innerHTML = markup;
        return icon;
    }

    const state = {
        queue: [],
        ignored: [], // thread ignorati (archivio)
        loaded: false,
        loading: false,
        serviceError: '',
        queueTab: 'active', // active (da smistare) | ignored (archivio)
        barOffset: 0, // di quanti px la barra secondaria è scivolata via (segue lo scorrimento)
        queueExpanded: null,
        queueChoice: {}, // scheda aperta: azione scelta nel selettore e gruppo dell'inoltro
        liftFrom: null, // scheda che si sta aprendo: disegnata non sollevata, poi si solleva
        previewScroll: null, // scheda aperta del tutto: il testo della mail può scorrere
        mailBodies: {}, // testo delle mail (per messageId)
        holdRender: false, // true durante un'animazione: il pannello non si ridisegna
        busy: false
    };
    let root = null;
    let statusTimer = null;

    // Crea un elemento: h('button', { class: 'x', onclick: fn, disabled: true }, 'testo', figlio, ...)
    function h(tag, props = {}, ...children) {
        const element = document.createElement(tag);
        Object.entries(props || {}).forEach(([key, value]) => {
            if (value === undefined || value === null || value === false) return;
            if (key === 'class') element.className = value;
            else if (key === 'text') element.textContent = value;
            else if (key === 'dataset') Object.assign(element.dataset, value);
            else if (key.startsWith('on')) element.addEventListener(key.slice(2), value);
            else if (key === 'value') element.value = value;
            else if (key === 'checked') element.checked = Boolean(value);
            else if (value === true) element.setAttribute(key, '');
            else element.setAttribute(key, value);
        });
        children.flat(Infinity).forEach(child => {
            if (child === null || child === undefined || child === false) return;
            element.append(child instanceof Node ? child : document.createTextNode(String(child)));
        });
        return element;
    }
    // Contatore delle schede: un solo numero, sempre il conteggio. Pulito di norma;
    // con una notifica lo stesso numero prende lo sfondo blu (il dettaglio si legge nella sezione).
    function counter(count, notify = false, title = '') {
        return h('span', { class: `ao-counter${notify ? ' ao-counter-notify' : ''}`, title: title || undefined, hidden: !count && !notify }, count || '');
    }

    function showStatus(message, tone = 'success', action = null) {
        const status = root?.querySelector('.ao-status');
        if (!status) return;
        status.replaceChildren(...[h('span', {}, message),
            action ? h('button', { type: 'button', class: 'ao-status-action', onclick: () => { status.hidden = true; action.onclick(); } }, action.label) : null].filter(Boolean));
        status.className = `ao-status ao-tone-${tone}`;
        status.hidden = false;
        clearTimeout(statusTimer);
        statusTimer = setTimeout(() => { status.hidden = true; }, tone === 'danger' || action ? 7000 : 4000);
    }
    // Esegue una chiamata ai servizi mostrando l'esito; restituisce il risultato o null.
    async function run(task, successMessage) {
        if (state.busy) return null;
        state.busy = true;
        root?.classList.add('ao-busy');
        try {
            const result = await task();
            if (successMessage) showStatus(typeof successMessage === 'function' ? successMessage(result) : successMessage);
            return result ?? true;
        } catch (error) {
            console.error('[AMP]', error);
            showStatus(error instanceof ServicesMissingError ? TEXTS.common.servicesMissing : (error?.message || String(error)), 'danger');
            return null;
        } finally {
            state.busy = false;
            root?.classList.remove('ao-busy');
        }
    }

    // Elenchi di In coda, sempre dai servizi.
    async function refreshLists() {
        if (!Services.available()) {
            state.loaded = true;
            render();
            return;
        }
        state.loading = !state.loaded;
        if (state.loading) render();
        try {
            const [queue, ignored] = await Promise.all([Services.listQueue(), Services.listIgnored()]);
            state.queue = queue;
            state.ignored = ignored;
            state.serviceError = '';
        } catch (error) {
            state.serviceError = error?.message || String(error);
        }
        state.loaded = true;
        state.loading = false;
        render();
    }

    // Cambio di scheda: il contenuto nuovo entra nella direzione dello spostamento
    // (da destra andando verso destra, da sinistra tornando indietro), così la navigazione si nota.
    // Contenuto sotto un selettore: nella barra fissa è l'elenco; altrove gli elementi che lo seguono.
    function slideInContent(bar, direction) {
        if (window.matchMedia?.('(prefers-reduced-motion: reduce)').matches) return;
        const fixed = bar.closest('.ao-fixed');
        const targets = fixed
            ? [...fixed.parentElement.children].filter(child => child !== fixed)
            : (() => { const after = []; for (let next = bar.nextElementSibling; next; next = next.nextElementSibling) after.push(next); return after; })();
        targets.forEach(target => target.animate?.([
            { opacity: 0, transform: `translateX(${24 * direction}px)` },
            { opacity: 1, transform: 'none' }
        ], { duration: 280, easing: 'cubic-bezier(.2, .8, .2, 1)' }));
    }

    function buildShell(container) {
        // Una sola barra: freccia per tornare all'add-in e la scheda In coda.
        root = h('div', { id: 'ao-root', dataset: { build: BUILD } },
            h('nav', { class: 'ao-tabs', role: 'tablist' },
                h('button', { type: 'button', class: 'ao-icon-button ao-back', title: TEXTS.back, 'aria-label': TEXTS.back, onclick: () => window.TestPluginAPI?.showMainView?.() }, '←'),
                h('button', { type: 'button', role: 'tab', class: 'ao-tab ao-active', 'aria-selected': 'true', dataset: { view: 'queue' } },
                    h('span', { class: 'ao-tab-label' }, TEXTS.tabs.queue), counter(0)),
                h('span', { class: 'ao-tab-indicator', 'aria-hidden': 'true' })),
            h('div', { class: 'ao-body' }),
            h('div', { class: 'ao-status', role: 'status', 'aria-live': 'polite', hidden: true }),
            // Barra di sistema fissa in fondo: stato dei servizi, sempre visibile.
            h('footer', { class: 'ao-statusbar', role: 'status', 'aria-live': 'polite' }));
        // Scheda aperta: un clic fuori prima la richiude (con l'animazione) e non fa altro;
        // solo il clic successivo apre un'altra richiesta. Scheda e selettore restano liberi.
        root.addEventListener('click', event => {
            if (!state.queueExpanded || state.holdRender) return;
            const target = event.target;
            if (target.closest('.ao-queue-item.ao-expanded, .ao-tabs, .ao-fixed, .ao-status, .ao-statusbar')) return;
            const form = root.querySelector('.ao-queue-item.ao-expanded .ao-inline-form');
            if (!form) return;
            event.preventDefault();
            event.stopPropagation();
            closeTakeCharge(form);
        }, true);
        container.replaceChildren(root);
        // La sottolineatura blu segue la larghezza della scheda.
        if (window.ResizeObserver) new ResizeObserver(placeTabIndicator).observe(root.querySelector('.ao-tabs'));
    }
    function placeTabIndicator() {
        const indicator = root?.querySelector('.ao-tab-indicator');
        const tab = root?.querySelector('.ao-tab');
        if (!indicator || !tab) return;
        indicator.style.width = `${tab.offsetWidth}px`;
        indicator.style.transform = `translateX(${tab.offsetLeft}px)`;
        indicator.classList.add('ao-ready');
    }
    function render() {
        if (!root) return;
        if (state.holdRender) { renderPending = true; return; }
        placeTabIndicator();
        const tabCounter = root.querySelector('.ao-tab .ao-counter');
        tabCounter.textContent = state.queue.length || '';
        tabCounter.hidden = !state.queue.length;
        renderStatusBar();
        const body = root.querySelector('.ao-body');
        // Il selettore Attive / Ignorate sta in cima all'elenco e scivola via scorrendo (watchScroll).
        // La posizione di scorrimento resta uguale tra un aggiornamento e l'altro.
        const scrollKey = state.queueTab;
        const previousScroll = body.dataset.scrollKey === scrollKey ? body.querySelector('.ao-scroll')?.scrollTop || 0 : 0;
        const view = renderQueue();
        const [bar, ...rest] = [...view.children];
        const scroll = h('div', { class: 'ao-scroll' }, h('div', { class: 'ao-fixed' }, bar), ...rest);
        if (!previousScroll) state.barOffset = 0;
        view.replaceChildren(scroll);
        body.replaceChildren(view);
        body.dataset.scrollKey = scrollKey;
        scroll.scrollTop = previousScroll;
        watchScroll(scroll, scroll.firstElementChild);
    }
    // Barra secondaria: sta in cima all'elenco (la sua "casa").
    //  - scorrendo verso il basso scivola via seguendo lo scorrimento, pixel per pixel;
    //  - scorrendo verso l'alto ricompare con un'animazione morbida (discesa con un lieve rimbalzo
    //    e l'ombra che si accende), sovrapposta al contenuto;
    //  - se lo scorrimento verso il basso si ferma a metà, finisce di uscire.
    const BAR_UP_DELTA = 3; // scorrimento minimo verso l'alto per richiamarla
    const BAR_SHADOW_FADE = 40; // px prima di casa in cui l'ombra sfuma
    const BAR_SNAP_DELAY = 140; // ms senza scorrimento prima di completare l'uscita
    function watchScroll(scroll, bar) {
        let last = scroll.scrollTop;
        let snapTimer = null;
        const travel = () => bar.offsetHeight + 12; // fuori vista del tutto (con un margine per l'ombra)
        const apply = (offset, motion = '') => {
            state.barOffset = Math.max(0, Math.min(travel(), offset));
            const home = scroll.scrollTop <= 0;
            bar.classList.toggle('ao-snap', motion === 'snap');
            bar.classList.toggle('ao-reveal', motion === 'reveal');
            bar.style.transform = state.barOffset ? `translateY(${-state.barOffset}px)` : '';
            // Ombra solo quando è sovrapposta e tutta visibile: mentre scivola via non ce l'ha.
            bar.classList.toggle('ao-floating', !home && state.barOffset === 0);
            // L'ombra non si spegne di colpo: negli ultimi pixel prima di "casa" sfuma con lo scorrimento.
            const lift = !home && state.barOffset === 0 ? Math.min(1, scroll.scrollTop / BAR_SHADOW_FADE) : 0;
            bar.style.setProperty('--ao-lift', lift.toFixed(3));
        };
        apply(scroll.scrollTop > 0 ? state.barOffset : 0);
        scroll.addEventListener('scroll', () => {
            const top = scroll.scrollTop;
            const delta = top - last;
            last = top;
            clearTimeout(snapTimer);
            if (top <= 0) return apply(0, bar.classList.contains('ao-reveal') ? 'reveal' : '');
            if (delta < -BAR_UP_DELTA) {
                // Richiamo: se non è già tutta visibile, scende con l'animazione e lì resta.
                if (state.barOffset > 0) apply(0, 'reveal');
                else apply(0, bar.classList.contains('ao-reveal') ? 'reveal' : '');
                return;
            }
            if (delta > 0) {
                apply(state.barOffset + delta);
                snapTimer = setTimeout(() => {
                    if (state.barOffset > 0 && state.barOffset < travel()) apply(travel(), 'snap');
                }, BAR_SNAP_DELAY);
            }
        }, { passive: true });
    }

    // Servizi mancanti, in errore o in caricamento: gli elenchi vuoti non dicono "nessun elemento".
    const servicesBlocked = () => !Services.available() || Boolean(state.serviceError) || state.loading;
    function renderStatusBar() {
        const bar = root?.querySelector('.ao-statusbar');
        if (!bar) return;
        let tone = 'ok';
        let message = '';
        if (!Services.available()) { tone = 'warning'; message = TEXTS.system.servicesMissing; }
        else if (state.serviceError) { tone = 'danger'; message = TEXTS.system.serviceError(state.serviceError); }
        else if (state.loading) { tone = 'neutral'; message = TEXTS.common.loading; }
        else message = TEXTS.system.servicesOk;
        bar.className = `ao-statusbar ao-statusbar-${tone}`;
        bar.replaceChildren(
            h('span', { class: 'ao-statusbar-dot', 'aria-hidden': 'true' }),
            h('span', { class: 'ao-statusbar-text', title: message }, message));
    }
    // Selettore a segmenti. La pillola della voce scelta è un elemento a parte (.ao-segment-thumb):
    // quando la scelta cambia, scivola dalla voce di prima a quella nuova (anche cambiando larghezza).
    // Il pannello si ridisegna a ogni clic, quindi la voce di prima si ricorda per selettore (segmentMemory).
    const segmentMemory = new Map();
    function filterBar(options, value, onchange) {
        const key = options.map(option => option.value).join('|');
        const thumb = h('span', { class: 'ao-segment-thumb', 'aria-hidden': 'true' });
        const bar = h('div', { class: 'ao-segmented', role: 'group' },
            thumb,
            options.map(option => h('button', {
                type: 'button', class: `ao-segment${option.value === value ? ' ao-active' : ''}`, dataset: { value: option.value },
                'aria-pressed': String(option.value === value), onclick: () => onchange(option.value)
            }, option.label, option.count !== undefined ? h('span', { class: 'ao-count' }, option.count) : null)));
        const previous = segmentMemory.get(key);
        segmentMemory.set(key, value);
        requestAnimationFrame(() => {
            if (!bar.isConnected) return;
            const place = (target, animate) => {
                const button = bar.querySelector(`.ao-segment[data-value="${CSS.escape(String(target))}"]`);
                if (!button) return false;
                thumb.classList.toggle('ao-moving', animate);
                thumb.style.width = `${button.offsetWidth}px`;
                thumb.style.transform = `translateX(${button.offsetLeft - 2}px)`; // 2px = bordo interno del selettore
                return true;
            };
            if (previous !== undefined && previous !== value && place(previous, false)) {
                thumb.getBoundingClientRect(); // parte dalla voce di prima
                place(value, true);
                // Anche il contenuto che dipende dal selettore entra nella direzione dello spostamento.
                const values = options.map(option => option.value);
                slideInContent(bar, values.indexOf(value) > values.indexOf(previous) ? 1 : -1);
            } else place(value, false);
            thumb.classList.add('ao-ready');
            // Se il riquadro cambia larghezza, la pillola segue la voce (senza animazione).
            if (window.ResizeObserver) new ResizeObserver(() => { if (!thumb.classList.contains('ao-moving')) place(value, false); }).observe(bar);
            thumb.addEventListener('transitionend', () => thumb.classList.remove('ao-moving'));
        });
        return bar;
    }
    // Selettore di scelta (scheda di In coda: azione e gruppo dell'inoltro). Come il selettore a segmenti,
    // ma all'inizio nessuna voce è scelta e la voce scelta si illumina del colore della sua azione
    // (sfondo pastello, senza bordo). Cambiando scelta la pillola scivola sulla voce nuova.
    const choiceMemory = new Map();
    function choiceBar(key, options, value, onpick) {
        const thumb = h('span', { class: 'ao-choice-thumb', 'aria-hidden': 'true' });
        const tone = options.find(option => option.value === value)?.tone;
        const bar = h('div', { class: `ao-choice${tone ? ` ao-choice-${tone}` : ''}`, role: 'group' },
            thumb,
            options.map(option => h('button', {
                type: 'button', class: `ao-choice-option${option.value === value ? ' ao-active' : ''}`, dataset: { value: option.value },
                'aria-pressed': String(option.value === value), disabled: option.disabled, onclick: () => onpick(option.value)
            }, h('span', { class: 'ao-choice-label' }, option.label), option.sub ? h('span', { class: 'ao-choice-sub' }, option.sub) : null)));
        const previous = choiceMemory.get(key);
        choiceMemory.set(key, value);
        if (value === undefined || value === null) return bar;
        requestAnimationFrame(() => {
            if (!bar.isConnected) return;
            const place = (target, animate) => {
                const button = bar.querySelector(`.ao-choice-option[data-value="${CSS.escape(String(target))}"]`);
                if (!button) return false;
                thumb.classList.toggle('ao-moving', animate);
                thumb.style.width = `${button.offsetWidth}px`;
                thumb.style.transform = `translateX(${button.offsetLeft - 3}px)`; // 3px = bordo interno del selettore
                return true;
            };
            if (previous != null && previous !== value && place(previous, false)) {
                thumb.classList.add('ao-ready');
                thumb.getBoundingClientRect(); // parte dalla voce di prima
                place(value, true);
            } else {
                place(value, false);
                if (previous == null) thumb.classList.add('ao-appearing'); // prima scelta: la pillola compare sul posto
            }
            thumb.classList.add('ao-ready');
            if (window.ResizeObserver) new ResizeObserver(() => { if (!thumb.classList.contains('ao-moving')) place(value, false); }).observe(bar);
            thumb.addEventListener('transitionend', () => thumb.classList.remove('ao-moving'));
        });
        return bar;
    }

    // ======================================================================
    // 3. IN CODA
    // ======================================================================

    // Barra secondaria: Attive (da smistare) | Ignorate (archivio).
    function renderQueue() {
        const active = state.queue.slice().sort((a, b) => String(b.receivedAt).localeCompare(String(a.receivedAt)));
        const showIgnored = state.queueTab === 'ignored';
        return h('div', { class: 'ao-view' },
            filterBar([
                { value: 'active', label: TEXTS.queue.active, count: active.length },
                { value: 'ignored', label: TEXTS.queue.ignoredTab, count: state.ignored.length }
            ], showIgnored ? 'ignored' : 'active', value => { state.queueTab = value; state.queueExpanded = null; render(); }),
            showIgnored
                ? [
                    state.ignored.length ? h('ul', { class: 'ao-list' }, state.ignored.map(ignoredItem))
                        : (servicesBlocked() ? null : h('p', { class: 'ao-empty' }, TEXTS.queue.ignoredEmpty))
                ]
                : (active.length ? h('ul', { class: 'ao-list' }, active.map(item => queueCard(item))) : (servicesBlocked() ? null : h('p', { class: 'ao-empty' }, TEXTS.queue.empty))));
    }
    // Thread ignorato: chi l'ha ignorato si legge solo aprendo la scheda, sopra il testo della mail.
    function ignoredItem(item) {
        return queueCard(item, { allowIgnore: false,
            info: TEXTS.queue.ignoredBy(item.discardedBy || TEXTS.common.none, formatWhen(item.discardedAt)) });
    }
    // Scheda chiusa: oggetto (su una riga, troncato; intero solo da aperta), ora e mittente. Il clic la apre.
    function queueCard(item, { allowIgnore = true, info = null } = {}) {
        const expanded = state.queueExpanded === item.conversationId;
        // Un secondo clic mentre la scheda si sta ancora aprendo (es. doppio clic) non la richiude.
        const toggle = () => {
            if (!expanded) return openTakeCharge(item);
            if (!state.holdRender) closeTakeCharge(queueDetailsOf(item.conversationId));
        };
        // Aperta: la scheda si solleva in primo piano e il titolo prende il colore dell'azione scelta.
        const choice = expanded ? state.queueChoice[item.conversationId]?.action : undefined;
        return h('li', {
            class: `ao-item ao-queue-item${expanded ? ' ao-expanded' : ''}${expanded && state.liftFrom === item.conversationId ? ' ao-unlifted' : ''}`,
            dataset: { conversation: item.conversationId, choice },
            // Chiusa: si apre con un clic in qualsiasi punto della scheda (anche sui bordi).
            onclick: () => { if (!expanded) openTakeCharge(item); }
        },
            h('div', {
                class: 'ao-item-main', tabindex: '0', role: 'button', 'aria-expanded': String(expanded),
                onclick: event => { event.stopPropagation(); toggle(); }, onkeydown: event => { if (event.key === 'Enter') toggle(); }
            },
                h('div', { class: 'ao-queue-row' },
                    h('span', { class: 'ao-item-title', title: expanded ? undefined : item.subject || undefined }, item.subject || '(senza oggetto)'),
                    h('span', { class: 'ao-item-time' }, formatWhen(item.receivedAt))),
                h('div', { class: 'ao-queue-row' },
                    h('span', { class: 'ao-item-sub' }, item.sender?.name || item.sender?.email || ''))),
            expanded ? queueDetails(item, allowIgnore, info) : null);
    }
    // Scheda aperta: testo della mail, poi il selettore dell'azione (nessuna scelta all'inizio).
    // Prendi in carico non è attiva; Inoltra mostra i gruppi; Ignora chiede "Confermi?". In fondo la freccia richiude.
    function queueDetails(item, allowIgnore, info = null) {
        const body = state.mailBodies[item.messageId];
        const choice = state.queueChoice[item.conversationId] || {};
        const actions = [
            { value: 'take', label: TEXTS.queue.take, tone: 'accent', disabled: true },
            { value: 'forward', label: TEXTS.queue.forward, tone: 'forward' },
            allowIgnore ? { value: 'ignore', label: TEXTS.queue.ignore, tone: 'danger' } : null
        ].filter(Boolean);
        return h('div', { class: 'ao-inline-form' },
            info ? h('p', { class: 'ao-hint ao-queue-info' }, info) : null,
            // Testo limitato a PREVIEW_MAX_CHARS; la barra di scorrimento compare solo a scheda aperta del tutto.
            h('div', { class: `ao-mail-preview${state.previewScroll === item.conversationId ? ' ao-scrollable' : ''}` },
                clipText(body ?? item.preview ?? '', PREVIEW_MAX_CHARS), body === undefined ? h('span', { class: 'ao-hint' }, ` ${TEXTS.common.loading}`) : null),
            choiceBar(`action:${item.conversationId}`, actions, choice.action, value => pickQueueAction(item, value)),
            h('div', { class: 'ao-choice-zone' }, choice.action ? h('div', { class: 'ao-choice-content' }, queueActionContent(item, choice)) : null),
            h('button', {
                type: 'button', class: 'ao-collapse', 'aria-label': TEXTS.common.close, title: TEXTS.common.close,
                onclick: event => closeTakeCharge(event.currentTarget.closest('.ao-inline-form'))
            }, svgIcon(CHEVRON_UP_ICON)));
    }
    function queueActionContent(item, choice) {
        const mail = { conversationId: item.conversationId };
        const card = button => button.closest('.ao-queue-item');
        if (choice.action === 'forward') {
            const group = FORWARD_GROUPS.find(entry => entry.id === choice.group);
            return [
                choiceBar(`group:${item.conversationId}`, FORWARD_GROUPS.map(entry => ({ value: entry.id, label: entry.id, sub: entry.name, tone: 'forward' })),
                    choice.group, value => { choice.group = value; render(); }),
                h('div', { class: 'ao-actions' },
                    h('button', { type: 'button', class: 'ao-button ao-primary ao-fill-forward', disabled: !group, onclick: event => forwardFromQueue(card(event.currentTarget), mail, group) }, TEXTS.queue.forward))
            ];
        }
        if (choice.action === 'ignore') return h('div', { class: 'ao-actions' },
            h('button', { type: 'button', class: 'ao-button ao-primary ao-fill-danger', onclick: event => ignoreFromQueue(card(event.currentTarget), mail) }, TEXTS.common.confirm));
        return null;
    }
    // Cambio di azione nella stessa scheda: il contenuto di prima svanisce, l'altezza si adatta, il nuovo compare (dissolvenza).
    async function pickQueueAction(item, action) {
        const id = item.conversationId;
        const choice = state.queueChoice[id] || (state.queueChoice[id] = {});
        if (choice.action === action || state.holdRender) return;
        holdRender();
        const zone = queueDetailsOf(id)?.querySelector('.ao-choice-zone');
        const card = zone?.closest('.ao-queue-item');
        if (card) card.dataset.choice = action; // il titolo cambia colore mentre il contenuto svanisce
        const from = zone?.offsetHeight || 0;
        const old = zone?.firstElementChild;
        if (old) await playOn(old, [{ opacity: 1 }, { opacity: 0 }], MOTION.swap * .45, 'linear');
        choice.action = action;
        releaseRender();
        render();
        const next = queueDetailsOf(id)?.querySelector('.ao-choice-zone');
        const content = next?.firstElementChild;
        if (!next || !content) return;
        await holdingRender(async () => {
            content.style.opacity = '0';
            const result = await playOn(next, [{ height: `${from}px` }, { height: `${next.scrollHeight}px` }], MOTION.swap);
            if (result) { result.animation.cancel(); next.style.overflow = result.overflow; }
            const shown = await playOn(content, [{ opacity: 0 }, { opacity: 1 }], MOTION.swap * .55, 'linear');
            if (shown) { shown.animation.cancel(); content.style.overflow = shown.overflow; }
            content.style.opacity = '';
        });
    }
    // Apre la scheda (una alla volta) con l'animazione e carica il testo della mail.
    async function openTakeCharge(item) {
        await whenIdle(); // se un'altra scheda si sta chiudendo, si apre subito dopo (il clic non va perso)
        if (state.queueExpanded === item.conversationId) return;
        state.queueExpanded = item.conversationId;
        state.queueChoice = { [item.conversationId]: {} }; // ogni apertura riparte senza scelta
        state.previewScroll = null;
        // Se il testo arriva in fretta si apre già completo; altrimenti si apre con le prime righe
        // e, quando arriva, la tendina si allunga senza scatti.
        const loading = state.mailBodies[item.messageId] === undefined && item.messageId
            ? Services.getMailBody(item.messageId).then(text => text || item.preview || '').catch(() => item.preview || '')
            : null;
        let arrived = false;
        if (loading) {
            loading.then(text => { state.mailBodies[item.messageId] = text; arrived = true; });
            await Promise.race([loading, wait(BODY_WAIT_MS)]);
        }
        if (state.queueExpanded !== item.conversationId) return;
        state.liftFrom = item.conversationId; // si disegna ancora "a terra", poi si solleva con la transizione
        render();
        state.liftFrom = null;
        const lifted = queueDetailsOf(item.conversationId)?.closest('.ao-queue-item');
        lifted?.getBoundingClientRect();
        lifted?.classList.remove('ao-unlifted');
        const complete = !loading || arrived; // aperta già con il testo intero
        await holdingRender(() => unfold(queueDetailsOf(item.conversationId)));
        if (complete) return allowPreviewScroll(item.conversationId);
        await loading;
        if (state.queueExpanded !== item.conversationId) return;
        const before = queueDetailsOf(item.conversationId)?.offsetHeight || 0;
        render();
        const form = queueDetailsOf(item.conversationId);
        if (form && before) await holdingRender(() => playOn(form, [{ maxHeight: `${before}px` }, { maxHeight: `${form.scrollHeight}px` }], 240)
            .then(result => { if (result) { result.animation.cancel(); form.style.overflow = result.overflow; } }));
        allowPreviewScroll(item.conversationId);
    }
    const BODY_WAIT_MS = 250;
    const PREVIEW_MAX_CHARS = 500;
    const clipText = (text, max) => (text.length > max ? `${text.slice(0, max).trimEnd()}…` : text);
    // Barra di scorrimento del testo della mail: solo ad animazione finita (durante le tendine non deve comparire).
    function allowPreviewScroll(conversationId) {
        if (state.queueExpanded !== conversationId) return;
        state.previewScroll = conversationId;
        queueDetailsOf(conversationId)?.querySelector('.ao-mail-preview')?.classList.add('ao-scrollable');
    }
    function stopPreviewScroll(element) {
        state.previewScroll = null;
        element?.querySelector('.ao-mail-preview')?.classList.remove('ao-scrollable');
    }
    // Chiudi (o clic sull'intestazione o fuori): la scheda si richiude con la stessa animazione, al contrario.
    async function closeTakeCharge(form) {
        const closing = state.queueExpanded;
        if (!closing) return;
        form?.closest('.ao-expanded')?.classList.remove('ao-expanded'); // il bordo blu va via subito
        stopPreviewScroll(form);
        holdRender();
        await fold(form);
        if (state.queueExpanded === closing) { state.queueExpanded = null; delete state.queueChoice[closing]; } // solo se nel frattempo non se n'è aperta un'altra
        releaseRender();
        render();
    }
    const queueDetailsOf = conversationId => root?.querySelector(`.ao-queue-item[data-conversation="${CSS.escape(conversationId)}"] .ao-inline-form`);

    // ---------------------------- Animazioni --------------------------------
    // Regola unica per tutte le tendine del pannello: il testo non si piega mai.
    //  - unfold: prima si apre lo spazio, poi il contenuto compare;
    //  - fold:   prima il contenuto si affievolisce, poi lo spazio si chiude;
    //  - removeCard: una scheda lascia l'elenco e le altre salgono.
    // Le azioni su una scheda di In coda usano fold e poi la loro uscita (outcome*).
    const MOTION = { unfold: 320, fold: 320, swap: 240, hold: 150, exit: 380, remove: 260, collapse: 380 };
    const reducedMotion = () => Boolean(window.matchMedia?.('(prefers-reduced-motion: reduce)').matches);
    const wait = ms => new Promise(next => setTimeout(next, ms));
    function playOn(element, frames, duration, easing = 'ease-in-out') {
        if (!element?.animate || reducedMotion()) return Promise.resolve();
        const overflow = element.style.overflow;
        element.style.overflow = 'hidden';
        const animation = element.animate(frames, { duration, easing, fill: 'forwards' });
        return animation.finished.catch(() => {}).then(() => ({ animation, overflow }));
    }
    async function unfold(element) {
        if (!element) return;
        const height = `${element.scrollHeight}px`;
        const pad = getComputedStyle(element).paddingTop;
        const result = await playOn(element, [
            { maxHeight: '0px', paddingTop: '0px', opacity: 0 },
            { maxHeight: height, paddingTop: pad, opacity: 0, offset: .55 }, // prima lo spazio…
            { maxHeight: height, paddingTop: pad, opacity: 1 } // …poi il contenuto compare
        ], MOTION.unfold);
        if (result) { result.animation.cancel(); element.style.overflow = result.overflow; } // torna libera (altezza naturale)
    }
    function fold(element) {
        if (!element) return Promise.resolve();
        const height = `${element.scrollHeight}px`;
        const pad = getComputedStyle(element).paddingTop;
        return playOn(element, [
            { maxHeight: height, paddingTop: pad, opacity: 1 },
            { maxHeight: height, paddingTop: pad, opacity: 0, offset: .4 }, // prima si affievolisce…
            { maxHeight: '0px', paddingTop: '0px', opacity: 0 } // …poi lo spazio si chiude
        ], MOTION.fold);
    }
    function removeCard(card) {
        if (!card) return Promise.resolve();
        return playOn(card, [
            { maxHeight: `${card.offsetHeight}px` },
            { maxHeight: '0px', paddingTop: '0px', paddingBottom: '0px', marginTop: '-6px', borderWidth: '0px' }
        ], MOTION.remove);
    }
    // Durante un'animazione il pannello non si ridisegna: i ridisegni richiesti nel frattempo
    // (dati aggiornati, clic) si fanno appena finisce, e le azioni che aspettano (whenIdle) ripartono.
    let renderHolds = 0;
    let renderPending = false;
    let idleWaiters = [];
    function holdRender() {
        renderHolds += 1;
        state.holdRender = true;
    }
    function releaseRender() {
        if (!renderHolds) return;
        renderHolds -= 1;
        if (renderHolds) return;
        state.holdRender = false;
        idleWaiters.splice(0).forEach(resume => resume());
        if (renderPending) {
            renderPending = false;
            render();
        }
    }
    const whenIdle = () => (renderHolds ? new Promise(resume => idleWaiters.push(resume)) : Promise.resolve());
    async function holdingRender(task) {
        holdRender();
        try { return await task(); } finally { releaseRender(); }
    }

    // Esito su una scheda di In coda, in un movimento solo: mentre la tendina si chiude la scheda prende
    // pian piano il colore dell'esito (banner che affiora), poi esce di lato e infine lo spazio si chiude.
    // Ignora → rosso, esce a sinistra; Inoltra → viola, esce a destra; Prendi in carico → blu, si richiude verso l'alto.
    // Se ServiceNow non ha ancora risposto, il banner resta fermo con la rotellina finché non arriva l'esito.
    const ARCHIVE_ICON = '<svg viewBox="0 0 20 20" width="16" height="16" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="2.5" y="3.5" width="15" height="4" rx="1"/><path d="M4 7.5v8a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1v-8"/><path d="M8 11h4"/></svg>';
    const CHECK_ICON = '<svg viewBox="0 0 20 20" width="16" height="16" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4.5 10.5l3.5 3.5 7.5-8"/></svg>';
    const SPINNER = '<span class="ao-outcome-spinner" aria-hidden="true"></span>';
    const FORWARD_ICON = '<svg viewBox="0 0 20 20" width="16" height="16" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M11.5 4.5l5 5-5 5"/><path d="M16.5 9.5H8a4.5 4.5 0 0 0-4.5 4.5v1.5"/></svg>';
    function outcomeStart(card, tone, html) {
        if (!card || reducedMotion()) return;
        holdRender(); // rilasciato da outcomeRelease, dopo che gli elenchi sono aggiornati
        card.classList.add('ao-resolving');
        stopPreviewScroll(card);
        card.aoOutcome = { tone, html, cancelled: false, held: true };
        const banner = h('div', { class: 'ao-outcome', role: 'status' });
        banner.innerHTML = html;
        card.append(banner);
        card.classList.add(`ao-outcome-${tone}`);
        // il colore affiora piano e diventa pieno solo alla fine, quando la tendina è del tutto chiusa
        banner.animate([{ opacity: 0 }, { opacity: .12, offset: .5 }, { opacity: .4, offset: .8 }, { opacity: 1 }], { duration: MOTION.fold, easing: 'linear', fill: 'backwards' });
        // anche l'intestazione (oggetto, mittente, ora, etichetta) si affievolisce e svanisce sotto il colore
        // (entro metà della chiusura il testo è già svanito del tutto)
        card.querySelector('.ao-item-main')?.animate([{ opacity: 1 }, { opacity: 0 }], { duration: MOTION.fold * .5, easing: 'ease-out', fill: 'forwards' });
        card.aoOutcome.ready = fold(card.querySelector('.ao-inline-form')).then(() => { if (!card.aoOutcome.cancelled) card.classList.add('ao-outcome-full'); });
    }
    function outcomeText(card, html) {
        if (!card?.aoOutcome) return;
        card.aoOutcome.html = html;
        const banner = card.querySelector('.ao-outcome');
        if (banner) banner.innerHTML = html;
    }
    async function outcomeEnd(card) {
        if (card?.aoOutcome) {
            await card.aoOutcome.ready;
            await wait(MOTION.hold); // un attimo per leggere l'esito
            if (card.aoOutcome.tone === 'accent') {
                // Prendi in carico: la scheda blu si richiude verso l'alto (la scritta svanisce mentre si chiude)
                [...(card.querySelector('.ao-outcome')?.children || [])].forEach(child =>
                    child.animate?.([{ opacity: 1 }, { opacity: 0 }], { duration: MOTION.collapse * .5, easing: 'linear', fill: 'forwards' }));
                await playOn(card, [
                    { maxHeight: `${card.offsetHeight}px` },
                    { maxHeight: '0px', paddingTop: '0px', paddingBottom: '0px', marginTop: '-6px', borderWidth: '0px' }
                ], MOTION.collapse, 'cubic-bezier(.4, 0, .2, 1)');
                return;
            }
            // Ignora a sinistra, Inoltra a destra: solo scorrimento laterale (la scheda resta intera),
            // lo spazio si chiude dopo che è uscita
            const direction = card.aoOutcome.tone === 'danger' ? -1 : 1;
            await playOn(card, [
                { transform: 'none' },
                { transform: `translateX(${direction * 110}%)` }
            ], MOTION.exit, 'cubic-bezier(.5, 0, .75, 0)');
            await removeCard(card);
        }
    }
    function outcomeRelease(card) {
        if (!card?.aoOutcome?.held) return;
        card.aoOutcome.held = false;
        releaseRender();
    }
    // Errore: niente banner e la scheda torna com'era (al prossimo ridisegno).
    function outcomeCancel(card) {
        if (card?.aoOutcome) card.aoOutcome.cancelled = true;
        card?.querySelector('.ao-outcome')?.remove();
        card?.classList.remove('ao-resolving', 'ao-outcome-accent', 'ao-outcome-forward', 'ao-outcome-danger', 'ao-outcome-full');
        card?.querySelector('.ao-item-main')?.getAnimations().forEach(animation => animation.cancel());
        outcomeRelease(card);
    }
    async function ignoreFromQueue(card, mail) {
        outcomeStart(card, 'danger', `${ARCHIVE_ICON}<span>${TEXTS.queue.ignoredDone}</span>`);
        const ok = await run(() => Services.discardMail(mail.conversationId, currentUser()));
        if (ok) await outcomeEnd(card);
        else outcomeCancel(card);
        if (ok) {
            if (state.queueExpanded === mail.conversationId) state.queueExpanded = null;
            delete state.queueChoice[mail.conversationId];
            await refreshLists();
        }
        outcomeRelease(card);
        render();
    }
    // Inoltro a un gruppo: viola, esce a destra come la presa in carico.
    async function forwardFromQueue(card, mail, group) {
        if (!group) return;
        outcomeStart(card, 'forward', `${SPINNER}<span>${TEXTS.queue.forwardWorking}</span>`);
        const ok = await run(() => Services.forwardMail(mail.conversationId, group.id, currentUser()));
        if (ok) {
            outcomeText(card, `${FORWARD_ICON}<span>${TEXTS.queue.forwardDone(group.id)}</span>`);
            await outcomeEnd(card);
        } else outcomeCancel(card);
        if (ok) {
            if (state.queueExpanded === mail.conversationId) state.queueExpanded = null;
            delete state.queueChoice[mail.conversationId];
            await refreshLists();
        }
        outcomeRelease(card);
        render();
    }

    // ======================================================================
    // 4. STILI (solo dentro #ao-root: non toccano il resto dell'add-in)
    // ======================================================================

    function themeCssVariables() {
        const kebab = value => value.replace(/[A-Z]/g, char => `-${char.toLowerCase()}`);
        const lines = [];
        ['colors', 'fontSize', 'fontWeight', 'radius', 'shadow', 'state', 'layout'].forEach(group => {
            const prefix = { colors: '', fontSize: 'fs-', fontWeight: 'fw-', radius: 'radius-', shadow: 'shadow-', state: '', layout: '' }[group];
            Object.entries(THEME[group]).forEach(([key, value]) => lines.push(`--ao-${prefix}${kebab(key)}: ${value};`));
        });
        lines.push(`--ao-font: ${THEME.font.family};`);
        const dark = Object.entries(THEME.darkColors).map(([key, value]) => `--ao-${kebab(key)}: ${value};`).join(' ') + ' color-scheme: dark;';
        return `#ao-root { ${lines.join(' ')} }
#ao-root[data-theme="dark"] { ${dark} }
@media (prefers-color-scheme: dark) { #ao-root:not([data-theme="light"]) { ${dark} } }`;
    }
    const STYLES = `
/* L'area AMP ha l'altezza esatta del riquadro visibile: schede e barra di sistema ferme, scorre solo il contenuto */
#ao-root { position: relative; display: flex; flex-direction: column; overflow: hidden; font-family: var(--ao-font); font-size: var(--ao-fs-base); color: var(--ao-text); background: var(--ao-bg); line-height: 1.4; }
#ao-root *, #ao-root *::before, #ao-root *::after { box-sizing: border-box; }
#ao-root [hidden] { display: none !important; }
#ao-root.ao-busy { cursor: progress; }
#ao-root button { font: inherit; color: inherit; }

/* Intestazione e schede */
#ao-root .ao-tabs { position: relative; flex: none; display: flex; align-items: center; gap: 2px; padding: 0 4px; background: var(--ao-surface); border-bottom: 1px solid var(--ao-line); }
#ao-root .ao-tabs .ao-icon-button { flex: none; font-size: var(--ao-fs-lg); }
#ao-root .ao-tab { flex: 1; display: flex; align-items: center; justify-content: center; gap: 5px; padding: 9px 4px 7px; border: 0; border-bottom: 2px solid transparent; background: none; color: var(--ao-text-muted); font-size: var(--ao-fs-lg); font-weight: var(--ao-fw-normal); cursor: pointer; }
#ao-root .ao-tab:hover { color: var(--ao-text); background: var(--ao-bg-soft); }
#ao-root .ao-tab.ao-active { color: var(--ao-text-accent); font-weight: var(--ao-fw-bold); }
#ao-root .ao-tab-indicator { position: absolute; left: 0; bottom: -1px; height: 2px; width: 0; background: var(--ao-accent); pointer-events: none; opacity: 0; }
#ao-root .ao-tab-indicator.ao-ready { opacity: 1; }
#ao-root .ao-tab-indicator.ao-moving { transition: transform .32s cubic-bezier(.3, .8, .2, 1), width .32s cubic-bezier(.3, .8, .2, 1); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-tab-indicator.ao-moving { transition: none; } }
/* Scheda aperta: leggero zoom della scritta (e del numero), che resta finché è selezionata */
#ao-root .ao-tab > * { transition: transform .35s ease-in-out; }
#ao-root .ao-tab.ao-active > * { transform: scale(1.04); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-tab > * { transition: none; } }
/* Schede: il conteggio è un numero pulito; con una notifica lo stesso numero ha lo sfondo blu. */
#ao-root .ao-counter { display: inline-block; min-width: 18px; height: 18px; padding: 0 4px; border-radius: var(--ao-radius-badge); color: var(--ao-text-muted); font-size: var(--ao-fs-md); font-weight: var(--ao-fw-medium); line-height: 18px; text-align: center; }
#ao-root .ao-tab.ao-active .ao-counter { color: var(--ao-text-accent); }
#ao-root .ao-counter-notify, #ao-root .ao-tab.ao-active .ao-counter-notify { background: var(--ao-accent); color: var(--ao-white); font-weight: var(--ao-fw-bold); }
#ao-root .ao-body { flex: 1; min-height: 0; display: flex; flex-direction: column; }
#ao-root .ao-view { flex: 1; min-height: 0; display: flex; flex-direction: column; }
/* In alto, fermi: barra secondaria o "←", con una linea sottile come confine */
/* Barra secondaria scontornata: nessuna fascia né separatore, solo il menu.
   A casa sta in cima all'elenco; fuori casa compare sovrapposto al contenuto, con un'ombra sotto il menu stesso. */
#ao-root .ao-fixed { position: sticky; top: 0; z-index: 3; display: grid; gap: 10px; padding-top: 12px; pointer-events: none; will-change: transform; }
/* Completamento del movimento quando lo scorrimento si ferma a metà */
#ao-root .ao-fixed.ao-snap { transition: transform .18s ease-out; }
/* Ricomparsa: discesa morbida con un lieve rimbalzo; l'ombra si accende mentre arriva */
#ao-root .ao-fixed.ao-reveal { transition: transform .42s cubic-bezier(.2, 1.3, .4, 1); }
#ao-root .ao-fixed.ao-reveal > * { transition: box-shadow .35s ease-out .08s, background-color .35s ease-out .08s; }
#ao-root .ao-fixed > * { pointer-events: auto; }
/* Ombra (e fondo della riga "indietro") graduati da --ao-lift: 0 a casa, 1 sovrapposta al contenuto */
#ao-root .ao-fixed { --ao-lift: 0; }
#ao-root .ao-fixed .ao-segmented,
#ao-root .ao-fixed .ao-back-row { box-shadow: 0 3px 10px rgba(0,0,0, calc(.18 * var(--ao-lift))); transition: box-shadow .25s ease-out, background-color .25s ease-out; }
#ao-root .ao-fixed .ao-back-row { justify-self: start; margin-left: -10px; padding: 4px 10px; border-radius: var(--ao-radius-base); background: color-mix(in srgb, var(--ao-surface) calc(100% * var(--ao-lift)), transparent); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-fixed.ao-snap, #ao-root .ao-fixed.ao-reveal, #ao-root .ao-fixed.ao-reveal > * { transition: none; } }
/* Sotto la linea scorre il contenuto, a tutta larghezza (barra di scorrimento sottile, spazio riservato) */
#ao-root .ao-scroll { flex: 1; min-height: 0; display: grid; align-content: start; gap: 10px; padding: 0 12px 12px; scroll-padding-top: 60px; overflow-y: auto; overflow-x: hidden; scrollbar-gutter: stable; scrollbar-width: thin; scrollbar-color: var(--ao-line-strong) transparent; }
#ao-root .ao-scroll::-webkit-scrollbar { width: 6px; }
#ao-root .ao-scroll::-webkit-scrollbar-thumb { border-radius: 3px; background: var(--ao-line-strong); }
#ao-root .ao-scroll::-webkit-scrollbar-track { background: transparent; }

/* Pulsanti */
#ao-root .ao-button { display: inline-flex; align-items: center; justify-content: center; gap: 6px; min-height: var(--ao-control-height); padding: 0 12px; border: 1px solid var(--ao-line-strong); border-radius: var(--ao-radius-base); background: var(--ao-surface); font-weight: var(--ao-fw-medium); cursor: pointer; }
#ao-root .ao-button:hover:where(:not(:disabled)) { background: var(--ao-bg-soft); border-color: var(--ao-text-subtle); }
#ao-root .ao-primary { background: var(--ao-accent); border-color: var(--ao-accent); color: var(--ao-white); }
#ao-root .ao-primary:hover:where(:not(:disabled)) { background: var(--ao-accent-dark); border-color: var(--ao-accent-dark); }
#ao-root .ao-subtle { border-color: transparent; background: transparent; color: var(--ao-text-muted); }
#ao-root .ao-small { min-height: 26px; padding: 0 9px; font-size: var(--ao-fs-md); }
#ao-root .ao-icon-button { display: inline-flex; align-items: center; justify-content: center; min-width: 28px; height: 28px; padding: 0 6px; border: 1px solid transparent; border-radius: var(--ao-radius-base); background: none; color: var(--ao-text-muted); cursor: pointer; }
#ao-root .ao-icon-button:hover { background: var(--ao-bg-soft); border-color: var(--ao-line); }
#ao-root :disabled { opacity: var(--ao-disabled-opacity); cursor: not-allowed; }
#ao-root :focus-visible { outline: none; box-shadow: var(--ao-shadow-focus); }

/* Filtri e scelte a segmenti */
#ao-root .ao-segmented { position: relative; display: flex; padding: 2px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); }
#ao-root .ao-segment { flex: 1; display: inline-flex; align-items: center; justify-content: center; gap: 5px; min-height: 28px; padding: 0 6px; border: 0; border-radius: calc(var(--ao-radius-base) - 2px); background: none; color: var(--ao-text-muted); font-size: var(--ao-fs-md); cursor: pointer; white-space: nowrap; }
/* Niente sfondo al passaggio del mouse: l'unica pillola è quella della voce scelta, che scivola */
#ao-root .ao-segment { position: relative; z-index: 1; transition: color .25s; }
#ao-root .ao-segment:hover { color: var(--ao-text); }
#ao-root .ao-segment.ao-active { color: var(--ao-text-accent); font-weight: var(--ao-fw-medium); }
#ao-root .ao-segment-thumb { position: absolute; top: 2px; bottom: 2px; left: 2px; width: 0; border-radius: calc(var(--ao-radius-base) - 2px); background: var(--ao-accent-bg); opacity: 0; pointer-events: none; }
#ao-root .ao-segment-thumb.ao-ready { opacity: 1; }
#ao-root .ao-segment-thumb.ao-moving { transition: transform .32s cubic-bezier(.3, .8, .2, 1), width .32s cubic-bezier(.3, .8, .2, 1); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-segment-thumb.ao-moving, #ao-root .ao-segment { transition: none; } }
#ao-root .ao-count { color: var(--ao-text-subtle); font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-normal); }
/* Elenco di In coda */
#ao-root .ao-list { display: grid; gap: 6px; margin: 0; padding: 0; list-style: none; }
#ao-root .ao-item { display: grid; gap: 3px; padding: 9px 10px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); box-shadow: var(--ao-shadow-card); cursor: pointer; }
/* Passaggio del mouse, uguale in tutte le sezioni: sfondo grigio leggerissimo, testo invariato.
   Il blu (titolo e bordo) è riservato alla scheda selezionata (modulo o dettagli aperti), che non reagisce all'hover. */
#ao-root .ao-item { transition: background-color .15s; }
#ao-root .ao-item:not(.ao-expanded):hover { background: var(--ao-hover); }
#ao-root .ao-expanded, #ao-root .ao-expanded:hover { background: var(--ao-surface); }
#ao-root .ao-item-title { font-weight: var(--ao-fw-medium); overflow-wrap: anywhere; }
#ao-root .ao-item-sub { color: var(--ao-text-muted); font-size: var(--ao-fs-md); overflow-wrap: anywhere; }
#ao-root .ao-item-time { color: var(--ao-text-subtle); font-size: var(--ao-fs-sm); white-space: nowrap; }
#ao-root .ao-empty { margin: 8px 0; color: var(--ao-text-subtle); text-align: center; }
#ao-root .ao-hint { color: var(--ao-text-subtle); font-size: var(--ao-fs-sm); }
#ao-root .ao-actions { display: flex; flex-wrap: wrap; gap: 8px; }
#ao-root .ao-actions .ao-primary { flex: 1; }
/* In coda: riga con azioni e modulo compatto */
#ao-root .ao-queue-item { gap: 6px; }
#ao-root .ao-item-main { display: grid; grid-template-columns: minmax(0, 1fr) auto; gap: 3px 10px; cursor: pointer; border-radius: calc(var(--ao-radius-base) - 2px); }
#ao-root .ao-inline-form { display: grid; gap: 8px; padding-top: 8px; border-top: 1px solid var(--ao-line-light); }
/* Scheda di In coda aperta: testo della mail */
#ao-root .ao-mail-preview { max-height: 180px; overflow: hidden; padding: 8px 10px; border-radius: var(--ao-radius-base); background: var(--ao-bg-soft); color: var(--ao-text); font-size: var(--ao-fs-md); line-height: 1.5; white-space: pre-line; overflow-wrap: anywhere; scrollbar-width: thin; }
#ao-root .ao-mail-preview.ao-scrollable { overflow: auto; }
#ao-root .ao-spacer { flex: 1; }
#ao-root .ao-queue-item .ao-item-main { display: grid; grid-template-columns: minmax(0, 1fr); gap: 3px; }
#ao-root .ao-queue-row { display: flex; align-items: center; justify-content: space-between; gap: 8px; min-width: 0; }
#ao-root .ao-queue-row > :first-child { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
#ao-root .ao-queue-row > :last-child { flex: none; }
#ao-root .ao-queue-item.ao-expanded .ao-queue-row .ao-item-title { white-space: normal; overflow-wrap: anywhere; }
#ao-root .ao-queue-info { margin: 0; }
/* Prendi in carico: su una riga e largo, occupa lo spazio libero accanto a Ignora e Chiudi */
#ao-root .ao-queue-item .ao-actions { flex-wrap: nowrap; }
#ao-root .ao-queue-item .ao-actions .ao-primary { flex: 1; white-space: nowrap; }
#ao-root .ao-queue-item .ao-actions .ao-spacer { display: none; }
#ao-root .ao-queue-item { position: relative; overflow: hidden; }
/* Scheda aperta: selettore dell'azione (e del gruppo), contenuto dell'azione, freccia per richiudere */
#ao-root .ao-choice { position: relative; display: flex; padding: 3px; border-radius: var(--ao-radius-base); background: var(--ao-bg); }
#ao-root .ao-choice-option { position: relative; z-index: 1; flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center; min-height: 30px; padding: 4px 6px; border: 0; border-radius: calc(var(--ao-radius-base) - 2px); background: none; color: var(--ao-text-muted); font-size: var(--ao-fs-md); font-weight: var(--ao-fw-medium); cursor: pointer; white-space: nowrap; line-height: 1.25; transition: color .25s; }
#ao-root .ao-choice-sub { font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-normal); }
#ao-root .ao-choice-option:hover:not(.ao-active) { color: var(--ao-text); }
#ao-root .ao-choice-thumb { position: absolute; top: 3px; bottom: 3px; left: 3px; width: 0; border-radius: calc(var(--ao-radius-base) - 2px); background: var(--ao-choice-bg); opacity: 0; pointer-events: none; transition: background-color .25s; }
#ao-root .ao-choice-thumb.ao-ready { opacity: 1; }
#ao-root .ao-choice-thumb.ao-appearing { animation: ao-fade-in .2s linear; }
#ao-root .ao-choice-thumb.ao-moving { transition: transform .32s cubic-bezier(.3, .8, .2, 1), width .32s cubic-bezier(.3, .8, .2, 1), background-color .25s; }
@keyframes ao-fade-in { from { opacity: 0; } }
#ao-root .ao-choice-accent { --ao-choice-bg: color-mix(in srgb, var(--ao-accent) 16%, var(--ao-surface)); --ao-choice-text: var(--ao-text-accent); } /* più deciso dell'azzurro dei filtri: si stacca dal fondo grigio */
#ao-root .ao-choice-forward { --ao-choice-bg: var(--ao-forward-bg); --ao-choice-text: var(--ao-forward-text); }
#ao-root .ao-choice-danger { --ao-choice-bg: var(--ao-danger-bg); --ao-choice-text: var(--ao-danger); }
#ao-root .ao-choice-option.ao-active { color: var(--ao-choice-text); }
#ao-root .ao-choice-zone:empty { display: none; }
#ao-root .ao-choice-content { display: grid; gap: 8px; }
#ao-root .ao-fill-forward, #ao-root .ao-fill-forward:hover:where(:not(:disabled)) { background: var(--ao-forward); border-color: var(--ao-forward); color: var(--ao-white); }
#ao-root .ao-fill-forward:hover:where(:not(:disabled)) { filter: brightness(.92); }
#ao-root .ao-fill-danger, #ao-root .ao-fill-danger:hover:where(:not(:disabled)) { background: #d13438; border-color: #d13438; color: var(--ao-white); }
#ao-root .ao-fill-danger:hover:where(:not(:disabled)) { filter: brightness(.92); }
#ao-root .ao-collapse { display: flex; justify-content: center; align-items: center; justify-self: center; width: 44px; height: 22px; margin-top: -2px; padding: 0; border: 0; border-radius: var(--ao-radius-base); background: none; color: var(--ao-text-muted); cursor: pointer; }
#ao-root .ao-collapse:hover { background: var(--ao-hover); color: var(--ao-text); }
#ao-root .ao-svg-icon { display: inline-flex; }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-choice-thumb, #ao-root .ao-choice-option { transition: none; animation: none; } }
/* Esito su una scheda di In coda: banner colorato (le animazioni sono nel blocco Animazioni dello script) */
#ao-root .ao-queue-item { transition: background-color .3s, border-color .3s, transform .32s cubic-bezier(.3, .8, .2, 1), box-shadow .32s ease, filter .32s ease; }
/* Scheda aperta: in primo piano (si solleva e si ingrandisce appena), bordo chiaro come le altre */
#ao-root .ao-queue-item.ao-expanded, #ao-root .ao-queue-item.ao-expanded:hover {
    z-index: 1; border-color: var(--ao-line); transform: translateY(-2px) scale(1.015);
    box-shadow: 0 6px 16px rgba(0,0,0,.12), 0 1px 3px rgba(0,0,0,.08);
}
#ao-root .ao-queue-item.ao-expanded.ao-unlifted, #ao-root .ao-queue-item.ao-resolving { transform: none; box-shadow: var(--ao-shadow-card); }
/* Con una scheda aperta le altre si sfocano (tornano nitide appena si richiude) */
#ao-root .ao-list:has(.ao-queue-item.ao-expanded:not(.ao-unlifted):not(.ao-resolving)) .ao-queue-item:not(.ao-expanded) { filter: blur(5px); }
/* Titolo: neutro finché non si sceglie l'azione, poi del suo colore */
#ao-root .ao-queue-item.ao-expanded .ao-item-title { color: var(--ao-text); transition: color .25s; }
#ao-root .ao-queue-item.ao-expanded[data-choice="take"] .ao-item-title { color: var(--ao-text-accent); }
#ao-root .ao-queue-item.ao-expanded[data-choice="forward"] .ao-item-title { color: var(--ao-forward-text); }
#ao-root .ao-queue-item.ao-expanded[data-choice="ignore"] .ao-item-title { color: var(--ao-danger); }
#ao-root .ao-resolving { pointer-events: none; }
#ao-root .ao-outcome-accent { --ao-outcome: var(--ao-accent); }
#ao-root .ao-outcome-danger { --ao-outcome: #d13438; }
#ao-root .ao-outcome-forward { --ao-outcome: var(--ao-forward); }
/* sfondo e bordo della scheda prendono il colore solo a tendina chiusa (prima lo dà solo il banner, che affiora) */
#ao-root .ao-queue-item.ao-outcome-full { background: var(--ao-outcome); border-color: var(--ao-outcome); }
#ao-root .ao-outcome {
    position: absolute; inset: 0; z-index: 2; display: flex; align-items: center; gap: 8px; padding: 0 12px;
    background: var(--ao-outcome); color: #fff; font-size: var(--ao-fs-md); font-weight: var(--ao-fw-medium);
}
#ao-root .ao-outcome-spinner { width: 14px; height: 14px; border: 2px solid rgba(255,255,255,.45); border-top-color: #fff; border-radius: 50%; animation: ao-spin .8s linear infinite; }
@keyframes ao-spin { to { transform: rotate(360deg); } }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-queue-item, #ao-root .ao-outcome { transition: none; animation: none; } }
/* Barra di sistema in fondo all'area (stato dei servizi, allineamento), sempre visibile */
#ao-root .ao-statusbar { flex: none; display: flex; align-items: center; gap: 6px; height: 28px; padding: 0 12px; border-top: 1px solid var(--ao-line); background: var(--ao-surface); color: var(--ao-text-muted); font-size: var(--ao-fs-sm); }
#ao-root .ao-statusbar-dot { width: 7px; height: 7px; flex: none; border-radius: 50%; background: var(--ao-success); }
#ao-root .ao-statusbar-warning .ao-statusbar-dot { background: var(--ao-warning); }
#ao-root .ao-statusbar-danger .ao-statusbar-dot { background: var(--ao-danger); }
#ao-root .ao-statusbar-danger { color: var(--ao-danger); }
#ao-root .ao-statusbar-neutral .ao-statusbar-dot { background: var(--ao-text-subtle); }
#ao-root .ao-statusbar-text { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
#ao-root .ao-statusbar-sync { flex: none; }

/* Messaggi temporanei, sopra la barra di sistema */
#ao-root .ao-status:not([hidden]) { display: flex; align-items: center; justify-content: space-between; gap: 10px; }
#ao-root .ao-status-action { flex: none; padding: 0; border: 0; background: none; color: var(--ao-text-accent); font: inherit; font-weight: var(--ao-fw-bold); cursor: pointer; }
#ao-root .ao-status-action:hover { text-decoration: underline; }
#ao-root .ao-status { position: absolute; left: 12px; right: 12px; bottom: 36px; z-index: 3; margin: 0; padding: 8px 10px; border-radius: var(--ao-radius-base); font-size: var(--ao-fs-md); font-weight: var(--ao-fw-medium); box-shadow: 0 2px 8px rgba(0,0,0,.12); }
#ao-root .ao-status.ao-tone-success { background: var(--ao-success-bg); color: var(--ao-success); }
#ao-root .ao-status.ao-tone-warning { background: var(--ao-warning-bg); color: var(--ao-warning); }
#ao-root .ao-status.ao-tone-danger { background: var(--ao-danger-bg); color: var(--ao-danger); }
`;
    function injectStyles() {
        if (document.getElementById('ao-styles')) return;
        document.head.append(h('style', { id: 'ao-styles' }, themeCssVariables() + STYLES));
    }

    // ======================================================================
    // 5. AVVIO E COLLEGAMENTO CON L'ADD-IN
    // ======================================================================

    let initialized = false;
    // Asincrona come quella originale dell'area AMP: taskpane.js chiama
    // AmpModule.initialize().catch(...) all'avvio dell'add-in.
    async function initialize() {
        if (initialized) return;
        const container = document.getElementById('amp-view');
        if (!container) return;
        initialized = true;
        injectStyles();
        buildShell(container);
        fillPane();
        // Il riquadro di Outlook si può ridimensionare: l'altezza dell'area si ricalcola.
        window.addEventListener('resize', fitHeight);
        if (window.ResizeObserver) new ResizeObserver(() => fitHeight()).observe(scrollContainer());
        document.getElementById('open-amp-button')?.addEventListener('click', show);
        Services.subscribe(() => { refreshLists(); });
        render();
        await refreshLists();
    }
    // L'area AMP occupa tutto il riquadro: annulla il margine interno del contenitore dell'add-in
    // e prende l'altezza esatta che resta sotto l'intestazione, così la pagina non scorre più:
    // intestazione e schede restano ferme, scorre solo il contenuto (.ao-body).
    function fillPane() {
        const host = root?.parentElement?.parentElement;
        if (!host) return;
        const style = getComputedStyle(host);
        root.style.margin = `-${style.paddingTop} -${style.paddingRight} -${style.paddingBottom} -${style.paddingLeft}`;
        fitHeight();
    }
    // Contenitore che scorre: nell'add-in la pagina, nell'anteprima il riquadro simulato.
    function scrollContainer() {
        for (let element = root?.parentElement; element && element !== document.body; element = element.parentElement) {
            if (/(auto|scroll)/.test(getComputedStyle(element).overflowY)) return element;
        }
        return document.scrollingElement || document.documentElement;
    }
    function fitHeight() {
        if (!root || root.offsetParent === null) return; // area nascosta
        const scroller = scrollContainer();
        scroller.scrollTop = 0;
        const isPage = scroller === document.scrollingElement || scroller === document.documentElement;
        const top = root.getBoundingClientRect().top - (isPage ? 0 : scroller.getBoundingClientRect().top);
        const visible = isPage ? window.innerHeight : scroller.clientHeight;
        root.style.height = `${Math.max(240, visible - top)}px`;
        // Quello che segue l'area (es. il piè di pagina dell'add-in) non deve far scorrere la pagina.
        const overflow = scroller.scrollHeight - scroller.clientHeight;
        if (overflow > 0) root.style.height = `${Math.max(240, visible - top - overflow)}px`;
        scroller.scrollTop = 0;
    }
    function show() {
        initialize();
        fillPane();
        const mainView = document.getElementById('main-view');
        const headerView = document.getElementById('header-view');
        const ampView = document.getElementById('amp-view');
        if (mainView) mainView.hidden = true;
        if (headerView) headerView.hidden = true;
        if (ampView) ampView.hidden = false;
        render();
        fitHeight();
    }
    window.AmpModule = Object.freeze({ initialize, show });
    const logInitError = error => console.error('AMP init error:', error);
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', () => initialize().catch(logInitError));
    else initialize().catch(logInitError);
})();
