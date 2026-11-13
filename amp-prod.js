/*
 * AMP - SNOW Assistant: area AMP dell'add-in Outlook — VERSIONE DI PRODUZIONE.
 *
 * Due schede: In coda (mail della casella condivisa) e Debug (le sezioni dell'add-in di jak).
 * In coda:
 *   - Attive / Ignorate;
 *   - scheda aperta: testo della mail (con più messaggi nel thread, le frecce per scorrerli) e selettore
 *     Prendi in carico · Inoltra · Ignora. Prendi in carico c'è ma non è attiva (fino a nuovo ordine);
 *     Inoltra (gruppi esterni della configurazione operativa) e Ignora sì.
 * Ticket, vista Mail, Impostazioni, etichette e stile vetro sono solo nel pannello di sviluppo
 * (fork Mark7458/TestOutlookAddIn-AMP, ramo amp-panel: amp.js). Questo file si ricava da lì togliendo il resto:
 * una modifica a In coda si fa in amp.js e si riporta qui.
 *
 * Collegamento con l'add-in di jak (nessun suo file viene modificato):
 *   - espone window.AmpModule { initialize, show }; il pannello si apre da solo all'avvio;
 *   - la scheda Debug contiene la sua schermata (#main-view, con i Developer Tools) e la vista dell'header
 *     (#header-view), spostate così come sono: il suo codice continua a funzionare;
 *   - se la pagina non ha #amp-view, il pannello lo crea in cima a <main>;
 *   - l'utente è quello di Outlook (TestPluginAPI.getCurrentMessageContext, altrimenti Office).
 *
 * Servizi (window.AmpBackend, da collegare alla Function App; vedi amp-servizi.md):
 *   listQueue(), listIgnored(), discardMail(conversationId, utente)       → necessari
 *   getMailBody(messageId), listThreadMessages(conversationId),
 *   forwardMail(conversationId, idGruppo, utente), subscribe(listener)     → facoltativi
 * Gruppi dell'inoltro: dalla configurazione operativa del core (localStorage test-plugin-operational-config),
 *   quelli attivi con type "external".
 */
(function () {
    'use strict';

    // Indice dei blocchi
    //  0. Aspetto: tema, testi                         3. In coda
    //  1. Servizi, configurazione operativa e utente   4. Stili (usano i valori del blocco 0)
    //  2. Interfaccia: struttura e navigazione         5. Avvio e collegamento con l'add-in

    // ======================================================================
    // 0. ASPETTO: TEMA E TESTI
    // ======================================================================

    // Tutti i valori di aspetto: il CSS li usa come variabili --ao-* (vedi themeCssVariables()).
    const THEME = {
        colors: {
            accent: '#0078d4', // pulsanti principali, scheda attiva, voci selezionate
            textAccent: '#106ebe', // tutte le scritte blu (numeri dei ticket, link, schede attive…)
            accentBg: '#eff6fc',
            white: '#fff', // testo su sfondo blu (pulsanti, notifiche)
            surface: 'rgba(255, 255, 255, .82)', // sfondo di card, campi, schede, barra di sistema
            text: '#242424',
            textMuted: '#605e5c',
            textSubtle: '#8a8886',
            bg: 'rgba(110, 140, 190, .1)', // zone incassate (campi raggruppati, dettagli)
            hover: 'rgba(255, 255, 255, .9)', // sfondo delle schede al passaggio del mouse
            lineLight: 'rgba(0, 0, 0, .07)',
            line: 'rgba(255, 255, 255, .75)',
            lineStrong: 'rgba(0, 0, 0, .14)',
            success: '#107c10',
            successBg: '#dff6dd',
            warning: '#8a5a00',
            warningBg: '#fff4ce',
            danger: '#a4262c',
            dangerBg: '#fde7e9',
            forward: '#7c5bbf', // Inoltra (In coda): pulsante ed esito
            // Selettore della scheda aperta (Inoltra, Ignora): pillola e scritta della voce scelta
            choiceForwardBg: '#f1ecf9',
            choiceForwardText: '#6b47b0',
            choiceDangerBg: '#fde7e9',
            choiceDangerText: '#a4262c',
            // Stile vetro: superfici semitrasparenti e sfocate sopra uno sfondo sfumato. Gli elementi usano solo questi nomi,
            // così chiaro e scuro cambiano tutto insieme.
            backdrop: 'transparent', // lo sfondo sfumato lo dipinge il contenitore dell'intestazione (pageBackdrop), dietro a tutto
            blur: 'blur(14px) saturate(1.2)', // sfocatura di quello che sta dietro alle superfici
            navBlur: 'blur(16px) saturate(1.3)', // sfocatura della fascia blu (intestazione e schede)
            navInk: '#fff', // scritte e icone dell'intestazione e delle schede
            navInkSoft: 'rgba(255, 255, 255, .72)', // schede non scelte
            raisedShadow: '0 6px 16px rgba(0, 0, 0, .12), 0 1px 3px rgba(0, 0, 0, .08)', // scheda sollevata (in viaggio, al clic)
            openShadow: '0 0 18px rgba(80, 150, 255, .45), 0 6px 16px rgba(0, 0, 0, .1)', // scheda aperta
            itemHover: 'var(--ao-surface)', // vetro: la scheda non si schiarisce; bordo vicino al colore della scheda e alone azzurro
            itemHoverEdge: 'rgba(255, 255, 255, .95)',
            itemHoverGlow: '0 0 16px rgba(110, 165, 255, .38)',
            openEdge: 'var(--ao-item-hover-edge)', // scheda aperta: tiene il bordo del passaggio del mouse
            quoteBg: 'var(--ao-bg)', // description come citazione incassata dentro la scheda
            segmentTrack: 'rgba(214, 224, 240, .6)', // un po' più scuro delle card: la voce scelta si stacca
            segmentHover: 'rgba(255, 255, 255, .4)', // voce non scelta di un selettore al passaggio del mouse
            choiceTrack: 'var(--ao-segment-track)', // fondo del selettore delle azioni (Prendi in carico / Inoltra / Ignora)
            selection: 'var(--ao-surface)', // voce scelta: come una card a riposo
            selectionLine: 'var(--ao-line)', // bordo della voce scelta
            selectionShadow: '0 1px 2px rgba(0, 0, 0, .06)', // ombra della voce scelta
            selectionTone: '0%', // quanto colore dell'azione nella voce scelta del selettore delle azioni
            selectionMotion: '.44s cubic-bezier(.32, 1.12, .5, 1)', // scivola con un leggero rimbalzo finale, come iOS
            // Colori delle azioni (pulsanti, esiti)
            toneTake: '#0f5fbf', // zaffiro
            toneForward: '#7b3fbf', // ametista
            toneIgnore: '#c4183c', // rubino
            // Pulsanti pieni del colore dell'azione, scritta bianca
            toneFill: '86%', // vetro colorato
            toneSolid: '0%', // quanto colore nella scheda a esito concluso
            toneLine: 'rgba(255, 255, 255, .22)', // bordo dei pulsanti
            toneHoverFilter: 'brightness(1.08)',
            btnHighlight: 'inset 0 1px 0 rgba(255, 255, 255, .22)', // filo di luce sul bordo alto dei pulsanti
            spinnerTrack: '35%',
        },
        // Tema scuro: segue Outlook (prefers-color-scheme). Testi e colori di stato come l'add-in di jak in tema scuro.
        darkColors: {
            textAccent: '#6cb8f6',
            accentBg: '#19344a',
            text: '#f3f2f1',
            textMuted: '#c8c6c4',
            textSubtle: '#a19f9d',
            success: '#6ccb5f',
            successBg: '#1f3a1d',
            warning: '#f2c661',
            warningBg: '#3d3214',
            danger: '#ffb3b8',
            dangerBg: '#442726',
            // Tema scuro: viola e rosso accesi quanto il blu (prima risultavano spenti)
            choiceForwardBg: 'color-mix(in srgb, #7c5bbf 26%, #2d2c2b)',
            choiceForwardText: '#b79bff',
            choiceDangerBg: 'color-mix(in srgb, #d13438 24%, #2d2c2b)',
            choiceDangerText: '#ff8f94',
            // Vetro scuro: le stesse superfici, più trasparenti; le pietre delle azioni scurite come l'intestazione
            surface: 'rgba(52, 56, 66, .62)',
            hover: 'rgba(70, 77, 92, .68)',
            bg: 'rgba(0, 0, 0, .2)',
            line: 'rgba(255, 255, 255, .09)',
            lineLight: 'rgba(255, 255, 255, .07)',
            lineStrong: 'rgba(255, 255, 255, .16)',
            navInk: 'rgba(236, 240, 247, .93)', // vetro scuro: bianco meno acceso
            navInkSoft: 'rgba(236, 240, 247, .64)',
            itemHoverEdge: 'rgba(255, 255, 255, .2)',
            itemHoverGlow: '0 0 16px rgba(90, 150, 255, .3)',
            segmentTrack: 'rgba(28, 31, 38, .62)',
            segmentHover: 'rgba(255, 255, 255, .06)',
            selectionShadow: '0 1px 2px rgba(0, 0, 0, .3)',
            openShadow: '0 0 18px rgba(70, 140, 255, .45), 0 6px 16px rgba(0, 0, 0, .3)',
            toneTake: 'color-mix(in srgb, #0f5fbf 82%, #000)',
            toneForward: 'color-mix(in srgb, #7b3fbf 82%, #000)',
            toneIgnore: 'color-mix(in srgb, #c4183c 82%, #000)',
            toneLine: 'rgba(255, 255, 255, .14)',
            btnHighlight: '0 0 transparent'
        },
        // Fuori dall'area AMP: sfondo sfumato dietro intestazione e pannello, e intestazione dell'add-in dello stesso vetro,
        // tinta di blu (riflesso diagonale, filo di luce in alto). La barra delle schede la prolunga.
        pageBackdrop: 'radial-gradient(ellipse 260px 220px at 8% 4%, rgba(121, 184, 255, .45), transparent 70%), radial-gradient(ellipse 240px 220px at 95% 40%, rgba(185, 168, 245, .32), transparent 70%), linear-gradient(180deg, #e9f1fb, #eef0fa) fixed',
        pageDarkBackdrop: 'radial-gradient(ellipse 260px 220px at 8% 4%, rgba(40, 100, 190, .45), transparent 70%), radial-gradient(ellipse 240px 220px at 95% 40%, rgba(110, 80, 190, .3), transparent 70%), linear-gradient(180deg, #1b1e26, #1d1b24) fixed',
        header: {
            light: {
                background: 'linear-gradient(125deg, rgba(255, 255, 255, .28), transparent 45%), linear-gradient(135deg, rgba(0, 90, 158, .9), rgba(0, 120, 212, .84))',
                light: 'inset 0 1px 0 rgba(255, 255, 255, .45)'
            },
            dark: {
                background: 'linear-gradient(125deg, rgba(255, 255, 255, .18), transparent 45%), linear-gradient(135deg, rgba(15, 60, 130, .62), rgba(25, 85, 170, .52))',
                light: 'inset 0 1px 0 rgba(255, 255, 255, .22)'
            }
        },
        font: { family: '"Segoe UI", system-ui, -apple-system, BlinkMacSystemFont, sans-serif' },
        fontSize: { xs: '10px', sm: '11px', md: '12px', base: '13px', lg: '14px', xl: '16px' },
        fontWeight: { normal: '400', medium: '600', bold: '700' },
        radius: { base: '8px', pill: '999px', badge: '5px' }, // badge: notifiche, quadrato con gli angoli smussati
        shadow: {
            card: '0 1px 2px rgba(0,0,0,.06)',
            focus: '0 0 0 2px rgba(0,120,212,.25)'
        },
        state: { disabledOpacity: '.45' },
        layout: { controlHeight: '32px' }
    };

    // Tutte le scritte del pannello. Le voci con (valore) sono funzioni.
    const TEXTS = {
        tabs: { queue: 'In coda', debug: 'Debug' },
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
            ignoredBy: (who, when) => `Ignorata da ${who} · ${when}`,
            messageOf: (index, total, when) => `${index} di ${total} · ${when}`,
            previousMessage: 'Messaggio precedente',
            nextMessage: 'Messaggio successivo'
        }
    };

    // Versione del pannello (attributo data-build dell'area): serve a capire se il browser usa l'ultima.
    const BUILD = '2026-11-13 18:30';

    // ======================================================================
    // 1. SERVIZI, CONFIGURAZIONE OPERATIVA E UTENTE
    // ======================================================================

    const REQUIRED_SERVICES = ['listQueue', 'listIgnored', 'discardMail'];
    class ServicesMissingError extends Error {}
    const backend = () => window.AmpBackend;
    const call = (name, ...args) => (typeof backend()?.[name] === 'function'
        ? Promise.resolve(backend()[name](...args)) : Promise.reject(new ServicesMissingError(name)));
    const Services = {
        available: () => REQUIRED_SERVICES.every(name => typeof backend()?.[name] === 'function'),
        listQueue: () => call('listQueue'),
        listIgnored: () => call('listIgnored'),
        discardMail: (conversationId, user) => call('discardMail', conversationId, user),
        forwardMail: (conversationId, group, user) => call('forwardMail', conversationId, group, user),
        getMailBody: messageId => call('getMailBody', messageId),
        listThreadMessages: conversationId => call('listThreadMessages', conversationId),
        subscribe: listener => backend()?.subscribe?.(listener) || (() => {})
    };

    // Configurazione operativa scaricata dal core dell'add-in (agenzie, operatori, gruppi).
    // Serve per menu e visibilità; l'autorizzazione resta al backend.
    const PLUGIN_CONFIG_KEY = 'test-plugin-operational-config';
    function pluginConfig() {
        try { return JSON.parse(localStorage.getItem(PLUGIN_CONFIG_KEY) || 'null')?.payload ?? null; }
        catch { return null; }
    }
    // Gruppi a cui si inoltra: quelli attivi di tipo "external" (es. ES, PEL, PO). La sigla va davanti all'oggetto: [ES].
    function forwardGroups() {
        return Object.entries(pluginConfig()?.groups || {})
            .filter(([, group]) => group?.enabled && group.type === 'external')
            .map(([id, group]) => ({ id, tag: group.shortName || id, name: group.displayName || '' }))
            .sort((a, b) => a.tag.localeCompare(b.tag, 'it'));
    }

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
    // 2. INTERFACCIA: STRUTTURA E NAVIGAZIONE
    // ======================================================================

    const CHEVRON_UP_ICON = '<svg viewBox="0 0 20 20" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l5-5 5 5"/></svg>';
    const CHEVRON_LEFT_ICON = '<svg viewBox="0 0 20 20" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12.5 5l-5 5 5 5"/></svg>';
    const CHEVRON_RIGHT_ICON = '<svg viewBox="0 0 20 20" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M7.5 5l5 5-5 5"/></svg>';
    // Icone delle azioni (selettore della scheda aperta)
    const ACTION_ICONS = {
        take: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6.5V4h14v2.5a2 2 0 0 0 0 4V16H3v-5.5a2 2 0 0 0 0-4z"/><path d="M10 7.5v5M7.5 10h5"/></svg>',
        forward: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M11.5 4.5l5 5-5 5"/><path d="M16.5 9.5H8a4.5 4.5 0 0 0-4.5 4.5v1.5"/></svg>',
        ignore: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><rect x="2.5" y="3.5" width="15" height="4" rx="1"/><path d="M4 7.5v8a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1v-8M8 11h4"/></svg>',
    };
    const ACTION_TONES = { take: 'accent', forward: 'forward', ignore: 'danger' };
    function svgIcon(markup) {
        const icon = h('span', { class: 'ao-svg-icon', 'aria-hidden': 'true' });
        icon.innerHTML = markup;
        return icon;
    }

    // Icone delle schede principali
    const TAB_ICONS = {
        queue: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"><path d="M3 11l2.5-6h9L17 11v5H3z"/><path d="M3 11h4l1 2h4l1-2h4"/></svg>',
        debug: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M12.5 3.5a3.5 3.5 0 0 0-3.3 4.6L3.5 13.8a1.5 1.5 0 0 0 2.1 2.1l5.7-5.7a3.5 3.5 0 0 0 4.6-3.3l-2 2-2-.5-.5-2z"/></svg>'
    };
    // Icone dei selettori a segmenti
    const SEGMENT_ICONS = {
        active: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><circle cx="10" cy="10" r="7"/><path d="M10 6v4l2.5 2"/></svg>',
        ignored: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="2.5" y="3.5" width="15" height="4" rx="1"/><path d="M4 7.5v8a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1v-8M8 11h4"/></svg>',
    };
    // Spazi tra icona, nome e contatore della scheda aperta (usati dal CSS e dalla sottolineatura).
    const TAB_GAP = { label: 9, counter: 5 };
    function tabIcon(view) {
        if (!TAB_ICONS[view]) return null;
        const icon = h('span', { class: 'ao-tab-icon' });
        icon.innerHTML = TAB_ICONS[view];
        return icon;
    }

    const state = {
        view: 'queue', // queue | debug
        queue: [],
        ignored: [], // thread ignorati (archivio)
        loaded: false,
        loading: false,
        serviceError: '',
        queueTab: 'active', // In coda: active (da prendere) | ignored (archivio)
        barOffset: 0, // di quanti px la barra secondaria è scivolata via (segue lo scorrimento)
        queueExpanded: null,
        queueChoice: {}, // scheda aperta: azione scelta nel selettore e gruppo dell'inoltro
        liftFrom: null, // scheda che si sta aprendo: disegnata non sollevata, poi si solleva
        previewScroll: null, // scheda aperta del tutto: il testo della mail può scorrere
        mailBodies: {}, // testo delle mail per l'anteprima (per messageId)
        threadMessages: {}, // messaggi ricevuti di un thread (per conversationId): con più di uno si scorrono con le frecce
        holdRender: false, // true durante un'animazione: il pannello non si ridisegna (le schede sparirebbero di colpo)
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
    // Contatore delle schede (solo sulla scheda aperta). Nascosto con una classe, non con hidden,
    // così compare e sparisce con un'animazione.
    const counter = () => h('span', { class: 'ao-counter ao-counter-off' });

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

    const VIEWS = ['queue', 'debug'];
    // Cambio di scheda: la sottolineatura scivola e il contenuto entra dal lato verso cui ci si sposta.
    function setView(view) {
        if (view === state.view) return;
        const direction = VIEWS.indexOf(view) > VIEWS.indexOf(state.view) ? 1 : -1;
        state.view = view;
        state.barOffset = 0;
        render();
        const shown = root.querySelector(view === 'debug' ? '.ao-debug' : '.ao-body');
        if (shown?.animate && !reducedMotion()) shown.animate([
            { opacity: 0, transform: `translateX(${24 * direction}px)` },
            { opacity: 1, transform: 'none' }
        ], { duration: 280, easing: 'cubic-bezier(.2, .8, .2, 1)' });
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

    let debugHost = null; // contenuto della scheda Debug: le sezioni dell'add-in di jak (sempre nella pagina)
    function buildShell(container) {
        // Una sola barra con le schede In coda e Debug. La scheda scelta mostra icona, nome e contatore.
        // Le sezioni di jak stanno sempre dentro il pannello (nascoste fuori da Debug): il suo codice le cerca per id.
        debugHost = h('div', { class: 'ao-debug', hidden: true },
            ['main-view', 'header-view'].map(id => document.getElementById(id)).filter(Boolean));
        root = h('div', { id: 'ao-root', dataset: { build: BUILD } },
            h('nav', { class: 'ao-tabs', role: 'tablist' },
                VIEWS.map(view => h('button', {
                    type: 'button', role: 'tab', class: 'ao-tab', dataset: { view }, 'aria-label': TEXTS.tabs[view], onclick: () => setView(view)
                }, tabIcon(view), h('span', { class: 'ao-tab-label' }, TEXTS.tabs[view]), counter())),
                h('span', { class: 'ao-tab-indicator', 'aria-hidden': 'true' })),
            h('div', { class: 'ao-body' }),
            debugHost,
            h('div', { class: 'ao-status', role: 'status', 'aria-live': 'polite', hidden: true }));
        // Scheda aperta: un clic fuori prima la richiude (con l'animazione) e non fa altro;
        // solo il clic successivo apre un'altra richiesta. Schede e selettori restano liberi.
        root.addEventListener('click', event => {
            if (!state.queueExpanded || state.holdRender) return;
            if (event.target.closest('.ao-queue-item.ao-expanded, .ao-tabs, .ao-fixed, .ao-status, .ao-debug')) return;
            const form = root.querySelector('.ao-queue-item.ao-expanded .ao-inline-form');
            if (!form) return;
            event.preventDefault();
            event.stopPropagation();
            closeTakeCharge(form);
        }, true);
        container.replaceChildren(root);
        // Se il riquadro cambia larghezza la sottolineatura resta sotto la scheda attiva.
        if (window.ResizeObserver) {
            const observer = new ResizeObserver(() => {
                blendTabsWithHeader();
                const indicator = root.querySelector('.ao-tab-indicator');
                if (indicator && !indicator.classList.contains('ao-moving')) placeTabIndicator();
            });
            observer.observe(root.querySelector('.ao-tabs'));
            const header = document.querySelector('.app-header');
            if (header) observer.observe(header);
        }
        blendTabsWithHeader();
        root.querySelector('.ao-tab-indicator').addEventListener('transitionend', event => event.currentTarget.classList.remove('ao-moving'));
        root.querySelector('.ao-tabs').addEventListener('transitionend', event => {
            if (event.target.classList.contains('ao-tab-label') && event.propertyName === 'max-width') settleTabIndicator();
        });
    }

    // Sottolineatura blu delle schede principali: un elemento solo che scivola sulla scheda attiva
    // (come la pillola dei selettori).
    // Barra delle schede come continuazione dell'intestazione: intestazione e schede sono un'unica lastra di vetro.
    // Sfumatura e riflesso si dipingono sull'altezza di tutte e due e la barra delle schede ne mostra la parte bassa,
    // così l'effetto continua senza giunture.
    function blendTabsWithHeader() {
        const nav = root?.querySelector('.ao-tabs');
        const header = document.querySelector('.app-header');
        if (!nav || !header || !nav.offsetWidth) return;
        const H = header.offsetHeight;
        const whole = { backgroundSize: `${header.offsetWidth}px ${H + nav.offsetHeight}px`, backgroundRepeat: 'no-repeat' };
        Object.assign(header.style, whole);
        nav.style.background = '';
        Object.assign(nav.style, whole, { backgroundImage: getComputedStyle(header).backgroundImage, backgroundColor: 'transparent', backgroundPosition: `0 ${-H}px` });
    }
    function placeTabIndicator() {
        const indicator = root?.querySelector('.ao-tab-indicator');
        if (!indicator) return;
        const activeView = state.view;
        const tab = root.querySelector(`.ao-tab[data-view="${activeView}"]`);
        if (!tab) { indicator.classList.remove('ao-ready'); return; }
        const changed = indicator.dataset.view && indicator.dataset.view !== activeView && indicator.classList.contains('ao-ready');
        indicator.classList.toggle('ao-moving', Boolean(changed));
        // Linea corta e arrotondata, centrata sotto icona e nome della scheda.
        // Larghezza a scheda aperta (nome e numero stanno ancora comparendo): dall'icona all'ultima cifra.
        const label = tab.querySelector('.ao-tab-label'), count = tab.querySelector('.ao-counter');
        const content = tab.querySelector('.ao-tab-icon').offsetWidth + TAB_GAP.label + label.scrollWidth
            + (count.classList.contains('ao-counter-off') ? 0 : TAB_GAP.counter + count.scrollWidth);
        const width = Math.min(tab.offsetWidth, Math.max(36, content + 12));
        indicator.style.width = `${width}px`;
        indicator.style.transform = `translateX(${tab.offsetLeft + (tab.offsetWidth - width) / 2}px)`;
        indicator.dataset.view = activeView;
        requestAnimationFrame(() => indicator.classList.add('ao-ready'));
    }
    // A scheda aperta (nome comparso) la sottolineatura si rimette esattamente sotto icona, nome e numero
    // come appaiono davvero (carattere, zoom della scheda), centrata su di loro.
    function settleTabIndicator() {
        const indicator = root?.querySelector('.ao-tab-indicator');
        const tab = root?.querySelector('.ao-tab.ao-active');
        if (!indicator || !tab) return;
        const parts = [...tab.children].filter(part => !part.classList.contains('ao-counter-off') && part.getBoundingClientRect().width > 0);
        const left = Math.min(...parts.map(part => part.getBoundingClientRect().left));
        const right = Math.max(...parts.map(part => part.getBoundingClientRect().right));
        const nav = root.querySelector('.ao-tabs').getBoundingClientRect();
        const width = Math.min(tab.offsetWidth, Math.max(36, right - left + 12));
        indicator.classList.add('ao-moving');
        indicator.style.width = `${width}px`;
        indicator.style.transform = `translateX(${(left + right) / 2 - nav.left - width / 2}px)`;
    }
    function render() {
        if (!root) return;
        if (state.holdRender) { renderPending = true; return; }
        // Solo In coda ha un numero (le richieste attive), sulla scheda aperta.
        const counts = { queue: state.queue.length };
        root.querySelectorAll('.ao-tab').forEach(tab => {
            const active = tab.dataset.view === state.view;
            tab.classList.toggle('ao-active', active);
            tab.setAttribute('aria-selected', String(active));
            tab.title = active ? '' : TEXTS.tabs[tab.dataset.view];
            const count = counts[tab.dataset.view];
            const tabCounter = tab.querySelector('.ao-counter');
            tabCounter.textContent = count || '';
            tabCounter.classList.toggle('ao-counter-off', !count || !active);
        });
        placeTabIndicator(); // dopo i contatori: la sottolineatura comprende il numero della scheda aperta
        const body = root.querySelector('.ao-body');
        body.hidden = state.view !== 'queue';
        debugHost.hidden = state.view !== 'debug';
        if (state.view !== 'queue') return;
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
    // Selettore a segmenti. La pillola della voce scelta è un elemento a parte (.ao-segment-thumb):
    // quando la scelta cambia, scivola dalla voce di prima a quella nuova (anche cambiando larghezza).
    // Il pannello si ridisegna a ogni clic, quindi la voce di prima si ricorda per selettore (segmentMemory).
    // Si ricorda anche dove stava la pillola: ridisegnando con la stessa scelta (es. un clic su una mail in Outlook)
    // compare subito al suo posto, senza un istante di vuoto.
    const segmentMemory = new Map();
    const segmentPlace = new Map();
    const segmentKey = options => options.map(option => option.value).join('|');
    // Mentre la barra cambia forma (filtri, ricerca) la pillola si mette al suo posto solo alla fine del movimento.
    let barSettled = Promise.resolve();
    // grow: la voce scelta è più larga delle altre (selettore con le icone); cambiando scelta le larghezze
    // si scambiano con lo stesso movimento della pillola, che le segue.
    // grow: voce scelta più larga (selettori a icone); key: memoria del selettore; slide: il contenuto che segue entra
    // scorrendo (navigazione). Dentro una scheda (slide: false) il contenuto cambia con swapZone, in dissolvenza.
    function filterBar(options, value, onchange, { grow = false, key: memoryKey = null, slide = true } = {}) {
        const key = memoryKey || segmentKey(options);
        const previous = segmentMemory.get(key);
        const chosen = options.some(option => option.value === value);
        // anche la prima scelta (prima nessuna voce) si allarga con il movimento del selettore
        const growing = grow && segmentMemory.has(key) && chosen && previous !== value;
        const wide = growing ? previous : value; // all'inizio è larga la voce di prima
        const thumb = h('span', { class: 'ao-segment-thumb', 'aria-hidden': 'true' });
        const kept = segmentMemory.get(key) === value && segmentPlace.get(key);
        if (kept) {
            thumb.style.width = kept.width;
            thumb.style.transform = kept.transform;
            thumb.classList.add('ao-ready');
        }
        const bar = h('div', { class: 'ao-segmented', role: 'group' },
            thumb,
            options.map(option => h('button', {
                type: 'button', class: `ao-segment${option.value === value ? ' ao-active' : ''}${grow && option.value === wide ? ' ao-grow' : ''}`, dataset: { value: option.value }, title: option.title, 'aria-label': option.title, disabled: option.disabled,
                'aria-pressed': String(option.value === value), onclick: () => onchange(option.value)
            }, option.label, option.count !== undefined ? h('span', { class: 'ao-count' }, option.count) : null)));
        segmentMemory.set(key, value);
        requestAnimationFrame(() => barSettled.then(() => {
            if (!bar.isConnected) return;
            const place = (target, animate) => {
                const button = bar.querySelector(`.ao-segment[data-value="${CSS.escape(String(target))}"]`);
                if (!button) return false;
                thumb.classList.toggle('ao-moving', animate);
                thumb.style.width = `${button.offsetWidth}px`;
                thumb.style.transform = `translateX(${button.offsetLeft - 2}px)`; // 2px = bordo interno del selettore
                if (!animate) segmentPlace.set(key, { width: thumb.style.width, transform: thumb.style.transform });
                return true;
            };
            const button = target => bar.querySelector(`.ao-segment[data-value="${CSS.escape(String(target))}"]`);
            if (growing && place(previous, false)) {
                // Dove arriverà la voce nuova, già larga: si misura e si torna subito com'era (senza disegnare)
                const from = button(previous), to = button(value);
                from.classList.remove('ao-grow'); to.classList.add('ao-grow');
                const end = { width: `${to.offsetWidth}px`, transform: thumb.style.transform.replace(/-?[\d.]+px/, `${to.offsetLeft - 2}px`) };
                to.classList.remove('ao-grow'); from.classList.add('ao-grow');
                bar.getBoundingClientRect();
                // Larghezze e pillola si muovono insieme, con la stessa durata e la stessa curva
                bar.classList.add('ao-growing');
                from.classList.remove('ao-grow'); to.classList.add('ao-grow');
                thumb.classList.add('ao-moving');
                thumb.style.width = end.width;
                thumb.style.transform = end.transform;
                const values = options.map(option => option.value);
                if (slide) slideInContent(bar, values.indexOf(value) > values.indexOf(previous) ? 1 : -1);
            } else if (growing && place(value, false)) {
                // Prima scelta: la pillola compare sulla voce premuta e si allarga con lei, con lo stesso movimento
                const to = button(value);
                to.classList.add('ao-grow');
                const end = { width: `${to.offsetWidth}px`, transform: `translateX(${to.offsetLeft - 2}px)` };
                to.classList.remove('ao-grow');
                bar.getBoundingClientRect();
                bar.classList.add('ao-growing');
                to.classList.add('ao-grow');
                thumb.classList.add('ao-moving');
                Object.assign(thumb.style, end);
                thumb.animate([{ opacity: 0 }, { opacity: 1 }], { duration: 200, easing: 'ease-out' });
            } else if (previous !== undefined && previous !== value && place(previous, false)) {
                thumb.getBoundingClientRect(); // parte dalla voce di prima
                place(value, true);
                // Anche il contenuto che dipende dal selettore entra nella direzione dello spostamento.
                const values = options.map(option => option.value);
                if (slide) slideInContent(bar, values.indexOf(value) > values.indexOf(previous) ? 1 : -1);
            } else place(value, false);
            thumb.classList.add('ao-ready');
            // Se il riquadro cambia larghezza, la pillola segue la voce (senza animazione).
            if (window.ResizeObserver) new ResizeObserver(() => { if (!thumb.classList.contains('ao-moving')) place(value, false); }).observe(bar);
            thumb.addEventListener('transitionend', () => {
                thumb.classList.remove('ao-moving');
                segmentPlace.set(key, { width: thumb.style.width, transform: thumb.style.transform });
            });
        }));
        return bar;
    }
    // Selettore con le icone: la voce scelta si allarga e mostra icona e nome; le altre solo l'icona (il nome passandoci sopra).
    // Il numero compare solo sulla voce scelta.
    // key: memoria del selettore quando le stesse voci compaiono in più schede (azioni di ogni scheda di In coda);
    // tone: la voce scelta prende il colore della sua azione (ACTION_TONES); senza scelta solo icone (il nome passandoci sopra).
    function iconBar(options, value, onchange, { key = null, tones = false, slide = true } = {}) {
        const memoryKey = key || segmentKey(options);
        const chosen = options.some(option => option.value === value);
        // Il nome compare in dissolvenza solo quando la scelta cambia, non a ogni ridisegno.
        const previous = segmentMemory.get(memoryKey);
        const entering = segmentMemory.has(memoryKey) && chosen && previous !== value;
        // il nome della voce di prima svanisce mentre si stringe, poi si toglie
        const leaving = (option, active) => entering && !active && option.value === previous
            ? h('span', { class: 'ao-segment-name ao-segment-name-out', onanimationend: event => event.currentTarget.remove() }, h('span', null, option.label)) : null;
        const bar = filterBar(options.map(option => {
            const active = option.value === value;
            return {
                value: option.value,
                title: active ? undefined : option.label,
                label: [svgIcon(option.icon), active ? h('span', { class: entering ? 'ao-segment-name ao-segment-name-in' : 'ao-segment-name' }, h('span', null, option.label)) : leaving(option, active)],
                count: active ? option.count : undefined,
                disabled: option.disabled
            };
        }), value, onchange, { grow: true, key: memoryKey, slide });
        bar.classList.add('ao-segmented-icons');
        if (tones && chosen) bar.classList.add('ao-toned', `ao-choice-${ACTION_TONES[value] || 'accent'}`);
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
                type: 'button', class: `ao-choice-option${option.value === value ? ' ao-active' : ''}`, dataset: { value: option.value }, disabled: option.disabled,
                'aria-pressed': String(option.value === value), onclick: () => onpick(option.value)
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

    // ---------------------------- Coda ------------------------------------
    // Coda condivisa con barra secondaria: Attive (da prendere in carico) | Ignorate (archivio).
    function renderQueue() {
        const active = state.queue.slice().sort((a, b) => String(b.receivedAt).localeCompare(String(a.receivedAt)));
        const showIgnored = state.queueTab === 'ignored';
        return h('div', { class: 'ao-view' },
            // Selettore della stessa larghezza del pannello di sviluppo, senza pulsanti ai lati.
            h('div', { class: 'ao-source-bar' }, h('div', { class: 'ao-bar-row' },
                h('span', { class: 'ao-bar-spacer', 'aria-hidden': 'true' }),
                iconBar([
                    { value: 'active', icon: SEGMENT_ICONS.active, label: TEXTS.queue.active, count: active.length },
                    { value: 'ignored', icon: SEGMENT_ICONS.ignored, label: TEXTS.queue.ignoredTab, count: state.ignored.length }
                ], showIgnored ? 'ignored' : 'active', value => { state.queueTab = value; state.queueExpanded = null; render(); }),
                h('span', { class: 'ao-bar-spacer', 'aria-hidden': 'true' }))),
            showIgnored
                ? [
                    state.ignored.length ? h('ul', { class: 'ao-list' }, state.ignored.map(item => ignoredItem(item)))
                        : (servicesBlocked() ? null : h('p', { class: 'ao-empty' }, TEXTS.queue.ignoredEmpty))
                ]
                : (active.length ? h('ul', { class: 'ao-list' }, active.map(item => queueCard(item))) : (servicesBlocked() ? null : h('p', { class: 'ao-empty' }, TEXTS.queue.empty))));
    }
    // Thread ignorato: chi l'ha ignorata si legge solo aprendo la scheda, sopra il testo della mail.
    function ignoredItem(item) {
        return queueCard(item, { allowIgnore: false, info: TEXTS.queue.ignoredBy(item.discardedBy || TEXTS.common.none, formatWhen(item.discardedAt)) });
    }
    // Scheda di In coda (Attive e Ignorate). Chiusa: oggetto, mittente, ora; il clic la apre.
    function queueCard(item, { allowIgnore = true, info = null } = {}) {
        const expanded = state.queueExpanded === item.conversationId;
        // Un secondo clic mentre la scheda si sta ancora aprendo (es. doppio clic) non la richiude.
        const toggle = () => {
            if (!expanded) return openTakeCharge(item);
            if (!state.holdRender) closeTakeCharge(queueDetailsOf(item.conversationId));
        };
        // Aperta: la scheda si solleva in primo piano (ombra, leggero ingrandimento) e il titolo prende il colore dell'azione scelta.
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
                // Due righe sempre uguali: oggetto (su una riga, troncato; intero solo da aperta) e ora;
                // mittente e casella condivisa.
                h('div', { class: 'ao-queue-row' },
                    h('span', { class: 'ao-item-title', title: expanded ? undefined : item.subject || undefined }, item.subject || '(senza oggetto)'),
                    h('span', { class: 'ao-item-time' }, formatWhen(item.receivedAt))),
                h('div', { class: 'ao-queue-row' },
                    h('span', { class: 'ao-item-sub' }, item.sender?.name || item.sender?.email || ''))),
            expanded ? queueDetails(item, allowIgnore, info) : null);
    }
    // Scheda aperta: testo della mail, poi il selettore dell'azione (nessuna scelta all'inizio: si decide dopo aver letto).
    // Sotto compare il contenuto dell'azione scelta, con il suo pulsante; in fondo la freccia richiude la scheda.
    //  - Prendi in carico: agenzia e attività;  - Inoltra: gruppo (ES, PEL, PO);  - Ignora: "Confermi?".
    function queueDetails(item, allowIgnore, info = null) {
        const messages = threadMessagesOf(item);
        const shown = messages.length > 1 ? chosenMessage(item) : null; // con più messaggi: quello che si sta guardando
        const body = state.mailBodies[shown?.messageId || item.messageId];
        const choice = state.queueChoice[item.conversationId] || {};
        const actions = [
            { value: 'take', label: TEXTS.queue.take, tone: 'accent', disabled: true }, // non attiva, fino a nuovo ordine
            // senza gruppi esterni nella configurazione non c'è a chi inoltrare
            { value: 'forward', label: TEXTS.queue.forward, tone: 'forward', disabled: !forwardGroups().length },
            allowIgnore ? { value: 'ignore', label: TEXTS.queue.ignore, tone: 'danger' } : null
        ].filter(Boolean);
        return h('div', { class: 'ao-inline-form' },
            info ? h('p', { class: 'ao-hint ao-queue-info' }, info) : null,
            shown ? messagePicker(item, messages, shown) : null,
            // Testo limitato a PREVIEW_MAX_CHARS; la barra di scorrimento compare solo a scheda aperta del tutto.
            h('div', { class: `ao-mail-preview${state.previewScroll === item.conversationId ? ' ao-scrollable' : ''}` },
                h('span', { class: 'ao-mail-text' }, clipText(body ?? shown?.preview ?? item.preview ?? '', PREVIEW_MAX_CHARS),
                    body === undefined ? h('span', { class: 'ao-hint' }, ` ${TEXTS.common.loading}`) : null)),
            iconBar(actions.map(action => ({ ...action, icon: ACTION_ICONS[action.value] })), choice.action, value => pickQueueAction(item, value), { key: `action:${item.conversationId}`, tones: true, slide: false }),
            choiceZone(`queue:${item.conversationId}`, choice.action ? queueActionContent(item, choice) : null),
            h('button', {
                type: 'button', class: 'ao-collapse', 'aria-label': TEXTS.common.close, title: TEXTS.common.close,
                onclick: event => closeTakeCharge(event.currentTarget.closest('.ao-inline-form'))
            }, svgIcon(CHEVRON_UP_ICON)));
    }
    // Messaggi del thread in In coda: con più messaggi l'operatore li scorre con le frecce (si parte dall'ultimo).
    const threadMessagesOf = item => state.threadMessages[item.conversationId] || [];
    function chosenMessage(item) {
        const messages = threadMessagesOf(item);
        return messages[state.queueChoice[item.conversationId]?.at ?? messages.length - 1]
            || { messageId: item.messageId, sender: item.sender, receivedAt: item.receivedAt, preview: item.preview };
    }
    function messagePicker(item, messages, shown) {
        const at = messages.indexOf(shown);
        const step = (direction, label) => h('button', {
            type: 'button', class: 'ao-message-step', 'aria-label': label, title: label, disabled: at + direction < 0 || at + direction >= messages.length,
            onclick: () => showQueueMessage(item, at + direction)
        }, svgIcon(direction < 0 ? CHEVRON_LEFT_ICON : CHEVRON_RIGHT_ICON));
        return h('div', { class: 'ao-message-nav' },
            step(-1, TEXTS.queue.previousMessage),
            h('span', { class: 'ao-message-where' }, TEXTS.queue.messageOf(at + 1, messages.length, formatWhen(shown.receivedAt))),
            step(1, TEXTS.queue.nextMessage));
    }
    // Un altro messaggio del thread: il riquadro grigio resta al suo posto; svanisce solo il testo, il riquadro si adatta
    // all'altezza del nuovo e questo compare.
    async function showQueueMessage(item, index) {
        if (state.holdRender) return;
        const id = item.conversationId;
        const choice = state.queueChoice[id];
        const message = threadMessagesOf(item)[index];
        if (!message || !choice) return;
        const form = queueDetailsOf(id);
        const before = form?.offsetHeight || 0;
        const boxBefore = form?.querySelector('.ao-mail-preview')?.offsetHeight || 0;
        const changing = () => [...(queueDetailsOf(id)?.querySelectorAll('.ao-message-where, .ao-mail-text') || [])];
        const body = state.mailBodies[message.messageId] === undefined
            ? Services.getMailBody(message.messageId).then(text => text || message.preview || '').catch(() => message.preview || '') : null;
        await holdingRender(async () => {
            await Promise.all(changing().map(element => element.animate([{ opacity: 1 }, { opacity: 0 }], { duration: MOTION.hold, easing: 'ease-out', fill: 'forwards' }).finished)).catch(() => {});
            if (body) state.mailBodies[message.messageId] = await body;
            choice.at = index;
        });
        render();
        const next = queueDetailsOf(id);
        const box = next?.querySelector('.ao-mail-preview');
        changing().forEach(element => element.animate([{ opacity: 0 }, { opacity: 1 }], { duration: MOTION.swap, easing: 'ease-in' }));
        const settle = result => { if (result) { result.animation.cancel(); box.style.overflow = result.overflow; } };
        await holdingRender(() => Promise.all([
            box && boxBefore && box.offsetHeight !== boxBefore ? playOn(box, [{ height: `${boxBefore}px` }, { height: `${box.offsetHeight}px` }], MOTION.swap).then(settle) : null,
            next && before && next.scrollHeight !== before ? playOn(next, [{ maxHeight: `${before}px` }, { maxHeight: `${next.scrollHeight}px` }], MOTION.swap)
                .then(result => { if (result) { result.animation.cancel(); next.style.overflow = result.overflow; } }) : null
        ]));
    }
    function queueActionContent(item, choice) {
        const mail = { conversationId: item.conversationId };
        const card = button => button.closest('.ao-queue-item');
        if (choice.action === 'forward') {
            const groups = forwardGroups();
            const group = groups.find(entry => entry.id === choice.group);
            return [
                choiceBar(`group:${item.conversationId}`, groups.map(entry => ({ value: entry.id, label: entry.tag, sub: entry.name, tone: 'forward' })),
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
        const card = queueDetailsOf(id)?.closest('.ao-queue-item');
        if (card) card.dataset.choice = action; // il titolo cambia colore mentre il contenuto svanisce
        await swapZone(`queue:${id}`, () => { choice.action = action; });
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
        if (loading) loading.then(text => { state.mailBodies[item.messageId] = text; arrived = true; });
        const listing = state.threadMessages[item.conversationId] ? null : Services.listThreadMessages(item.conversationId)
            .then(list => { state.threadMessages[item.conversationId] = list || []; }).catch(() => { state.threadMessages[item.conversationId] = []; });
        if (loading || listing) await Promise.race([Promise.all([loading, listing]), wait(BODY_WAIT_MS)]);
        listing?.then(() => { if (state.queueExpanded === item.conversationId) render(); }); // arrivati dopo: le frecce compaiono
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
        const preview = queueDetailsOf(conversationId)?.querySelector('.ao-mail-preview');
        if (!preview || preview.scrollHeight <= preview.clientHeight + 1) return; // il testo ci sta: niente barra
        state.previewScroll = conversationId;
        preview.classList.add('ao-scrollable');
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
        // anche l'intestazione (oggetto, mittente, ora) si affievolisce e svanisce sotto il colore
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
            outcomeText(card, `${FORWARD_ICON}<span>${TEXTS.queue.forwardDone(group.tag)}</span>`);
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
    // Contenuto che dipende dal selettore dentro una scheda (azioni di In coda): quello vecchio svanisce,
    // lo spazio si adatta, il nuovo compare. Sempre la stessa sequenza, ovunque.
    const choiceZone = (id, content) => h('div', { class: 'ao-choice-zone', dataset: { zone: id } }, content ? h('div', { class: 'ao-choice-content' }, content) : null);
    async function swapZone(id, change) {
        if (state.holdRender) return;
        const find = () => root.querySelector(`.ao-choice-zone[data-zone="${CSS.escape(id)}"]`);
        holdRender();
        const zone = find();
        const from = zone?.offsetHeight || 0;
        const old = zone?.firstElementChild;
        if (old) await playOn(old, [{ opacity: 1 }, { opacity: 0 }], MOTION.swap * .45, 'linear');
        change();
        releaseRender();
        render();
        const next = find();
        if (!next) return;
        const content = next.firstElementChild;
        await holdingRender(async () => {
            if (content) content.style.opacity = '0';
            const to = content ? next.scrollHeight : 0;
            if (to !== from) {
                if (!content) next.style.display = 'block';
                const result = await playOn(next, [{ height: `${from}px` }, { height: `${to}px` }], MOTION.swap);
                if (result) { result.animation.cancel(); next.style.overflow = result.overflow; }
                next.style.display = '';
            }
            if (!content) return;
            const shown = await playOn(content, [{ opacity: 0 }, { opacity: 1 }], MOTION.swap * .55, 'linear');
            if (shown) { shown.animation.cancel(); content.style.overflow = shown.overflow; }
            content.style.opacity = '';
        });
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
    // Regole del CSS (per tenerlo ordinato):
    // - colori, ombre e effetti dipendono da chiaro/scuro solo tramite le variabili di THEME:
    //   niente regole "solo per lo scuro" sui singoli elementi;
    // - una sola regola per elemento e proprietà; i gruppi comuni (sfocatura delle superfici, passaggio del mouse)
    //   sono nel blocco Superfici, in fondo;
    // - classi e variabili che non servono più si tolgono subito.
    const STYLES = `
/* L'area AMP ha l'altezza esatta del riquadro visibile: schede ferme, scorre solo il contenuto */
#ao-root { position: relative; display: flex; flex-direction: column; overflow: hidden; font-family: var(--ao-font); font-size: var(--ao-fs-base); color: var(--ao-text); background: var(--ao-backdrop); line-height: 1.4; }
#ao-root *, #ao-root *::before, #ao-root *::after { box-sizing: border-box; }
#ao-root [hidden] { display: none !important; }
#ao-root.ao-busy { cursor: progress; }
#ao-root button { font: inherit; color: inherit; }

/* Intestazione e schede */
#ao-root .ao-tabs { position: relative; flex: none; display: flex; align-items: center; gap: 2px; padding: 0 4px; background: #0078d4; color: var(--ao-nav-ink); -webkit-backdrop-filter: var(--ao-nav-blur); backdrop-filter: var(--ao-nav-blur); } /* sfondo: continua l'intestazione (blendTabsWithHeader) */
#ao-root .ao-tab { flex: 1; min-width: 0; display: flex; align-items: center; justify-content: center; gap: 0; padding: 9px 4px 7px; border: 0; border-bottom: 2px solid transparent; border-radius: var(--ao-radius-base) var(--ao-radius-base) 0 0; background: none; color: var(--ao-nav-ink-soft); font-size: var(--ao-fs-lg); font-weight: var(--ao-fw-normal); cursor: pointer; }
#ao-root .ao-tab:hover { color: var(--ao-nav-ink); }
#ao-root .ao-tab.ao-active { color: var(--ao-nav-ink); }
#ao-root .ao-tab-icon { position: relative; display: inline-flex; opacity: .8; transition: color .25s, opacity .25s; }
#ao-root .ao-tab.ao-active .ao-tab-icon { color: var(--ao-nav-ink); opacity: 1; }
/* Ogni scheda ha il suo posto: chiusa solo l'icona, aperta anche nome e contatore (si allargano nel suo posto) */
#ao-root .ao-tab-label { max-width: 0; margin-left: 0; overflow: hidden; white-space: nowrap; opacity: 0; transition: max-width var(--ao-selection-motion), margin-left var(--ao-selection-motion), opacity .25s; }
#ao-root .ao-tab.ao-active .ao-tab-label { max-width: 120px; margin-left: ${TAB_GAP.label}px; opacity: 1; } /* spazio anche per il segnalino */
@media (prefers-reduced-motion: reduce) { #ao-root .ao-tab-label { transition: none; }}
#ao-root .ao-tab-indicator { position: absolute; left: 0; bottom: 0; height: 3px; width: 0; border-radius: 3px 3px 0 0; background: var(--ao-nav-ink); pointer-events: none; opacity: 0; }
#ao-root .ao-tab-indicator.ao-ready { opacity: 1; }
#ao-root .ao-tab-indicator.ao-moving { transition: transform .32s cubic-bezier(.3, .8, .2, 1), width .32s cubic-bezier(.3, .8, .2, 1); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-tab-indicator.ao-moving { transition: none; }}
/* Scheda aperta: leggero zoom della scritta (e del numero), che resta finché è selezionata */
#ao-root .ao-tab > * { transition: transform .35s ease-in-out; }
#ao-root .ao-tab.ao-active > * { transform: scale(1.04); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-tab > * { transition: none; }}
/* Schede: il conteggio è un numero pulito, solo sulla scheda aperta. */
#ao-root .ao-counter { display: inline-block; min-width: 18px; height: 18px; padding: 0 4px; border-radius: var(--ao-radius-badge); color: var(--ao-text-muted); font-size: var(--ao-fs-md); font-weight: var(--ao-fw-medium); line-height: 18px; text-align: center; }
#ao-root .ao-tab.ao-active .ao-counter { color: var(--ao-nav-ink); }
/* Segnalino della scheda: sparisce rimpicciolendosi e svanendo, e la scritta si ricentra senza scatti (e il contrario quando torna) */
#ao-root .ao-tab .ao-counter { min-width: 0; margin-left: ${TAB_GAP.counter}px; padding: 0; max-width: 40px; color: var(--ao-nav-ink-soft); overflow: hidden; transition: opacity .2s ease, max-width .28s ease, margin-left .28s ease, transform .28s ease; } /* solo il numero: la sottolineatura si centra sulle cifre */
#ao-root .ao-tab .ao-counter.ao-counter-off { opacity: 0; max-width: 0; margin-left: 0; transform: scale(.6); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-tab .ao-counter { transition: none; }}
#ao-root .ao-body { flex: 1; min-height: 0; display: flex; flex-direction: column; }
#ao-root .ao-view { flex: 1; min-height: 0; display: flex; flex-direction: column; }
/* Barra secondaria scontornata: nessuna fascia né separatore, solo il menu.
   A casa sta in cima all'elenco; fuori casa compare sovrapposto al contenuto, con un'ombra sotto il menu stesso. */
#ao-root .ao-fixed { --ao-lift: 0; position: sticky; top: 0; z-index: 3; display: grid; gap: 10px; padding-top: 12px; pointer-events: none; will-change: transform; }
/* Completamento del movimento quando lo scorrimento si ferma a metà */
#ao-root .ao-fixed.ao-snap { transition: transform .18s ease-out; }
/* Ricomparsa: discesa morbida con un lieve rimbalzo; l'ombra si accende mentre arriva */
#ao-root .ao-fixed.ao-reveal { transition: transform .42s cubic-bezier(.2, 1.3, .4, 1); }
#ao-root .ao-fixed.ao-reveal > * { transition: box-shadow .35s ease-out .08s, background-color .35s ease-out .08s; }
#ao-root .ao-fixed > * { pointer-events: auto; }
/* Ombra graduata da --ao-lift: 0 a casa, 1 sovrapposta al contenuto */
#ao-root .ao-fixed .ao-segmented{ box-shadow: 0 3px 10px rgba(0,0,0, calc(.18 * var(--ao-lift))); transition: box-shadow .25s ease-out, background-color .25s ease-out; }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-fixed.ao-snap, #ao-root .ao-fixed.ao-reveal, #ao-root .ao-fixed.ao-reveal > * { transition: none; }}
/* Sotto la linea scorre il contenuto, a tutta larghezza (barra di scorrimento sottile, spazio riservato) */
/* Barra di scorrimento sottile (4px): lo spazio a destra è quasi uguale a quello a sinistra */
#ao-root .ao-scroll { flex: 1; min-height: 0; display: grid; align-content: start; gap: 10px; padding: 0 8px 12px 12px; scroll-padding-top: 60px; overflow-y: auto; overflow-x: hidden; scrollbar-gutter: stable; }
@supports not selector(::-webkit-scrollbar) { #ao-root .ao-scroll { scrollbar-width: thin; scrollbar-color: var(--ao-line-strong) transparent; }}
#ao-root .ao-scroll::-webkit-scrollbar { width: 4px; }
#ao-root .ao-scroll::-webkit-scrollbar-thumb { border-radius: 2px; background: var(--ao-line-strong); }
#ao-root .ao-scroll::-webkit-scrollbar-track { background: transparent; }

/* Pulsanti */
#ao-root .ao-button { display: inline-flex; align-items: center; justify-content: center; gap: 6px; min-height: var(--ao-control-height); padding: 0 12px; border: 1px solid var(--ao-line-strong); border-radius: var(--ao-radius-base); background: var(--ao-surface); font-weight: var(--ao-fw-medium); cursor: pointer; transition: background-color .3s, box-shadow .35s, filter .3s; }
#ao-root .ao-button:hover:where(:not(:disabled)) { background: var(--ao-hover); border-color: var(--ao-text-subtle); }
/* Pulsante principale, ovunque: pieno del colore della sua azione (--ao-btn, classi ao-fill-*), scritta bianca */
#ao-root .ao-button.ao-primary { --ao-btn: var(--ao-tone-take); background: color-mix(in srgb, var(--ao-btn) var(--ao-tone-fill), transparent); border-color: var(--ao-tone-line); color: var(--ao-white); box-shadow: var(--ao-btn-highlight); }
#ao-root .ao-button.ao-primary:hover:where(:not(:disabled)) { background: color-mix(in srgb, var(--ao-btn) var(--ao-tone-fill), transparent); border-color: var(--ao-tone-line); filter: var(--ao-tone-hover-filter); box-shadow: var(--ao-btn-highlight); }
#ao-root .ao-button.ao-fill-forward { --ao-btn: var(--ao-tone-forward); }
#ao-root .ao-button.ao-fill-danger { --ao-btn: var(--ao-tone-ignore); }
#ao-root :disabled { opacity: var(--ao-disabled-opacity); cursor: not-allowed; }
#ao-root :focus-visible { outline: none; box-shadow: var(--ao-shadow-focus); }

/* Filtri e scelte a segmenti */
#ao-root .ao-segmented { position: relative; display: flex; padding: 2px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-segment-track); }
/* Al passaggio del mouse la voce cambia colore; la pillola della voce scelta scivola */
#ao-root .ao-segment { position: relative; z-index: 1; flex: 1; display: inline-flex; align-items: center; justify-content: center; gap: 5px; min-height: 28px; padding: 0 6px; border: 0; border-radius: calc(var(--ao-radius-base) - 2px); background: none; color: var(--ao-text-muted); font-size: var(--ao-fs-md); cursor: pointer; white-space: nowrap; transition: color .25s; }
#ao-root .ao-segment:hover { color: var(--ao-text); }
#ao-root .ao-segment:not(.ao-active):hover { background: var(--ao-segment-hover); }
#ao-root .ao-segment.ao-active { color: var(--ao-text-accent); font-weight: var(--ao-fw-medium); }
#ao-root .ao-segment-thumb { position: absolute; top: 2px; bottom: 2px; left: 2px; width: 0; border-radius: calc(var(--ao-radius-base) - 2px); background: var(--ao-selection); box-shadow: inset 0 0 0 1px var(--ao-selection-line), var(--ao-selection-shadow); opacity: 0; pointer-events: none; }
#ao-root .ao-segment-thumb.ao-ready { opacity: 1; }
#ao-root .ao-segment-thumb.ao-moving { transition: transform var(--ao-selection-motion), width var(--ao-selection-motion); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-segment-thumb.ao-moving, #ao-root .ao-segment { transition: none; }}
#ao-root .ao-source-bar { position: relative; display: grid; gap: 6px; }
/* Riga della barra secondaria: il selettore con uno spazio vuoto ai lati (nel pannello di sviluppo lì ci sono i pulsanti) */
#ao-root .ao-bar-row { display: flex; align-items: stretch; gap: 6px; min-width: 0; }
#ao-root .ao-bar-row > .ao-segmented { flex: 1; min-width: 0; }
#ao-root .ao-bar-spacer { position: relative; flex: none; width: 34px; }
/* Selettore con icone (Attive / Ignorate, azioni): la voce scelta è più larga e ha il nome, che compare in dissolvenza */
#ao-root .ao-segmented-icons .ao-segment { flex: 1; min-width: 0; overflow: hidden; }
#ao-root .ao-segmented-icons .ao-segment.ao-grow { flex: 2.6; }
#ao-root .ao-segmented-icons.ao-growing .ao-segment { transition: color .25s, flex-grow var(--ao-selection-motion); }
/* Nome della voce: compare e svanisce allargandosi e stringendosi con la voce, così l'icona si ricentra senza scatti
   (la colonna passa da 0 alla larghezza del nome; il margine negativo annulla lo spazio tra icona e nome) */
#ao-root .ao-segment-name { display: inline-grid; grid-template-columns: 1fr; }
#ao-root .ao-segment-name > span { min-width: 0; overflow: hidden; white-space: nowrap; }
#ao-root .ao-segment-name-in { animation: ao-name-in var(--ao-selection-motion) backwards; }
#ao-root .ao-segment-name-out { animation: ao-name-out var(--ao-selection-motion) forwards; }
@keyframes ao-name-in { from { grid-template-columns: 0fr; margin-left: -5px; opacity: 0; } }
@keyframes ao-name-out { to { grid-template-columns: 0fr; margin-left: -5px; opacity: 0; } }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-segment-name-in { animation: none; } #ao-root .ao-segment-name-out { display: none; } #ao-root .ao-segmented-icons.ao-growing .ao-segment { transition: none; }}
#ao-root .ao-count { color: var(--ao-text-subtle); font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-normal); }

/* Elenco di In coda */
#ao-root .ao-list { display: grid; gap: 6px; margin: 0; padding: 0; list-style: none; }
#ao-root .ao-item { display: grid; gap: 3px; padding: 9px 10px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); box-shadow: var(--ao-shadow-card), 0 0 0 transparent, inset 0 0 0 1px transparent; cursor: pointer; transition: background-color .35s ease, box-shadow .4s ease, filter .3s, transform .32s cubic-bezier(.3, .8, .2, 1); } /* stessa struttura d'ombra del passaggio del mouse: sfuma invece di scattare */
/* Il blu (titolo e bordo) è riservato alla scheda selezionata (modulo o dettagli aperti), che non reagisce all'hover
   (il passaggio del mouse, uguale per tutte le superfici, è nel blocco Superfici). */
#ao-root .ao-expanded, #ao-root .ao-expanded:hover { background: var(--ao-surface); } /* aperta: niente schiarimento al passaggio del mouse */
#ao-root .ao-item-title { font-weight: var(--ao-fw-medium); overflow-wrap: anywhere; }
#ao-root .ao-item-sub { display: flex; align-items: center; justify-content: space-between; gap: 8px; color: var(--ao-text-muted); font-size: var(--ao-fs-md); overflow-wrap: anywhere; } /* a destra l'etichetta dell'operatore */
#ao-root .ao-item-time { color: var(--ao-text-subtle); font-size: var(--ao-fs-sm); white-space: nowrap; }
#ao-root .ao-empty{ margin: 8px 0; color: var(--ao-text-subtle); text-align: center; }

/* Riquadri */
#ao-root .ao-hint { color: var(--ao-text-subtle); font-size: var(--ao-fs-sm); }
#ao-root .ao-actions { display: flex; flex-wrap: wrap; gap: 8px; }
#ao-root .ao-actions .ao-primary { flex: 1; }

/* In coda: riga con azioni e modulo compatto */
#ao-root .ao-item-main { display: grid; grid-template-columns: minmax(0, 1fr) auto; gap: 3px 10px; cursor: pointer; border-radius: calc(var(--ao-radius-base) - 2px); }
#ao-root .ao-inline-form { display: grid; gap: 8px; padding-top: 8px; border-top: 1px solid var(--ao-line-light); }
/* Scheda di In coda aperta: testo della mail */
#ao-root .ao-mail-preview { max-height: 230px; overflow: hidden; padding: 6px 8px; border-radius: calc(var(--ao-radius-base) - 2px); background: var(--ao-quote-bg); color: var(--ao-text-muted); font-size: var(--ao-fs-md); line-height: 1.45; white-space: pre-line; overflow-wrap: anywhere; scrollbar-width: thin; } /* come la description nel thread */
#ao-root .ao-mail-preview.ao-scrollable { overflow-y: auto; } /* testo oltre 230px (scheda di sola lettura grande circa quanto il modulo di Senza mail): solo allora la barra */
/* Più messaggi nel thread: frecce per scorrerli */
#ao-root .ao-message-nav { display: flex; align-items: center; gap: 4px; min-width: 0; }
#ao-root .ao-message-step { display: inline-flex; align-items: center; justify-content: center; width: 26px; height: 26px; padding: 0; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); color: var(--ao-text); cursor: pointer; }
#ao-root .ao-message-step:hover:where(:not(:disabled)) { background: var(--ao-hover); }
#ao-root .ao-message-step:disabled { opacity: .4; cursor: default; }
#ao-root .ao-message-where { color: var(--ao-text-muted); font-size: var(--ao-fs-sm); font-variant-numeric: tabular-nums; white-space: nowrap; }
#ao-root .ao-queue-item .ao-item-main { display: grid; grid-template-columns: minmax(0, 1fr); gap: 3px; }
#ao-root .ao-queue-row { display: flex; align-items: center; justify-content: space-between; gap: 8px; min-width: 0; }
#ao-root .ao-queue-row > :first-child { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
#ao-root .ao-queue-row > :last-child { flex: none; }
#ao-root .ao-queue-item.ao-expanded .ao-queue-row .ao-item-title { white-space: normal; overflow-wrap: anywhere; }
#ao-root .ao-queue-info { margin: 0; }
/* Prendi in carico: su una riga e largo, occupa lo spazio libero accanto a Ignora e Chiudi */
#ao-root .ao-queue-item .ao-actions { flex-wrap: nowrap; }
#ao-root .ao-queue-item .ao-actions .ao-primary { flex: 1; white-space: nowrap; }
/* Scheda aperta: selettore dell'azione (e del gruppo), contenuto dell'azione, freccia per richiudere */
#ao-root .ao-choice { position: relative; display: flex; padding: 3px; border-radius: var(--ao-radius-base); background: var(--ao-choice-track); }
#ao-root .ao-choice-option { position: relative; z-index: 1; flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center; min-height: 30px; padding: 4px 6px; border: 0; border-radius: calc(var(--ao-radius-base) - 2px); background: none; color: var(--ao-text-muted); font-size: var(--ao-fs-md); font-weight: var(--ao-fw-medium); cursor: pointer; white-space: nowrap; line-height: 1.25; transition: color .25s; }
#ao-root .ao-choice-sub { font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-normal); }
#ao-root .ao-choice-option:hover:not(.ao-active) { color: var(--ao-text); }
#ao-root .ao-choice-option:hover:not(.ao-active):not(:disabled) { background: var(--ao-segment-hover); }
#ao-root .ao-choice-thumb { position: absolute; top: 3px; bottom: 3px; left: 3px; width: 0; border-radius: calc(var(--ao-radius-base) - 2px); background: color-mix(in srgb, var(--ao-choice-bg) var(--ao-selection-tone), var(--ao-selection)); box-shadow: inset 0 0 0 1px var(--ao-selection-line), var(--ao-selection-shadow); opacity: 0; pointer-events: none; transition: background-color .25s; }
#ao-root .ao-choice-thumb.ao-ready { opacity: 1; }
#ao-root .ao-choice-thumb.ao-appearing { animation: ao-fade-in .2s linear; }
#ao-root .ao-choice-thumb.ao-moving { transition: transform var(--ao-selection-motion), width var(--ao-selection-motion), background-color .25s; }
@keyframes ao-fade-in { from { opacity: 0; } }
#ao-root .ao-choice-accent { --ao-choice-bg: color-mix(in srgb, var(--ao-accent) 16%, var(--ao-surface)); --ao-choice-glow: var(--ao-accent); --ao-choice-text: var(--ao-text-accent); } /* più deciso dell'azzurro dei filtri: si stacca dal fondo grigio */
#ao-root .ao-choice-forward { --ao-choice-bg: var(--ao-choice-forward-bg); --ao-choice-glow: var(--ao-tone-forward); --ao-choice-text: var(--ao-choice-forward-text); }
#ao-root .ao-choice-danger { --ao-choice-bg: var(--ao-choice-danger-bg); --ao-choice-glow: var(--ao-tone-ignore); --ao-choice-text: var(--ao-choice-danger-text); }
/* Selettori delle azioni: la voce scelta prende il colore della sua azione */
#ao-root .ao-segmented.ao-toned .ao-segment-thumb { background: color-mix(in srgb, var(--ao-choice-bg) var(--ao-selection-tone), var(--ao-selection)); }
#ao-root .ao-segmented.ao-toned .ao-segment.ao-active { color: var(--ao-choice-text); }
#ao-root .ao-choice-option.ao-active { color: var(--ao-choice-text); }
#ao-root .ao-choice-zone:empty { display: none; }
#ao-root .ao-choice-content { display: grid; gap: 8px; }
#ao-root .ao-collapse { position: relative; top: 5px; /* più distante dai pulsanti, dentro il margine della scheda: l'altezza non cambia */ display: flex; justify-content: center; align-items: center; justify-self: center; width: 44px; height: 22px; margin-top: -2px; padding: 0; border: 0; border-radius: var(--ao-radius-base); background: none; color: var(--ao-text-muted); cursor: pointer; }
#ao-root .ao-collapse:hover { background: var(--ao-hover); color: var(--ao-text); }
#ao-root .ao-svg-icon { display: inline-flex; }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-choice-thumb, #ao-root .ao-choice-option { transition: none; animation: none; }}
/* Esito su una scheda di In coda: banner colorato (le animazioni sono nel blocco Animazioni dello script) */
#ao-root .ao-queue-item { position: relative; overflow: hidden; gap: 6px; transition: background-color .3s, border-color .3s, transform .32s cubic-bezier(.3, .8, .2, 1), box-shadow .32s ease, filter .32s ease; }
/* Scheda aperta: in primo piano (si solleva e si ingrandisce appena), bordo chiaro come le altre */
#ao-root .ao-queue-item.ao-expanded, #ao-root .ao-queue-item.ao-expanded:hover {
    z-index: 1; border-color: var(--ao-line); transform: translateY(-2px) scale(1.015);
    box-shadow: var(--ao-open-shadow), inset 0 0 0 1px var(--ao-open-edge);
}
#ao-root .ao-queue-item.ao-resolving { transform: none; box-shadow: var(--ao-shadow-card); }
/* Con una scheda aperta le altre si sfocano (tornano nitide appena si richiude) */
#ao-root .ao-list:has(.ao-queue-item.ao-expanded:not(.ao-unlifted):not(.ao-resolving)) .ao-queue-item:not(.ao-expanded) { filter: blur(5px); }
/* Titolo: neutro finché non si sceglie l'azione, poi del suo colore */
#ao-root .ao-queue-item.ao-expanded .ao-item-title { color: var(--ao-text); transition: color .25s; }
#ao-root .ao-queue-item.ao-expanded[data-choice="take"] .ao-item-title { color: var(--ao-text-accent); }
#ao-root .ao-queue-item.ao-expanded[data-choice="forward"] .ao-item-title { color: var(--ao-choice-forward-text); }
#ao-root .ao-queue-item.ao-expanded[data-choice="ignore"] .ao-item-title { color: var(--ao-choice-danger-text); }
#ao-root .ao-resolving { pointer-events: none; }
/* Colore dell'esito come quello dei pulsanti, scritta bianca */
#ao-root .ao-outcome-accent { --ao-outcome: var(--ao-tone-take); }
#ao-root .ao-outcome-danger { --ao-outcome: var(--ao-tone-ignore); }
#ao-root .ao-outcome-forward { --ao-outcome: var(--ao-tone-forward); }
/* sfondo e bordo della scheda prendono il colore solo a tendina chiusa (prima lo dà solo il banner, che affiora) */
#ao-root .ao-queue-item.ao-outcome-full { background: color-mix(in srgb, var(--ao-outcome) var(--ao-tone-solid), var(--ao-surface)); border-color: color-mix(in srgb, var(--ao-outcome) var(--ao-tone-solid), var(--ao-line)); }
#ao-root .ao-outcome {
    position: absolute; inset: 0; z-index: 2; display: flex; align-items: center; gap: 8px; padding: 0 12px;
    background: color-mix(in srgb, var(--ao-outcome) var(--ao-tone-fill), transparent); color: #fff;
    font-size: var(--ao-fs-md); font-weight: var(--ao-fw-medium);
}
#ao-root .ao-outcome-spinner { width: 14px; height: 14px; border: 2px solid color-mix(in srgb, currentColor var(--ao-spinner-track), transparent); border-top-color: currentColor; border-radius: 50%; animation: ao-spin .8s linear infinite; }
@keyframes ao-spin { to { transform: rotate(360deg); } }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-queue-item, #ao-root .ao-outcome { transition: none; animation: none; }}

/* Appena aperta, ancora a terra: lo stesso aspetto del passaggio del mouse (bagliore e bordo), così sollevandosi il
   bagliore prosegue senza spegnersi; richiudendo torna allo stesso aspetto */
#ao-root .ao-queue-item.ao-expanded.ao-unlifted { transform: none; box-shadow: var(--ao-item-hover-glow), var(--ao-shadow-card), inset 0 0 0 1px var(--ao-item-hover-edge); }

/* Superfici: sfocatura di quello che sta dietro e passaggio del mouse, uguale per tutte */
#ao-root .ao-item, #ao-root .ao-segmented{ -webkit-backdrop-filter: var(--ao-blur); backdrop-filter: var(--ao-blur); }
#ao-root .ao-item:not(.ao-expanded):hover { background: var(--ao-item-hover); box-shadow: var(--ao-shadow-card), var(--ao-item-hover-glow), inset 0 0 0 1px var(--ao-item-hover-edge); } /* solo la scheda: niente filtro sul contenuto (timer, etichette) */

/* Messaggi temporanei, in basso */
#ao-root .ao-status:not([hidden]) { display: flex; align-items: center; justify-content: space-between; gap: 10px; }
#ao-root .ao-status-action { flex: none; padding: 0; border: 0; background: none; color: var(--ao-text-accent); font: inherit; font-weight: var(--ao-fw-bold); cursor: pointer; }
#ao-root .ao-status-action:hover { text-decoration: underline; }
#ao-root .ao-status { position: absolute; left: 12px; right: 12px; bottom: 12px; z-index: 3; margin: 0; padding: 8px 10px; border-radius: var(--ao-radius-base); font-size: var(--ao-fs-md); font-weight: var(--ao-fw-medium); box-shadow: 0 2px 8px rgba(0,0,0,.12); }
#ao-root .ao-status.ao-tone-success { background: var(--ao-success-bg); color: var(--ao-success); }
#ao-root .ao-status.ao-tone-warning { background: var(--ao-warning-bg); color: var(--ao-warning); }
#ao-root .ao-status.ao-tone-danger { background: var(--ao-danger-bg); color: var(--ao-danger); }
`;
    // Fuori dall'area AMP, solo mentre è aperta: intestazione dell'add-in più bassa, con le schede più vicine al logo.
    // Lo sfondo sfumato del pannello (fisso, quindi senza giunture) passa anche dietro l'intestazione, che diventa di vetro.
    const headerRule = (look, ink) => `background: ${look.background}; box-shadow: ${look.light}; color: ${ink};`;
    const OUTSIDE_STYLES = `
:has(> .app-header):has(#amp-view:not([hidden])) { background: ${THEME.pageBackdrop}; }
:has(> .app-header):has(#amp-view:not([hidden]) #ao-root[data-theme="dark"]) { background: ${THEME.pageDarkBackdrop}; }
@media (prefers-color-scheme: dark) { :has(> .app-header):has(#amp-view:not([hidden]) #ao-root:not([data-theme="light"])) { background: ${THEME.pageDarkBackdrop}; } }
:has(> .app-header):has(#amp-view:not([hidden])) > .app-header { padding-bottom: 6px; -webkit-backdrop-filter: ${THEME.colors.navBlur}; backdrop-filter: ${THEME.colors.navBlur}; ${headerRule(THEME.header.light, THEME.colors.navInk)} }
:has(> .app-header):has(#amp-view:not([hidden]) #ao-root[data-theme="dark"]) > .app-header { ${headerRule(THEME.header.dark, THEME.darkColors.navInk)} }
@media (prefers-color-scheme: dark) { :has(> .app-header):has(#amp-view:not([hidden]) #ao-root:not([data-theme="light"])) > .app-header { ${headerRule(THEME.header.dark, THEME.darkColors.navInk)} } }
`;
    function injectStyles() {
        if (document.getElementById('ao-styles')) return;
        document.head.append(h('style', { id: 'ao-styles' }, themeCssVariables() + STYLES + OUTSIDE_STYLES));
    }

    // ======================================================================
    // 5. AVVIO E COLLEGAMENTO CON L'ADD-IN
    // ======================================================================

    let initialized = false;
    // Asincrona come quella originale dell'area AMP: taskpane.js chiama
    // AmpModule.initialize().catch(...) all'avvio dell'add-in.
    async function initialize() {
        if (initialized) return;
        // Nella pagina di jak l'area AMP può mancare: il pannello la crea in cima a <main>.
        let container = document.getElementById('amp-view');
        if (!container) {
            const host = document.querySelector('main') || document.body;
            container = h('section', { class: 'amp-view', id: 'amp-view' });
            host.prepend(container);
        }
        container.hidden = false;
        initialized = true;
        injectStyles();
        buildShell(container);
        fillPane();
        // Il riquadro di Outlook si può ridimensionare: l'altezza dell'area si ricalcola.
        window.addEventListener('resize', fitHeight);
        if (window.ResizeObserver) new ResizeObserver(() => fitHeight()).observe(scrollContainer());
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
        const ampView = document.getElementById('amp-view');
        if (ampView) ampView.hidden = false;
        render();
        fitHeight();
    }

    window.AmpModule = Object.freeze({ initialize, show });
    const logInitError = error => console.error('AMP init error:', error);
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', () => initialize().catch(logInitError));
    else initialize().catch(logInitError);
})();
