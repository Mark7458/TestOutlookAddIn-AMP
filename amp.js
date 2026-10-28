/*
 * AMP - SNOW Assistant: area AMP dell'add-in Outlook.
 *
 * Nell'add-in i ticket si gestiscono da Outlook, senza aprire ServiceNow:
 *   - Coda         mail arrivate nella casella condivisa che non hanno ancora un ticket;
 *   - Ticket       ticket aperti su ServiceNow, con chi li gestisce;
 *   - Mail         la mail aperta in Outlook: presa in carico, apertura del ticket, tempo,
 *                  storico di ServiceNow e azione da eseguire all'invio della risposta;
 *   - Impostazioni profilo, Agenzie (con i domini email), Attività, risposte salvate.
 *
 * Collegamento con l'add-in di jak (nessun suo file viene modificato):
 *   - espone window.AmpModule { initialize, show }; "Apri area AMP" chiama show();
 *   - la freccia in alto torna alla schermata principale (TestPluginAPI.showMainView);
 *   - legge la mail aperta da TestPluginAPI.getCurrentMessageContext / onMessageChanged.
 *
 * Tutto ciò che sta fuori dal pannello (casella condivisa, SharePoint, ServiceNow) passa
 * dal blocco 3 (Servizi). Oggi i servizi esistono solo nell'anteprima (amp-preview-backend.js,
 * simulati); nell'add-in vanno collegati alla Function App.
 */
(function () {
    'use strict';

    // Indice dei blocchi
    //  0. Aspetto: tema, testi                 5. Interfaccia: struttura e navigazione
    //  1. Valori di ServiceNow                 6. Viste: Coda, Ticket, Mail, Impostazioni
    //  2. Configurazioni (salvate in locale)   7. Stili (usano i valori del blocco 0)
    //  3. Servizi e Outlook                    8. Avvio e collegamento con l'add-in
    //  4. Regole: agenzia suggerita, date, tempo

    // ======================================================================
    // 0. ASPETTO: TEMA E TESTI
    // ======================================================================

    // Tutti i valori di aspetto: il CSS li usa come variabili --ao-* (vedi themeCssVariables()).
    const THEME = {
        colors: {
            accent: '#0078d4', // pulsanti principali, scheda attiva, voci selezionate
            accentDark: '#106ebe', // colore principale al passaggio del mouse
            accentBorder: '#c7e0f4',
            accentBg: '#eff6fc',
            headerBackground: 'linear-gradient(135deg, #005a9e, #0078d4)',
            white: '#fff',
            text: '#242424',
            textMuted: '#605e5c',
            textSubtle: '#8a8886',
            bg: '#f5f5f5', // sfondo dell'area (come il riquadro dell'add-in)
            bgSoft: '#faf9f8',
            lineLight: '#edebe9',
            line: '#e1dfdd',
            lineStrong: '#c8c6c4',
            success: '#107c10',
            successBg: '#dff6dd',
            warning: '#8a5a00',
            warningBg: '#fff4ce',
            danger: '#a4262c',
            dangerBg: '#fde7e9',
            hold: '#8764b8', // On Hold
            holdBg: '#f4f0fa'
        },
        font: { family: '"Segoe UI", system-ui, -apple-system, BlinkMacSystemFont, sans-serif' },
        fontSize: { xs: '10px', sm: '11px', md: '12px', base: '13px', lg: '14px', xl: '16px' },
        fontWeight: { normal: '400', medium: '600', bold: '700' },
        radius: { base: '6px', pill: '999px' },
        shadow: {
            card: '0 1px 2px rgba(0,0,0,.06)',
            focus: '0 0 0 2px rgba(0,120,212,.25)'
        },
        state: { disabledOpacity: '.45' },
        layout: { controlHeight: '32px' }
    };

    // Tutte le scritte del pannello. Le voci con (valore) sono funzioni.
    const TEXTS = {
        title: 'AMP',
        back: 'Torna all’add-in',
        tabs: { queue: 'Coda', tickets: 'Ticket', mail: 'Mail', settings: 'Impostazioni' },
        common: {
            me: 'me',
            loading: 'Caricamento…',
            cancel: 'Annulla',
            save: 'Salva',
            edit: 'Modifica',
            delete: 'Elimina',
            confirm: 'Confermi?',
            add: 'Aggiungi',
            none: '—',
            unassigned: 'Libera',
            servicesMissing: 'I servizi non sono ancora collegati: qui compariranno i dati della casella condivisa e di ServiceNow.',
            serviceError: message => `Servizio non disponibile: ${message}`
        },
        queue: {
            all: 'Tutte',
            mine: 'Mie',
            free: 'Libere',
            empty: 'Nessuna richiesta in coda.',
            take: 'Prendi in carico',
            discarded: n => `${n} scartate`,
            suggested: name => `Agenzia: ${name}`
        },
        tickets: {
            mine: 'Miei',
            all: 'Tutti',
            empty: 'Nessun ticket aperto.',
            holdExpired: 'On Hold scaduto',
            holdUntil: date => `fino al ${date}`
        },
        mail: {
            noMail: 'Apri una mail in Outlook, oppure scegli una richiesta dalla Coda o un ticket.',
            notRegistered: 'Questa mail non è tra le richieste della casella condivisa.',
            register: 'Aggiungi alla coda',
            openInOutlook: 'Apri la mail',
            takeCharge: 'Prendi in carico',
            takeOver: owner => `Prendi da ${owner}`,
            release: 'Rilascia',
            inCharge: owner => `In carico a ${owner}`,
            notInCharge: 'Nessuno l’ha presa in carico',
            discardedBy: (who, when) => `Scartata da ${who} · ${when}`,
            restore: 'Rimetti in coda',
            // Apertura del ticket
            newTicket: 'Nuovo ticket',
            company: 'Agenzia',
            activity: 'Attività',
            choose: 'Seleziona…',
            suggestedFrom: reason => `Suggerita ${reason}`,
            byDomain: domain => `dal dominio @${domain}`,
            byText: 'dal testo della mail',
            favorites: 'Preferiti',
            shortDescription: 'Short description',
            fieldsPreview: 'Su ServiceNow',
            openTicket: 'Apri ticket',
            discard: 'Non serve un ticket',
            profileMissing: 'Imposta prima il tuo profilo (Impostazioni → Profilo).',
            selectBoth: 'Scegli Agenzia e Attività.',
            ticketOpened: number => `Ticket ${number} aperto.`,
            // Ticket
            ticket: 'Ticket',
            openedBy: (who, when) => `Aperto da ${who} · ${when}`,
            openInServiceNow: 'Apri in ServiceNow',
            timeWorked: 'Tempo lavorato',
            minutes: n => `${n} min`,
            timeChanged: (delta, total) => `Tempo ${delta > 0 ? '+' : ''}${delta} min · totale ${total} min.`,
            closed: state => `Ticket ${state}: nessuna azione disponibile.`,
            // Azione all'invio
            onSend: 'Quando invii la risposta',
            onSendHint: 'Il testo della mail diventa il commento su ServiceNow; il ticket si aggiorna come scegli qui.',
            actions: { comment: 'Solo risposta', resolve: 'Risolvi', hold: 'On Hold', cancel: 'Annulla ticket' },
            resolutionCode: 'Resolution code',
            holdReason: 'On hold reason',
            holdExpiration: 'Scadenza',
            workNotes: 'Nota interna (facoltativa)',
            replyText: 'Testo di partenza',
            replyEmpty: 'Vuoto',
            replyCombination: 'Risposta di Agenzia + Attività',
            prepareReply: 'Salva e rispondi',
            pending: label => `All’invio: ${label}`,
            pendingNone: 'All’invio: solo commento',
            clearPending: 'Togli',
            pendingSaved: 'Azione salvata: verrà eseguita all’invio della risposta.',
            replyOpened: 'Risposta aperta in Outlook.',
            replyUnavailable: 'Apri la mail per rispondere.',
            // Storico
            history: 'Storico ServiceNow',
            loadHistory: 'Carica',
            loadMore: n => `Mostra precedenti (${n})`,
            historyEmpty: 'Nessuna attività.',
            kinds: { comment: 'Commento', work_note: 'Nota interna', system: 'Sistema' }
        },
        settings: {
            profile: 'Profilo',
            tag: 'Tag in Outlook',
            tagHint: 'Il tag che metti sulle mail della casella condivisa (es. Marco).',
            assignedTo: 'Assigned to (ServiceNow)',
            displayValue: 'Nome',
            sysId: 'sys_id',
            companies: 'Agenzie',
            activities: 'Attività',
            responses: 'Risposte salvate',
            data: 'Dati',
            companyName: 'Nome',
            company: 'Company',
            assignmentGroup: 'Assignment group',
            domains: 'Domini email',
            domainsHint: 'Separati da virgola, es. agenziaalfa.it, alfa.it',
            category: 'Category',
            impact: 'Impact',
            urgency: 'Urgency',
            timeWorked: 'Minuti',
            prefix: 'Prefisso Short description',
            title: 'Titolo',
            response: 'Testo',
            empty: 'Nessuna voce.',
            import: 'Importa JSON',
            export: 'Esporta JSON',
            imported: 'Configurazioni importate.',
            importInvalid: 'File JSON non valido.',
            saved: 'Salvato.',
            deleted: 'Eliminato.',
            nameRequired: 'Il nome è obbligatorio.'
        }
    };

    // ======================================================================
    // 1. VALORI DI SERVICENOW
    // ======================================================================

    const TICKET_STATES = {
        '1': { label: 'New', tone: 'neutral' },
        '2': { label: 'In Progress', tone: 'accent' },
        '3': { label: 'On Hold', tone: 'hold' },
        '6': { label: 'Resolved', tone: 'success' },
        '7': { label: 'Closed', tone: 'neutral' },
        '8': { label: 'Canceled', tone: 'neutral' }
    };
    const OPEN_STATES = ['2', '3'];
    const RESOLUTION_CODES = [
        'Solved (Permanently)', 'Solved by Third Party Vendor', 'Closed on Initial', 'Not Reproducible (Alert Cleared)',
        'Not Solved (Not Reproducible)', 'Caused by change', 'Not Solved (Too Costly)', 'No Client Response',
        'Solved (Work Around)', 'Duplicate', 'Out of Scope'
    ];
    const HOLD_REASONS = [
        { value: '7', label: 'Awaiting Client Action/Update' },
        { value: '8', label: 'Awaiting Client Resolution' },
        { value: '9', label: 'Awaiting Review' },
        { value: '6', label: 'Awaiting Third Party' },
        { value: '5', label: 'Awaiting Change' },
        { value: '3', label: 'Awaiting Problem Resolution' }
    ];
    const SEVERITY_LABELS = { '1': '1 - High', '2': '2 - Medium', '3': '3 - Low' };
    const TIME_STEPS = [-15, -5, 5, 15]; // pulsanti del tempo lavorato, in minuti
    const DEFAULT_HOLD_DAYS = 2; // scadenza proposta per l'On Hold
    const HISTORY_PAGE = 5; // voci dello storico caricate per volta

    // ======================================================================
    // 2. CONFIGURAZIONI (salvate nel browser, stesso formato dello script Tampermonkey)
    // Un file esportato da Tampermonkey si importa qui senza modifiche; in più le Agenzie
    // hanno i domini email e il profilo ha il tag di Outlook.
    // ======================================================================

    const STORAGE = { key: 'snow_ticket_assistant_v1', version: 3 };

    function createDefaultConfig() {
        const now = new Date().toISOString();
        return {
            version: STORAGE.version, created: now, updated: now,
            profile: { assignedTo: { value: '', displayValue: '' }, tag: '' },
            companies: {}, activities: {}, favorites: [], recent: [],
            combinationResponses: {}, usefulResponses: {},
            preferences: { favoritesGroupBy: 'none' }
        };
    }
    function normalizeReference(reference) {
        if (!reference) return { value: '', displayValue: '' };
        if (typeof reference === 'string') return { value: '', displayValue: reference };
        return { value: reference.value || '', displayValue: reference.displayValue || '' };
    }
    function normalizeDomains(value) {
        const list = Array.isArray(value) ? value : String(value || '').split(/[,;\s]+/);
        return [...new Set(list.map(item => String(item).trim().toLowerCase().replace(/^@/, '')).filter(Boolean))];
    }
    function normalizeConfig(data) {
        const defaults = createDefaultConfig();
        if (!data || typeof data !== 'object') return defaults;
        const isObject = value => value && typeof value === 'object' && !Array.isArray(value);
        const companies = isObject(data.companies) ? data.companies : {};
        Object.values(companies).forEach(company => { company.domains = normalizeDomains(company.domains); });
        return {
            ...defaults,
            ...data,
            version: Number(data.version) || STORAGE.version,
            profile: { assignedTo: normalizeReference(data.profile?.assignedTo), tag: String(data.profile?.tag || '') },
            companies,
            activities: isObject(data.activities) ? data.activities : {},
            favorites: Array.isArray(data.favorites) ? data.favorites : [],
            recent: Array.isArray(data.recent) ? data.recent : [],
            combinationResponses: isObject(data.combinationResponses) ? data.combinationResponses : {},
            usefulResponses: isObject(data.usefulResponses) ? data.usefulResponses : {}
        };
    }
    function loadConfig() {
        try { return normalizeConfig(JSON.parse(localStorage.getItem(STORAGE.key) || 'null')); }
        catch (error) {
            console.error('[AMP] Lettura configurazioni non riuscita:', error);
            return createDefaultConfig();
        }
    }
    function saveConfig(config) {
        config.version = STORAGE.version;
        config.updated = new Date().toISOString();
        try { localStorage.setItem(STORAGE.key, JSON.stringify(config)); }
        catch (error) { console.error('[AMP] Salvataggio configurazioni non riuscito:', error); }
    }
    function createId(label) {
        const base = String(label || '').normalize('NFD').replace(/[̀-ͯ]/g, '')
            .toLowerCase().trim().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
        return `${base || 'voce'}_${Date.now().toString(36)}`;
    }
    const combinationKey = (companyId, activityId) => `${companyId}::${activityId}`;
    const byName = (a, b) => String(a[1].name || a[1].title || '').localeCompare(String(b[1].name || b[1].title || ''), 'it');

    // ======================================================================
    // 3. SERVIZI E OUTLOOK
    // Unico punto in cui il pannello parla con l'esterno.
    //
    // Servizi (Function App): ogni metodo restituisce una Promise.
    //   listRequests()                          → Richiesta[]
    //   getRequestByConversation(conversationId) → Richiesta | null
    //   registerRequest(mail)                    → Richiesta   (mail non ancora in coda)
    //   takeCharge(id, tag)                      → Richiesta   (tag '' = rilascia; mette la categoria in Outlook)
    //   discardRequest(id, tag) / restoreRequest(id) → Richiesta
    //   createTicket(id, { companyId, companyName, activityId, activityName, openedBy, fields }) → Richiesta
    //   addTime(id, minuti, tag)                 → Richiesta   (minuti anche negativi)
    //   getTicketActivity(id, { offset, limit }) → { entries: [{ at, author, kind, text }], total }
    //   setPendingAction(id, azione | null)      → Richiesta   (eseguita dal server all'invio della risposta)
    //   subscribe(listener)                      → funzione per smettere (facoltativo)
    //
    // Richiesta (una riga della lista SharePoint, una per conversazione):
    //   { id, conversationId, messageId, subject, sender: { name, email }, receivedAt, preview,
    //     owner, status: 'queue' | 'discarded' | 'ticket', discardedBy, discardedAt,
    //     ticket: { number, sysId, url, state, companyId, companyName, activityId, activityName,
    //               shortDescription, timeWorkedMinutes, holdReason, holdExpiration, openedBy, openedAt } | null,
    //     pendingAction: { type: 'comment' | 'resolve' | 'cancel' | 'hold', resolutionCode,
    //                      holdReason, holdExpiration, workNotes, by } | null }
    // ======================================================================

    const SERVICE_METHODS = ['listRequests', 'getRequestByConversation', 'registerRequest', 'takeCharge',
        'discardRequest', 'restoreRequest', 'createTicket', 'addTime', 'getTicketActivity', 'setPendingAction'];

    class ServicesMissingError extends Error {}

    // Oggi: window.AmpBackend (simulato, solo nell'anteprima). Domani: le chiamate alla Function App.
    const Services = Object.fromEntries(SERVICE_METHODS.map(name => [name, async (...args) => {
        const backend = window.AmpBackend;
        if (typeof backend?.[name] !== 'function') throw new ServicesMissingError(name);
        return backend[name](...args);
    }]));
    Services.available = () => SERVICE_METHODS.every(name => typeof window.AmpBackend?.[name] === 'function');
    Services.subscribe = listener => window.AmpBackend?.subscribe?.(listener) || (() => {});

    // Outlook: mail aperta (dall'add-in di jak) e apertura di mail e risposte.
    const Outlook = {
        // Mail aperta, nel formato del pannello.
        currentMail() {
            const context = window.TestPluginAPI?.getCurrentMessageContext?.();
            if (!context) return null;
            return {
                itemId: context.itemId || null,
                conversationId: context.normalizedConversationId || context.conversationId || null,
                subject: String(context.subject || ''),
                sender: { name: context.sender?.displayName || '', email: context.sender?.emailAddress || '' },
                receivedAt: context.dateTimeCreated || null,
                body: String(context.bodyText || ''),
                userName: context.mailbox?.userDisplayName || ''
            };
        },
        onMailChanged(listener) {
            return window.TestPluginAPI?.onMessageChanged?.(listener) || (() => {});
        },
        // Apre una mail della casella (anteprima: la seleziona; Outlook: la apre in una finestra).
        openMessage(messageId) {
            if (window.AmpPreviewOutlook) return window.AmpPreviewOutlook.openMessage(messageId);
            const mailbox = window.Office?.context?.mailbox;
            if (messageId && typeof mailbox?.displayMessageForm === 'function') {
                mailbox.displayMessageForm(messageId);
                return true;
            }
            return false;
        },
        // Apre la risposta alla mail aperta con un testo di partenza.
        openReply(text) {
            if (window.AmpPreviewOutlook) return window.AmpPreviewOutlook.openReply(text);
            const item = window.Office?.context?.mailbox?.item;
            if (typeof item?.displayReplyForm !== 'function') return false;
            item.displayReplyForm({ htmlBody: escapeHtml(text).replace(/\n/g, '<br>') });
            return true;
        }
    };

    // ======================================================================
    // 4. REGOLE: AGENZIA SUGGERITA, DATE, TEMPO
    // ======================================================================

    // Agenzia della mail: prima il dominio del mittente, poi il nome dell'Agenzia nell'oggetto o nel corpo.
    function suggestCompany(config, sender, subject, body) {
        const domain = String(sender?.email || '').split('@')[1]?.toLowerCase() || '';
        const companies = Object.entries(config.companies);
        if (domain) {
            const match = companies.find(([, company]) => company.domains.some(item => domain === item || domain.endsWith(`.${item}`)));
            if (match) return { id: match[0], reason: TEXTS.mail.byDomain(domain) };
        }
        const text = `${subject || ''}\n${body || ''}`.toLowerCase();
        const byText = companies.find(([, company]) => company.name && text.includes(company.name.toLowerCase()));
        return byText ? { id: byText[0], reason: TEXTS.mail.byText } : null;
    }

    const pad = value => String(value).padStart(2, '0');
    function parseDate(value) {
        if (!value) return null;
        const text = String(value);
        // Formato di ServiceNow "AAAA-MM-GG hh:mm:ss" (ora locale) oppure ISO.
        const date = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}(:\d{2})?$/.test(text) ? new Date(text.replace(' ', 'T')) : new Date(text);
        return Number.isNaN(date.getTime()) ? null : date;
    }
    function formatServiceNowDate(date) {
        return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
    }
    function toDateTimeInput(date) {
        return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
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
    function defaultHoldExpiration() {
        const date = new Date(Date.now() + DEFAULT_HOLD_DAYS * 86400000);
        date.setHours(9, 0, 0, 0);
        return toDateTimeInput(date);
    }
    function isHoldExpired(ticket) {
        const date = ticket?.state === '3' ? parseDate(ticket.holdExpiration) : null;
        return Boolean(date && date < new Date());
    }
    function escapeHtml(value) {
        return String(value ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
    }

    // ======================================================================
    // 5. INTERFACCIA: STRUTTURA E NAVIGAZIONE
    // ======================================================================

    const state = {
        view: 'queue', // queue | tickets | mail | settings
        requests: [],
        loaded: false,
        loading: false,
        serviceError: '',
        queueFilter: 'all',
        ticketFilter: 'mine',
        mail: null, // mail aperta in Outlook
        focusId: null, // richiesta mostrata nella vista Mail
        mailLookup: null, // richiesta trovata per la mail aperta (anche se non in elenco)
        ticketDrafts: {}, // per richiesta: Agenzia, Attività, Short description
        actionDrafts: {}, // per richiesta: azione all'invio
        history: {}, // per richiesta: storico caricato
        editing: null, // impostazioni: { kind, id }
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
            else if (value === true) element.setAttribute(key, '');
            else element.setAttribute(key, value);
        });
        children.flat(Infinity).forEach(child => {
            if (child === null || child === undefined || child === false) return;
            element.append(child instanceof Node ? child : document.createTextNode(String(child)));
        });
        return element;
    }
    const chip = (text, tone = 'neutral', title) => h('span', { class: `ao-chip ao-tone-${tone}`, title }, text);
    const stateChip = ticket => {
        const info = TICKET_STATES[ticket.state] || { label: ticket.state, tone: 'neutral' };
        return isHoldExpired(ticket) ? chip(TEXTS.tickets.holdExpired, 'danger') : chip(info.label, info.tone);
    };
    function selectField(label, options, value, onchange, extra = {}) {
        return h('label', { class: 'ao-field' },
            h('span', { class: 'ao-label' }, label),
            h('select', { class: 'ao-input', onchange: event => onchange(event.target.value), ...extra },
                options.map(option => {
                    const element = h('option', { value: option.value }, option.label);
                    element.selected = String(option.value) === String(value ?? '');
                    return element;
                })));
    }
    function textField(label, value, oninput, extra = {}) {
        const { multiline, hint, ...attrs } = extra;
        return h('label', { class: 'ao-field' },
            h('span', { class: 'ao-label' }, label),
            h(multiline ? 'textarea' : 'input', { class: 'ao-input', value: value ?? '', oninput: event => oninput(event.target.value), ...attrs }),
            hint ? h('span', { class: 'ao-hint' }, hint) : null);
    }
    // Pulsante che chiede conferma con un secondo clic (confirm() non esiste negli add-in Office).
    function confirmButton(label, className, action) {
        const button = h('button', { type: 'button', class: className }, label);
        let armed = null;
        button.addEventListener('click', () => {
            if (armed) {
                clearTimeout(armed);
                armed = null;
                action();
                return;
            }
            button.textContent = TEXTS.common.confirm;
            button.classList.add('ao-armed');
            armed = setTimeout(() => {
                armed = null;
                button.textContent = label;
                button.classList.remove('ao-armed');
            }, 4000);
        });
        return button;
    }

    function myTag(config = loadConfig()) {
        if (config.profile.tag) return config.profile.tag;
        const name = state.mail?.userName || window.TestPluginAPI?.getCurrentMessageContext?.()?.mailbox?.userDisplayName || '';
        return name.split(/\s+/)[0] || '';
    }
    function showStatus(message, tone = 'success') {
        const status = root?.querySelector('.ao-status');
        if (!status) return;
        status.textContent = message;
        status.className = `ao-status ao-tone-${tone}`;
        status.hidden = false;
        clearTimeout(statusTimer);
        statusTimer = setTimeout(() => { status.hidden = true; }, tone === 'danger' ? 7000 : 4000);
    }
    // Esegue una chiamata ai servizi mostrando l'esito; restituisce il risultato o null.
    async function run(task, successMessage) {
        if (state.busy) return null;
        state.busy = true;
        root?.classList.add('ao-busy');
        try {
            const result = await task();
            if (successMessage) showStatus(typeof successMessage === 'function' ? successMessage(result) : successMessage);
            return result;
        } catch (error) {
            console.error('[AMP]', error);
            showStatus(error instanceof ServicesMissingError ? TEXTS.common.servicesMissing : (error?.message || String(error)), 'danger');
            return null;
        } finally {
            state.busy = false;
            root?.classList.remove('ao-busy');
        }
    }
    function upsertRequest(request) {
        if (!request) return;
        const index = state.requests.findIndex(item => item.id === request.id);
        if (index >= 0) state.requests[index] = request;
        else state.requests.push(request);
        if (state.mailLookup?.id === request.id) state.mailLookup = request;
    }

    async function refreshRequests() {
        if (!Services.available()) {
            state.loaded = true;
            state.serviceError = '';
            render();
            return;
        }
        state.loading = !state.loaded;
        if (state.loading) render();
        try {
            state.requests = await Services.listRequests();
            state.serviceError = '';
            if (state.mail?.conversationId) {
                state.mailLookup = state.requests.find(item => item.conversationId === state.mail.conversationId) || null;
            }
        } catch (error) {
            state.serviceError = error?.message || String(error);
        }
        state.loaded = true;
        state.loading = false;
        // Lo storico aperto si ricarica per mostrare le novità (es. risposta appena inviata).
        Object.keys(state.history).forEach(id => {
            if (state.history[id].open) loadHistory(id, true);
        });
        render();
    }

    // Mail aperta in Outlook: la vista Mail mostra la sua richiesta.
    async function handleMailChanged() {
        state.mail = Outlook.currentMail();
        state.mailLookup = null;
        if (state.mail?.conversationId) {
            state.mailLookup = state.requests.find(item => item.conversationId === state.mail.conversationId) || null;
            if (!state.mailLookup && Services.available()) {
                try { state.mailLookup = await Services.getRequestByConversation(state.mail.conversationId); }
                catch (error) { console.warn('[AMP]', error); }
            }
        }
        state.focusId = state.mailLookup?.id || null;
        if (state.mail && state.view !== 'settings') state.view = 'mail';
        render();
    }

    function focusedRequest() {
        if (state.focusId) return state.requests.find(item => item.id === state.focusId) || (state.mailLookup?.id === state.focusId ? state.mailLookup : null);
        return state.mailLookup;
    }
    // Da Coda o Ticket: apre la mail in Outlook e passa alla vista Mail.
    function openRequest(request) {
        state.focusId = request.id;
        state.view = 'mail';
        render();
        scrollToTop();
        Outlook.openMessage(request.messageId);
    }
    function setView(view) {
        state.view = view;
        state.editing = null;
        render();
        scrollToTop();
    }
    // Cambiando vista si riparte dall'alto (se l'area è già scrollata oltre le schede).
    function scrollToTop() {
        if (root && root.getBoundingClientRect().top < 0) root.scrollIntoView({ block: 'start' });
    }

    function buildShell(container) {
        root = h('div', { id: 'ao-root' },
            h('header', { class: 'ao-header' },
                h('button', { type: 'button', class: 'ao-icon-button ao-back', title: TEXTS.back, 'aria-label': TEXTS.back, onclick: () => window.TestPluginAPI?.showMainView?.() }, '←'),
                h('div', { class: 'ao-title' }, TEXTS.title),
                h('span', { class: 'ao-user' })),
            h('nav', { class: 'ao-tabs', role: 'tablist' },
                ['queue', 'tickets', 'mail', 'settings'].map(view => h('button', {
                    type: 'button', role: 'tab', class: 'ao-tab', dataset: { view },
                    onclick: () => setView(view)
                }, h('span', { class: 'ao-tab-label' }, TEXTS.tabs[view]), h('span', { class: 'ao-badge', hidden: true })))),
            h('div', { class: 'ao-body' }),
            h('div', { class: 'ao-status', role: 'status', 'aria-live': 'polite', hidden: true }));
        container.replaceChildren(root);
    }

    function render() {
        if (!root) return;
        const config = loadConfig();
        const tag = myTag(config);
        const user = root.querySelector('.ao-user');
        user.textContent = tag ? `● ${tag}` : '';
        const queueCount = state.requests.filter(item => item.status === 'queue').length;
        const ticketCount = state.requests.filter(item => item.ticket && OPEN_STATES.includes(item.ticket.state) && item.owner === tag).length;
        root.querySelectorAll('.ao-tab').forEach(tab => {
            const active = tab.dataset.view === state.view;
            tab.classList.toggle('ao-active', active);
            tab.setAttribute('aria-selected', String(active));
            const badge = tab.querySelector('.ao-badge');
            const count = { queue: queueCount, tickets: ticketCount }[tab.dataset.view];
            badge.hidden = !count;
            badge.textContent = count || '';
        });
        const views = { queue: renderQueue, tickets: renderTickets, mail: renderMail, settings: renderSettings };
        root.querySelector('.ao-body').replaceChildren(views[state.view](config, tag));
    }

    // Avviso comune quando i servizi mancano o non rispondono.
    function serviceNotice() {
        if (!Services.available()) return h('div', { class: 'ao-notice' }, TEXTS.common.servicesMissing);
        if (state.serviceError) return h('div', { class: 'ao-notice ao-tone-danger' }, TEXTS.common.serviceError(state.serviceError));
        if (state.loading) return h('div', { class: 'ao-notice' }, TEXTS.common.loading);
        return null;
    }
    function filterBar(options, value, onchange) {
        return h('div', { class: 'ao-segmented', role: 'group' },
            options.map(option => h('button', {
                type: 'button', class: `ao-segment${option.value === value ? ' ao-active' : ''}`,
                'aria-pressed': String(option.value === value), onclick: () => onchange(option.value)
            }, option.label, option.count !== undefined ? h('span', { class: 'ao-count' }, option.count) : null)));
    }

    // ======================================================================
    // 6. VISTE
    // ======================================================================

    // ---------------------------- Coda ------------------------------------
    function renderQueue(config, tag) {
        const notice = serviceNotice();
        const queue = state.requests.filter(item => item.status === 'queue')
            .sort((a, b) => String(b.receivedAt).localeCompare(String(a.receivedAt)));
        const mine = queue.filter(item => item.owner && item.owner === tag);
        const free = queue.filter(item => !item.owner);
        const lists = { all: queue, mine, free };
        const items = lists[state.queueFilter] || queue;
        const discarded = state.requests.filter(item => item.status === 'discarded').length;
        return h('div', { class: 'ao-view' },
            notice,
            filterBar([
                { value: 'all', label: TEXTS.queue.all, count: queue.length },
                { value: 'mine', label: TEXTS.queue.mine, count: mine.length },
                { value: 'free', label: TEXTS.queue.free, count: free.length }
            ], state.queueFilter, value => { state.queueFilter = value; render(); }),
            items.length ? h('ul', { class: 'ao-list' }, items.map(request => queueItem(request, config, tag)))
                : (notice ? null : h('p', { class: 'ao-empty' }, TEXTS.queue.empty)),
            discarded ? h('p', { class: 'ao-footnote' }, TEXTS.queue.discarded(discarded)) : null);
    }
    function queueItem(request, config, tag) {
        const suggestion = suggestCompany(config, request.sender, request.subject, request.preview);
        const company = suggestion && config.companies[suggestion.id];
        return h('li', { class: 'ao-item', tabindex: '0', role: 'button', onclick: () => openRequest(request), onkeydown: event => { if (event.key === 'Enter') openRequest(request); } },
            h('div', { class: 'ao-item-top' },
                h('span', { class: 'ao-item-title' }, request.subject || '(senza oggetto)'),
                h('span', { class: 'ao-item-time' }, formatWhen(request.receivedAt))),
            h('div', { class: 'ao-item-sub' }, request.sender?.name || request.sender?.email || ''),
            h('div', { class: 'ao-item-meta' },
                company ? chip(TEXTS.queue.suggested(company.name), 'accent') : null,
                request.owner ? chip(request.owner === tag ? TEXTS.common.me : request.owner, request.owner === tag ? 'success' : 'neutral')
                    : h('button', {
                        type: 'button', class: 'ao-button ao-small',
                        onclick: event => { event.stopPropagation(); takeCharge(request, tag); }
                    }, TEXTS.queue.take)));
    }

    // ---------------------------- Ticket ----------------------------------
    function renderTickets(config, tag) {
        const notice = serviceNotice();
        const open = state.requests.filter(item => item.ticket && OPEN_STATES.includes(item.ticket.state))
            .sort((a, b) => Number(isHoldExpired(b.ticket)) - Number(isHoldExpired(a.ticket))
                || String(b.lastMessageAt || b.receivedAt).localeCompare(String(a.lastMessageAt || a.receivedAt)));
        const mine = open.filter(item => item.owner === tag);
        const items = state.ticketFilter === 'mine' ? mine : open;
        return h('div', { class: 'ao-view' },
            notice,
            filterBar([
                { value: 'mine', label: TEXTS.tickets.mine, count: mine.length },
                { value: 'all', label: TEXTS.tickets.all, count: open.length }
            ], state.ticketFilter, value => { state.ticketFilter = value; render(); }),
            items.length ? h('ul', { class: 'ao-list' }, items.map(request => ticketItem(request, tag)))
                : (notice ? null : h('p', { class: 'ao-empty' }, TEXTS.tickets.empty)));
    }
    function ticketItem(request, tag) {
        const ticket = request.ticket;
        const holdDate = ticket.state === '3' ? parseDate(ticket.holdExpiration) : null;
        return h('li', { class: `ao-item${isHoldExpired(ticket) ? ' ao-item-alert' : ''}`, tabindex: '0', role: 'button', onclick: () => openRequest(request), onkeydown: event => { if (event.key === 'Enter') openRequest(request); } },
            h('div', { class: 'ao-item-top' },
                h('span', { class: 'ao-item-number' }, ticket.number),
                stateChip(ticket)),
            h('div', { class: 'ao-item-title' }, ticket.shortDescription || request.subject),
            h('div', { class: 'ao-item-sub' }, [ticket.companyName, ticket.activityName].filter(Boolean).join(' · ')),
            h('div', { class: 'ao-item-meta' },
                chip(request.owner === tag ? TEXTS.common.me : (request.owner || TEXTS.common.unassigned), request.owner === tag ? 'success' : 'neutral'),
                h('span', { class: 'ao-item-time' }, TEXTS.mail.minutes(ticket.timeWorkedMinutes || 0)),
                holdDate ? h('span', { class: 'ao-item-time' }, TEXTS.tickets.holdUntil(formatWhen(holdDate))) : null,
                request.pendingAction ? chip(pendingLabel(request.pendingAction), 'warning') : null));
    }

    // ---------------------------- Mail ------------------------------------
    function renderMail(config, tag) {
        const notice = serviceNotice();
        const request = focusedRequest();
        if (!request) {
            if (state.mail && Services.available() && !state.serviceError) {
                return h('div', { class: 'ao-view' },
                    mailHeader({ subject: state.mail.subject, sender: state.mail.sender, receivedAt: state.mail.receivedAt }),
                    h('div', { class: 'ao-card' },
                        h('p', { class: 'ao-muted' }, TEXTS.mail.notRegistered),
                        h('button', { type: 'button', class: 'ao-button', onclick: registerCurrentMail }, TEXTS.mail.register)));
            }
            return h('div', { class: 'ao-view' }, notice, h('p', { class: 'ao-empty' }, TEXTS.mail.noMail));
        }
        const isCurrentMail = state.mail?.conversationId === request.conversationId;
        return h('div', { class: 'ao-view' },
            notice,
            mailHeader(request, !isCurrentMail),
            ownerCard(request, tag),
            request.status === 'discarded' ? discardedCard(request) : null,
            request.status === 'queue' ? newTicketCard(request, config, tag) : null,
            request.ticket ? ticketCard(request, tag) : null,
            request.ticket && OPEN_STATES.includes(request.ticket.state) ? onSendCard(request, config, isCurrentMail) : null,
            request.ticket ? historyCard(request) : null);
    }
    function mailHeader(request, showOpen) {
        return h('div', { class: 'ao-mail-head' },
            h('div', { class: 'ao-mail-subject' }, request.subject || '(senza oggetto)'),
            h('div', { class: 'ao-mail-from' },
                h('span', {}, `${request.sender?.name || ''}${request.sender?.email ? ` <${request.sender.email}>` : ''}`),
                h('span', { class: 'ao-item-time' }, formatWhen(request.receivedAt))),
            showOpen ? h('button', { type: 'button', class: 'ao-link', onclick: () => Outlook.openMessage(request.messageId) }, TEXTS.mail.openInOutlook) : null);
    }
    function ownerCard(request, tag) {
        const mine = request.owner && request.owner === tag;
        const action = !request.owner
            ? h('button', { type: 'button', class: 'ao-button ao-small', onclick: () => takeCharge(request, tag) }, TEXTS.mail.takeCharge)
            : mine
                ? h('button', { type: 'button', class: 'ao-button ao-small ao-subtle', onclick: () => takeCharge(request, '') }, TEXTS.mail.release)
                : confirmButton(TEXTS.mail.takeOver(request.owner), 'ao-button ao-small', () => takeCharge(request, tag));
        return h('div', { class: 'ao-owner' },
            h('span', { class: `ao-owner-dot${request.owner ? (mine ? ' ao-tone-success' : ' ao-tone-neutral') : ''}` }),
            h('span', { class: 'ao-owner-text' }, request.owner ? TEXTS.mail.inCharge(mine ? `${request.owner} (${TEXTS.common.me})` : request.owner) : TEXTS.mail.notInCharge),
            request.status === 'discarded' ? null : action);
    }
    function discardedCard(request) {
        return h('div', { class: 'ao-card' },
            h('p', { class: 'ao-muted' }, TEXTS.mail.discardedBy(request.discardedBy || TEXTS.common.none, formatWhen(request.discardedAt))),
            h('button', { type: 'button', class: 'ao-button', onclick: () => run(async () => upsertRequest(await Services.restoreRequest(request.id))).then(render) }, TEXTS.mail.restore));
    }

    // Apertura del ticket: Agenzia suggerita, Attività, anteprima dei campi.
    function ticketDraft(request, config) {
        if (!state.ticketDrafts[request.id]) {
            const suggestion = suggestCompany(config, request.sender, request.subject, request.preview);
            state.ticketDrafts[request.id] = { companyId: suggestion?.id || '', activityId: '', subject: request.subject || '', suggestion };
        }
        return state.ticketDrafts[request.id];
    }
    function newTicketCard(request, config, tag) {
        const draft = ticketDraft(request, config);
        const companies = Object.entries(config.companies).sort(byName);
        const activities = Object.entries(config.activities).sort(byName);
        const company = config.companies[draft.companyId];
        const activity = config.activities[draft.activityId];
        const favorites = config.favorites.filter(item => config.companies[item.company] && config.activities[item.activity]);
        const shortDescription = buildShortDescription(draft.subject, activity);
        const rerender = () => render();
        return h('section', { class: 'ao-card' },
            h('h3', { class: 'ao-card-title' }, TEXTS.mail.newTicket),
            favorites.length ? h('div', { class: 'ao-favorites' },
                h('span', { class: 'ao-label' }, TEXTS.mail.favorites),
                h('div', { class: 'ao-chips' }, favorites.map(item => h('button', {
                    type: 'button',
                    class: `ao-fav${item.company === draft.companyId && item.activity === draft.activityId ? ' ao-active' : ''}`,
                    onclick: () => { Object.assign(draft, { companyId: item.company, activityId: item.activity }); rerender(); }
                }, `${config.companies[item.company].name} · ${config.activities[item.activity].name}`)))) : null,
            selectField(TEXTS.mail.company, [{ value: '', label: TEXTS.mail.choose }, ...companies.map(([id, item]) => ({ value: id, label: item.name }))],
                draft.companyId, value => { draft.companyId = value; rerender(); }),
            draft.suggestion && draft.suggestion.id === draft.companyId
                ? h('div', { class: 'ao-hint ao-hint-accent' }, TEXTS.mail.suggestedFrom(draft.suggestion.reason)) : null,
            selectField(TEXTS.mail.activity, [{ value: '', label: TEXTS.mail.choose }, ...activities.map(([id, item]) => ({ value: id, label: item.name }))],
                draft.activityId, value => { draft.activityId = value; rerender(); }),
            textField(TEXTS.mail.shortDescription, draft.subject, value => {
                draft.subject = value;
                const preview = root.querySelector('[data-preview="short"]');
                if (preview) preview.textContent = buildShortDescription(value, activity);
            }, { maxlength: '160' }),
            company && activity ? h('dl', { class: 'ao-fields' },
                h('dt', {}, TEXTS.mail.fieldsPreview), h('dd', {}),
                fieldRow('Company', company.company?.displayValue || company.name),
                fieldRow('Assignment group', company.assignmentGroup?.displayValue),
                fieldRow('Assigned to', config.profile.assignedTo.displayValue || TEXTS.common.none),
                fieldRow('Category', activity.category?.displayValue),
                fieldRow('Impact / Urgency', `${activity.impact || '?'} / ${activity.urgency || '?'}`),
                fieldRow('Time worked', TEXTS.mail.minutes(Number(activity.timeWorked) || 0)),
                fieldRow('Short description', shortDescription, 'short'),
                fieldRow('Reported', formatWhen(request.receivedAt))) : null,
            h('div', { class: 'ao-actions' },
                h('button', { type: 'button', class: 'ao-button ao-primary', disabled: !(company && activity), onclick: () => openTicket(request, tag) }, TEXTS.mail.openTicket),
                confirmButton(TEXTS.mail.discard, 'ao-button ao-subtle', () => run(async () => upsertRequest(await Services.discardRequest(request.id, tag))).then(render))));
    }
    function fieldRow(label, value, previewKey) {
        return [h('dt', {}, label), h('dd', previewKey ? { dataset: { preview: previewKey } } : {}, value || TEXTS.common.none)];
    }
    function buildShortDescription(subject, activity) {
        const clean = String(subject || '').replace(/^\s*((re|r|fw|fwd|i)\s*:\s*)+/i, '').trim();
        const prefix = activity?.shortDescriptionPrefix || '';
        return prefix ? `${prefix} ${clean.replace(/^\s*\[[^\]]+\]\s*/, '')}`.trim() : clean;
    }
    async function openTicket(request, tag) {
        const config = loadConfig();
        const draft = ticketDraft(request, config);
        const company = config.companies[draft.companyId];
        const activity = config.activities[draft.activityId];
        if (!company || !activity) return showStatus(TEXTS.mail.selectBoth, 'danger');
        if (!config.profile.assignedTo.displayValue || !tag) return showStatus(TEXTS.mail.profileMissing, 'danger');
        // Corpo completo dalla mail aperta, se è quella della richiesta; altrimenti l'anteprima salvata.
        const body = state.mail?.conversationId === request.conversationId ? state.mail.body : request.preview;
        const fields = {
            company: normalizeReference(company.company),
            assignmentGroup: normalizeReference(company.assignmentGroup),
            assignedTo: config.profile.assignedTo,
            category: normalizeReference(activity.category),
            impact: activity.impact,
            urgency: activity.urgency,
            timeWorkedMinutes: Number(activity.timeWorked) || 0,
            shortDescription: buildShortDescription(draft.subject, activity),
            description: String(body || '').slice(0, 4000),
            reportedOn: request.receivedAt ? formatServiceNowDate(parseDate(request.receivedAt)) : ''
        };
        const updated = await run(() => Services.createTicket(request.id, {
            companyId: draft.companyId, companyName: company.name,
            activityId: draft.activityId, activityName: activity.name,
            openedBy: tag, fields
        }), result => TEXTS.mail.ticketOpened(result.ticket.number));
        if (!updated) return;
        upsertRequest(updated);
        delete state.ticketDrafts[request.id];
        config.recent = [{ company: draft.companyId, activity: draft.activityId, timestamp: new Date().toISOString() },
            ...config.recent.filter(item => !(item.company === draft.companyId && item.activity === draft.activityId))].slice(0, 10);
        saveConfig(config);
        render();
    }

    // Ticket collegato: stato, tempo lavorato.
    function ticketCard(request) {
        const ticket = request.ticket;
        const open = OPEN_STATES.includes(ticket.state);
        return h('section', { class: 'ao-card' },
            h('div', { class: 'ao-ticket-head' },
                h('a', { class: 'ao-ticket-number', href: ticket.url || '#', target: '_blank', rel: 'noopener', title: TEXTS.mail.openInServiceNow }, ticket.number),
                stateChip(ticket)),
            h('div', { class: 'ao-item-sub' }, [ticket.companyName, ticket.activityName].filter(Boolean).join(' · ')),
            ticket.state === '3' && ticket.holdExpiration
                ? h('div', { class: `ao-hint${isHoldExpired(ticket) ? ' ao-hint-danger' : ''}` },
                    `${HOLD_REASONS.find(item => item.value === ticket.holdReason)?.label || 'On Hold'} · ${TEXTS.tickets.holdUntil(formatWhen(ticket.holdExpiration))}`) : null,
            h('div', { class: 'ao-hint' }, TEXTS.mail.openedBy(ticket.openedBy || TEXTS.common.none, formatWhen(ticket.openedAt))),
            h('div', { class: 'ao-time' },
                h('span', { class: 'ao-label' }, TEXTS.mail.timeWorked),
                h('div', { class: 'ao-time-row' },
                    TIME_STEPS.filter(step => step < 0).map(step => timeButton(request, step, open)),
                    h('span', { class: 'ao-time-value' }, TEXTS.mail.minutes(ticket.timeWorkedMinutes || 0)),
                    TIME_STEPS.filter(step => step > 0).map(step => timeButton(request, step, open)))),
            open ? null : h('p', { class: 'ao-muted' }, TEXTS.mail.closed(TICKET_STATES[ticket.state]?.label || ticket.state)));
    }
    function timeButton(request, step, enabled) {
        return h('button', {
            type: 'button', class: 'ao-button ao-small ao-time-step',
            disabled: !enabled || (step < 0 && (request.ticket.timeWorkedMinutes || 0) <= 0),
            onclick: async () => {
                const tag = myTag();
                const before = request.ticket.timeWorkedMinutes || 0;
                const updated = await run(() => Services.addTime(request.id, step, tag));
                if (!updated) return;
                upsertRequest(updated);
                render();
                showStatus(TEXTS.mail.timeChanged(updated.ticket.timeWorkedMinutes - before, updated.ticket.timeWorkedMinutes));
            }
        }, `${step > 0 ? '+' : '−'}${Math.abs(step)}`);
    }

    // Azione da eseguire su ServiceNow quando parte la risposta alla mail.
    function pendingLabel(action) {
        if (!action || action.type === 'comment') return TEXTS.mail.actions.comment;
        if (action.type === 'resolve') return `${TEXTS.mail.actions.resolve} · ${action.resolutionCode}`;
        if (action.type === 'cancel') return TEXTS.mail.actions.cancel;
        if (action.type === 'hold') return `${TEXTS.mail.actions.hold} · ${formatWhen(action.holdExpiration)}`;
        return action.type;
    }
    function actionDraft(request, config) {
        if (!state.actionDrafts[request.id]) {
            const saved = request.pendingAction || {};
            const key = combinationKey(request.ticket.companyId, request.ticket.activityId);
            state.actionDrafts[request.id] = {
                type: saved.type || 'comment',
                resolutionCode: saved.resolutionCode || RESOLUTION_CODES[0],
                holdReason: saved.holdReason || HOLD_REASONS[0].value,
                holdExpiration: saved.holdExpiration ? toDateTimeInput(parseDate(saved.holdExpiration)) : defaultHoldExpiration(),
                workNotes: saved.workNotes || '',
                template: config.combinationResponses[key]?.response ? 'combination' : ''
            };
        }
        return state.actionDrafts[request.id];
    }
    function onSendCard(request, config, isCurrentMail) {
        const draft = actionDraft(request, config);
        const rerender = () => render();
        const key = combinationKey(request.ticket.companyId, request.ticket.activityId);
        const templates = [
            { value: '', label: TEXTS.mail.replyEmpty },
            ...(config.combinationResponses[key]?.response ? [{ value: 'combination', label: TEXTS.mail.replyCombination }] : []),
            ...Object.entries(config.usefulResponses).sort(byName).map(([id, item]) => ({ value: `useful:${id}`, label: item.title }))
        ];
        return h('section', { class: 'ao-card' },
            h('h3', { class: 'ao-card-title' }, TEXTS.mail.onSend),
            h('p', { class: 'ao-hint' }, TEXTS.mail.onSendHint),
            request.pendingAction ? h('div', { class: 'ao-pending' },
                h('span', {}, TEXTS.mail.pending(pendingLabel(request.pendingAction))),
                h('button', { type: 'button', class: 'ao-link', onclick: () => savePendingAction(request, null) }, TEXTS.mail.clearPending)) : null,
            filterBar(Object.entries(TEXTS.mail.actions).map(([value, label]) => ({ value, label })), draft.type, value => { draft.type = value; rerender(); }),
            draft.type === 'resolve' ? selectField(TEXTS.mail.resolutionCode, RESOLUTION_CODES.map(code => ({ value: code, label: code })), draft.resolutionCode, value => { draft.resolutionCode = value; }) : null,
            draft.type === 'hold' ? [
                selectField(TEXTS.mail.holdReason, HOLD_REASONS, draft.holdReason, value => { draft.holdReason = value; }),
                textField(TEXTS.mail.holdExpiration, draft.holdExpiration, value => { draft.holdExpiration = value; }, { type: 'datetime-local' })
            ] : null,
            draft.type !== 'comment' ? textField(TEXTS.mail.workNotes, draft.workNotes, value => { draft.workNotes = value; }, { multiline: true, rows: '2' }) : null,
            selectField(TEXTS.mail.replyText, templates, draft.template, value => { draft.template = value; }),
            h('div', { class: 'ao-actions' },
                h('button', {
                    type: 'button', class: 'ao-button ao-primary',
                    title: isCurrentMail ? '' : TEXTS.mail.replyUnavailable,
                    onclick: () => prepareReply(request, isCurrentMail)
                }, TEXTS.mail.prepareReply)));
    }
    function buildPendingAction(draft, tag) {
        if (draft.type === 'comment') return null;
        const action = { type: draft.type, workNotes: draft.workNotes.trim(), by: tag };
        if (draft.type === 'resolve') action.resolutionCode = draft.resolutionCode;
        if (draft.type === 'hold') {
            const date = parseDate(draft.holdExpiration.replace('T', ' '));
            action.holdReason = draft.holdReason;
            action.holdExpiration = date ? formatServiceNowDate(date) : '';
        }
        return action;
    }
    async function savePendingAction(request, action) {
        const updated = await run(() => Services.setPendingAction(request.id, action));
        if (!updated) return null;
        upsertRequest(updated);
        delete state.actionDrafts[request.id];
        render();
        return updated;
    }
    async function prepareReply(request, isCurrentMail) {
        const config = loadConfig();
        const draft = actionDraft(request, config);
        if (draft.type === 'hold' && !parseDate(draft.holdExpiration)) return showStatus(TEXTS.mail.holdExpiration, 'danger');
        const template = draft.template === 'combination'
            ? config.combinationResponses[combinationKey(request.ticket.companyId, request.ticket.activityId)]?.response
            : config.usefulResponses[draft.template.replace(/^useful:/, '')]?.response;
        const updated = await savePendingAction(request, buildPendingAction(draft, myTag(config)));
        if (!updated) return;
        // In Outlook la risposta si apre solo sulla mail aperta: per un'altra mail
        // si apre la mail e l'azione resta salvata. L'anteprima le gestisce entrambe.
        let opened = false;
        if (isCurrentMail) opened = Outlook.openReply(template || '');
        else {
            Outlook.openMessage(request.messageId);
            if (window.AmpPreviewOutlook) opened = Outlook.openReply(template || '');
        }
        showStatus(opened ? TEXTS.mail.replyOpened : TEXTS.mail.pendingSaved, opened ? 'success' : 'warning');
    }

    // Storico di ServiceNow, caricato solo quando serve e a pagine.
    function historyCard(request) {
        const history = state.history[request.id];
        return h('section', { class: 'ao-card' },
            h('div', { class: 'ao-card-head' },
                h('h3', { class: 'ao-card-title' }, TEXTS.mail.history),
                history?.open ? null : h('button', { type: 'button', class: 'ao-button ao-small', onclick: () => loadHistory(request.id) }, TEXTS.mail.loadHistory)),
            history?.open ? [
                history.loading && !history.entries.length ? h('p', { class: 'ao-muted' }, TEXTS.common.loading) : null,
                !history.loading && !history.entries.length ? h('p', { class: 'ao-muted' }, TEXTS.mail.historyEmpty) : null,
                h('ol', { class: 'ao-history' }, history.entries.map(entry => h('li', { class: `ao-entry ao-entry-${entry.kind}` },
                    h('div', { class: 'ao-entry-head' },
                        h('span', { class: 'ao-entry-kind' }, TEXTS.mail.kinds[entry.kind] || entry.kind),
                        h('span', {}, entry.author),
                        h('span', { class: 'ao-item-time' }, formatWhen(entry.at))),
                    h('div', { class: 'ao-entry-text' }, entry.text)))),
                history.total > history.entries.length
                    ? h('button', { type: 'button', class: 'ao-link', onclick: () => loadHistory(request.id, false, true) }, TEXTS.mail.loadMore(history.total - history.entries.length)) : null
            ] : null);
    }
    async function loadHistory(requestId, reload = false, more = false) {
        const current = state.history[requestId] || { entries: [], total: 0 };
        const limit = reload ? Math.max(HISTORY_PAGE, current.entries.length) : HISTORY_PAGE;
        const offset = more ? current.entries.length : 0;
        state.history[requestId] = { ...current, open: true, loading: true };
        if (!reload) render();
        try {
            const page = await Services.getTicketActivity(requestId, { offset, limit });
            state.history[requestId] = {
                open: true, loading: false, total: page.total,
                entries: more ? [...current.entries, ...page.entries] : page.entries
            };
        } catch (error) {
            state.history[requestId] = { ...current, open: false, loading: false };
            showStatus(error?.message || String(error), 'danger');
        }
        render();
    }

    // tag = chi prende in carico; '' = rilascia (solo chi ce l'ha in carico).
    async function takeCharge(request, tag) {
        const release = tag === '';
        if (!release && !String(tag || '').trim()) return showStatus(TEXTS.mail.profileMissing, 'danger');
        if (release && request.owner !== myTag()) return;
        const updated = await run(() => Services.takeCharge(request.id, release ? '' : tag.trim()));
        if (!updated) return;
        upsertRequest(updated);
        render();
    }
    async function registerCurrentMail() {
        if (!state.mail) return;
        const created = await run(() => Services.registerRequest(state.mail));
        if (!created) return;
        upsertRequest(created);
        state.mailLookup = created;
        state.focusId = created.id;
        render();
    }

    // ---------------------------- Impostazioni ----------------------------
    // Ogni sezione è un elenco con modifica in linea. EDITORS descrive i campi di ogni tipo di voce.
    const EDITORS = {
        company: {
            collection: 'companies',
            title: TEXTS.settings.companies,
            describe: item => [item.assignmentGroup?.displayValue, (item.domains || []).map(domain => `@${domain}`).join(' ')].filter(Boolean).join(' · '),
            fields: [
                { key: 'name', label: TEXTS.settings.companyName, required: true },
                { key: 'company', label: TEXTS.settings.company, reference: true },
                { key: 'assignmentGroup', label: TEXTS.settings.assignmentGroup, reference: true },
                { key: 'domains', label: TEXTS.settings.domains, hint: TEXTS.settings.domainsHint, list: true }
            ]
        },
        activity: {
            collection: 'activities',
            title: TEXTS.settings.activities,
            describe: item => [item.category?.displayValue, `I/U ${item.impact || '?'}/${item.urgency || '?'}`, TEXTS.mail.minutes(Number(item.timeWorked) || 0), item.shortDescriptionPrefix].filter(Boolean).join(' · '),
            fields: [
                { key: 'name', label: TEXTS.settings.companyName, required: true },
                { key: 'category', label: TEXTS.settings.category, reference: true },
                { key: 'impact', label: TEXTS.settings.impact, options: ['1', '2', '3'] },
                { key: 'urgency', label: TEXTS.settings.urgency, options: ['1', '2', '3'] },
                { key: 'timeWorked', label: TEXTS.settings.timeWorked, number: true },
                { key: 'shortDescriptionPrefix', label: TEXTS.settings.prefix }
            ]
        },
        response: {
            collection: 'usefulResponses',
            title: TEXTS.settings.responses,
            nameKey: 'title',
            describe: item => String(item.response || '').replace(/\s+/g, ' ').slice(0, 70),
            fields: [
                { key: 'title', label: TEXTS.settings.title, required: true },
                { key: 'response', label: TEXTS.settings.response, multiline: true }
            ]
        }
    };

    function renderSettings(config) {
        return h('div', { class: 'ao-view' },
            profileSection(config),
            Object.keys(EDITORS).map(kind => collectionSection(kind, config)),
            h('section', { class: 'ao-card' },
                h('h3', { class: 'ao-card-title' }, TEXTS.settings.data),
                h('div', { class: 'ao-actions' },
                    h('button', { type: 'button', class: 'ao-button', onclick: importConfig }, TEXTS.settings.import),
                    h('button', { type: 'button', class: 'ao-button', onclick: exportConfig }, TEXTS.settings.export))));
    }
    function profileSection(config) {
        const draft = { tag: config.profile.tag, name: config.profile.assignedTo.displayValue, sysId: config.profile.assignedTo.value };
        return h('section', { class: 'ao-card' },
            h('h3', { class: 'ao-card-title' }, TEXTS.settings.profile),
            textField(TEXTS.settings.tag, draft.tag, value => { draft.tag = value; }, { placeholder: myTag(createDefaultConfig()) || 'Marco', hint: TEXTS.settings.tagHint }),
            h('div', { class: 'ao-grid-2' },
                textField(`${TEXTS.settings.assignedTo} · ${TEXTS.settings.displayValue}`, draft.name, value => { draft.name = value; }),
                textField(TEXTS.settings.sysId, draft.sysId, value => { draft.sysId = value; })),
            h('div', { class: 'ao-actions' },
                h('button', {
                    type: 'button', class: 'ao-button ao-primary',
                    onclick: () => {
                        const current = loadConfig();
                        current.profile = { tag: draft.tag.trim(), assignedTo: { displayValue: draft.name.trim(), value: draft.sysId.trim() } };
                        saveConfig(current);
                        render();
                        showStatus(TEXTS.settings.saved);
                    }
                }, TEXTS.common.save)));
    }
    function collectionSection(kind, config) {
        const editor = EDITORS[kind];
        const items = Object.entries(config[editor.collection]).sort(byName);
        const editing = state.editing?.kind === kind ? state.editing : null;
        return h('section', { class: 'ao-card' },
            h('div', { class: 'ao-card-head' },
                h('h3', { class: 'ao-card-title' }, editor.title, h('span', { class: 'ao-count' }, items.length)),
                editing?.id === null ? null : h('button', { type: 'button', class: 'ao-button ao-small', onclick: () => { state.editing = { kind, id: null }; render(); } }, `+ ${TEXTS.common.add}`)),
            editing?.id === null ? editorForm(kind, null, {}) : null,
            items.length ? h('ul', { class: 'ao-config-list' }, items.map(([id, item]) => editing?.id === id
                ? h('li', {}, editorForm(kind, id, item))
                : h('li', { class: 'ao-config-item' },
                    h('div', { class: 'ao-config-text' },
                        h('span', { class: 'ao-config-name' }, item[editor.nameKey || 'name']),
                        h('span', { class: 'ao-hint' }, editor.describe(item))),
                    h('button', { type: 'button', class: 'ao-icon-button', title: TEXTS.common.edit, 'aria-label': TEXTS.common.edit, onclick: () => { state.editing = { kind, id }; render(); } }, '✎'),
                    confirmButton('🗑', 'ao-icon-button ao-danger', () => deleteItem(kind, id)))))
                : (editing ? null : h('p', { class: 'ao-muted' }, TEXTS.settings.empty)));
    }
    function editorForm(kind, id, item) {
        const editor = EDITORS[kind];
        const values = {};
        const fields = editor.fields.map(field => {
            if (field.reference) {
                const reference = normalizeReference(item[field.key]);
                values[field.key] = { ...reference };
                return h('div', { class: 'ao-grid-2' },
                    textField(`${field.label} · ${TEXTS.settings.displayValue}`, reference.displayValue, value => { values[field.key].displayValue = value; }),
                    textField(TEXTS.settings.sysId, reference.value, value => { values[field.key].value = value; }));
            }
            if (field.options) {
                values[field.key] = String(item[field.key] || '3');
                return selectField(field.label, field.options.map(value => ({ value, label: SEVERITY_LABELS[value] || value })), values[field.key], value => { values[field.key] = value; });
            }
            const initial = field.list ? (item[field.key] || []).join(', ') : (item[field.key] ?? '');
            values[field.key] = String(initial);
            return textField(field.label, initial, value => { values[field.key] = value; },
                { hint: field.hint, multiline: field.multiline, rows: field.multiline ? '5' : undefined, type: field.number ? 'number' : undefined, min: field.number ? '0' : undefined });
        });
        const save = () => {
            const nameKey = editor.nameKey || 'name';
            if (!String(values[nameKey] || '').trim()) return showStatus(TEXTS.settings.nameRequired, 'danger');
            const config = loadConfig();
            const targetId = id || createId(values[nameKey]);
            const next = { ...(config[editor.collection][targetId] || {}) };
            editor.fields.forEach(field => {
                const value = values[field.key];
                if (field.reference) next[field.key] = { value: value.value.trim(), displayValue: value.displayValue.trim() };
                else if (field.list) next[field.key] = normalizeDomains(value);
                else if (field.number) next[field.key] = Math.max(0, Number.parseInt(value, 10) || 0);
                else next[field.key] = field.multiline ? String(value) : String(value).trim();
            });
            if (kind === 'response') {
                const now = new Date().toISOString();
                next.created ||= now;
                next.updated = now;
            }
            config[editor.collection][targetId] = next;
            saveConfig(config);
            state.editing = null;
            render();
            showStatus(TEXTS.settings.saved);
        };
        return h('div', { class: 'ao-editor' },
            fields,
            h('div', { class: 'ao-actions' },
                h('button', { type: 'button', class: 'ao-button ao-primary', onclick: save }, TEXTS.common.save),
                h('button', { type: 'button', class: 'ao-button ao-subtle', onclick: () => { state.editing = null; render(); } }, TEXTS.common.cancel)));
    }
    function deleteItem(kind, id) {
        const config = loadConfig();
        const editor = EDITORS[kind];
        delete config[editor.collection][id];
        if (kind === 'company' || kind === 'activity') {
            const key = kind === 'company' ? 'company' : 'activity';
            config.favorites = config.favorites.filter(item => item[key] !== id);
            config.recent = config.recent.filter(item => item[key] !== id);
            Object.keys(config.combinationResponses).forEach(responseKey => {
                if (config.combinationResponses[responseKey][key] === id) delete config.combinationResponses[responseKey];
            });
        }
        saveConfig(config);
        render();
        showStatus(TEXTS.settings.deleted);
    }
    function exportConfig() {
        const blob = new Blob([JSON.stringify(loadConfig(), null, 2)], { type: 'application/json' });
        const link = h('a', { href: URL.createObjectURL(blob), download: `amp-configurazioni-${new Date().toISOString().slice(0, 10)}.json` });
        document.body.append(link);
        link.click();
        setTimeout(() => { URL.revokeObjectURL(link.href); link.remove(); }, 0);
    }
    function importConfig() {
        const input = h('input', { type: 'file', accept: 'application/json,.json', hidden: true });
        input.addEventListener('change', () => {
            const file = input.files?.[0];
            input.remove();
            if (!file) return;
            const reader = new FileReader();
            reader.onload = () => {
                try {
                    const parsed = JSON.parse(reader.result);
                    if (!parsed || typeof parsed !== 'object' || (!parsed.companies && !parsed.activities)) throw new Error('formato');
                    saveConfig(normalizeConfig(parsed));
                    state.ticketDrafts = {};
                    state.actionDrafts = {};
                    render();
                    showStatus(TEXTS.settings.imported);
                } catch (error) {
                    showStatus(TEXTS.settings.importInvalid, 'danger');
                }
            };
            reader.readAsText(file);
        });
        document.body.append(input);
        input.click();
    }

    // ======================================================================
    // 7. STILI (solo dentro #ao-root: non toccano il resto dell'add-in)
    // ======================================================================

    function themeCssVariables() {
        const kebab = value => value.replace(/[A-Z]/g, char => `-${char.toLowerCase()}`);
        const lines = [];
        ['colors', 'fontSize', 'fontWeight', 'radius', 'shadow', 'state', 'layout'].forEach(group => {
            const prefix = { colors: '', fontSize: 'fs-', fontWeight: 'fw-', radius: 'radius-', shadow: 'shadow-', state: '', layout: '' }[group];
            Object.entries(THEME[group]).forEach(([key, value]) => lines.push(`--ao-${prefix}${kebab(key)}: ${value};`));
        });
        lines.push(`--ao-font: ${THEME.font.family};`);
        return `#ao-root { ${lines.join(' ')} }`;
    }
    const STYLES = `
#ao-root { display: flex; flex-direction: column; min-height: 100%; font-family: var(--ao-font); font-size: var(--ao-fs-base); color: var(--ao-text); background: var(--ao-bg); line-height: 1.4; }
#ao-root *, #ao-root *::before, #ao-root *::after { box-sizing: border-box; }
#ao-root [hidden] { display: none !important; }
#ao-root.ao-busy { cursor: progress; }
#ao-root button { font: inherit; color: inherit; }

/* Intestazione e schede */
#ao-root .ao-header { display: flex; align-items: center; gap: 8px; padding: 10px 12px; background: var(--ao-header-background); color: var(--ao-white); }
#ao-root .ao-title { flex: 1; font-size: var(--ao-fs-xl); font-weight: var(--ao-fw-medium); }
#ao-root .ao-user { font-size: var(--ao-fs-md); opacity: .9; }
#ao-root .ao-header .ao-icon-button { color: var(--ao-white); border-color: rgba(255,255,255,.35); }
#ao-root .ao-header .ao-icon-button:hover { background: rgba(255,255,255,.15); }
#ao-root .ao-tabs { position: sticky; top: 0; z-index: 2; display: flex; background: var(--ao-white); border-bottom: 1px solid var(--ao-line); }
#ao-root .ao-tab { flex: 1; display: flex; align-items: center; justify-content: center; gap: 5px; padding: 9px 4px 7px; border: 0; border-bottom: 2px solid transparent; background: none; color: var(--ao-text-muted); font-size: var(--ao-fs-md); font-weight: var(--ao-fw-medium); cursor: pointer; }
#ao-root .ao-tab:hover { color: var(--ao-text); background: var(--ao-bg-soft); }
#ao-root .ao-tab.ao-active { color: var(--ao-accent); border-bottom-color: var(--ao-accent); }
#ao-root .ao-badge { min-width: 18px; padding: 0 5px; border-radius: var(--ao-radius-pill); background: var(--ao-accent); color: var(--ao-white); font-size: var(--ao-fs-xs); line-height: 16px; text-align: center; }
#ao-root .ao-body { flex: 1; }
#ao-root .ao-view { display: grid; gap: 10px; padding: 12px; }

/* Pulsanti */
#ao-root .ao-button { display: inline-flex; align-items: center; justify-content: center; gap: 6px; min-height: var(--ao-control-height); padding: 0 12px; border: 1px solid var(--ao-line-strong); border-radius: var(--ao-radius-base); background: var(--ao-white); font-weight: var(--ao-fw-medium); cursor: pointer; }
#ao-root .ao-button:hover:where(:not(:disabled)) { background: var(--ao-bg-soft); border-color: var(--ao-text-subtle); }
#ao-root .ao-primary { background: var(--ao-accent); border-color: var(--ao-accent); color: var(--ao-white); }
#ao-root .ao-primary:hover:where(:not(:disabled)) { background: var(--ao-accent-dark); border-color: var(--ao-accent-dark); }
#ao-root .ao-subtle { border-color: transparent; background: transparent; color: var(--ao-text-muted); }
#ao-root .ao-small { min-height: 26px; padding: 0 9px; font-size: var(--ao-fs-md); }
#ao-root .ao-armed, #ao-root .ao-armed:hover { background: var(--ao-danger-bg) !important; border-color: var(--ao-danger) !important; color: var(--ao-danger) !important; }
#ao-root .ao-icon-button { display: inline-flex; align-items: center; justify-content: center; min-width: 28px; height: 28px; padding: 0 6px; border: 1px solid transparent; border-radius: var(--ao-radius-base); background: none; color: var(--ao-text-muted); cursor: pointer; }
#ao-root .ao-icon-button:hover { background: var(--ao-bg-soft); border-color: var(--ao-line); }
#ao-root .ao-danger:hover { color: var(--ao-danger); }
#ao-root .ao-link { padding: 0; border: 0; background: none; color: var(--ao-accent); font-size: var(--ao-fs-md); cursor: pointer; text-align: left; }
#ao-root .ao-link:hover { text-decoration: underline; }
#ao-root :disabled { opacity: var(--ao-disabled-opacity); cursor: not-allowed; }
#ao-root :focus-visible { outline: none; box-shadow: var(--ao-shadow-focus); }

/* Filtri e scelte a segmenti */
#ao-root .ao-segmented { display: flex; padding: 2px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-white); }
#ao-root .ao-segment { flex: 1; display: inline-flex; align-items: center; justify-content: center; gap: 5px; min-height: 28px; padding: 0 6px; border: 0; border-radius: calc(var(--ao-radius-base) - 2px); background: none; color: var(--ao-text-muted); font-size: var(--ao-fs-md); cursor: pointer; white-space: nowrap; }
#ao-root .ao-segment:hover { color: var(--ao-text); background: var(--ao-bg-soft); }
#ao-root .ao-segment.ao-active { background: var(--ao-accent-bg); color: var(--ao-accent); font-weight: var(--ao-fw-medium); }
#ao-root .ao-count { color: var(--ao-text-subtle); font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-normal); }
#ao-root .ao-card-title .ao-count { margin-left: 6px; }

/* Elenchi di Coda e Ticket */
#ao-root .ao-list { display: grid; gap: 6px; margin: 0; padding: 0; list-style: none; }
#ao-root .ao-item { display: grid; gap: 3px; padding: 9px 10px; border: 1px solid var(--ao-line); border-left: 3px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-white); box-shadow: var(--ao-shadow-card); cursor: pointer; }
#ao-root .ao-item:hover { border-color: var(--ao-accent-border); border-left-color: var(--ao-accent); background: var(--ao-accent-bg); }
#ao-root .ao-item-alert { border-left-color: var(--ao-danger); }
#ao-root .ao-item-top { display: flex; align-items: baseline; justify-content: space-between; gap: 8px; }
#ao-root .ao-item-title { font-weight: var(--ao-fw-medium); overflow-wrap: anywhere; }
#ao-root .ao-item-number { font-weight: var(--ao-fw-bold); color: var(--ao-accent); }
#ao-root .ao-item-sub { color: var(--ao-text-muted); font-size: var(--ao-fs-md); overflow-wrap: anywhere; }
#ao-root .ao-item-time { color: var(--ao-text-subtle); font-size: var(--ao-fs-sm); white-space: nowrap; }
#ao-root .ao-item-meta { display: flex; flex-wrap: wrap; align-items: center; gap: 6px; margin-top: 3px; }
#ao-root .ao-empty, #ao-root .ao-footnote { margin: 8px 0; color: var(--ao-text-subtle); text-align: center; }
#ao-root .ao-footnote { font-size: var(--ao-fs-sm); }

/* Etichette colorate */
#ao-root .ao-chip { display: inline-flex; align-items: center; padding: 1px 8px; border-radius: var(--ao-radius-pill); font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-medium); white-space: nowrap; }
#ao-root .ao-chip.ao-tone-neutral { background: var(--ao-line-light); color: var(--ao-text-muted); }
#ao-root .ao-chip.ao-tone-accent { background: var(--ao-accent-bg); color: var(--ao-accent); }
#ao-root .ao-chip.ao-tone-success { background: var(--ao-success-bg); color: var(--ao-success); }
#ao-root .ao-chip.ao-tone-warning { background: var(--ao-warning-bg); color: var(--ao-warning); }
#ao-root .ao-chip.ao-tone-danger { background: var(--ao-danger-bg); color: var(--ao-danger); }
#ao-root .ao-chip.ao-tone-hold { background: var(--ao-hold-bg); color: var(--ao-hold); }

/* Riquadri */
#ao-root .ao-card { display: grid; gap: 8px; padding: 12px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-white); box-shadow: var(--ao-shadow-card); }
#ao-root .ao-card-head { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
#ao-root .ao-card-title { margin: 0; font-size: var(--ao-fs-lg); font-weight: var(--ao-fw-medium); }
#ao-root .ao-notice { padding: 8px 10px; border-radius: var(--ao-radius-base); background: var(--ao-accent-bg); color: var(--ao-text-muted); font-size: var(--ao-fs-md); }
#ao-root .ao-notice.ao-tone-danger { background: var(--ao-danger-bg); color: var(--ao-danger); }
#ao-root .ao-muted { margin: 0; color: var(--ao-text-muted); }
#ao-root .ao-hint { color: var(--ao-text-subtle); font-size: var(--ao-fs-sm); }
#ao-root .ao-hint-accent { margin-top: -4px; color: var(--ao-accent); }
#ao-root .ao-hint-danger { color: var(--ao-danger); font-weight: var(--ao-fw-medium); }
#ao-root .ao-actions { display: flex; flex-wrap: wrap; gap: 8px; }
#ao-root .ao-actions .ao-primary { flex: 1; }

/* Mail aperta e presa in carico */
#ao-root .ao-mail-head { display: grid; gap: 2px; padding: 2px 2px 0; }
#ao-root .ao-mail-subject { font-size: var(--ao-fs-lg); font-weight: var(--ao-fw-bold); overflow-wrap: anywhere; }
#ao-root .ao-mail-from { display: flex; justify-content: space-between; gap: 8px; color: var(--ao-text-muted); font-size: var(--ao-fs-md); overflow-wrap: anywhere; }
#ao-root .ao-owner { display: flex; align-items: center; gap: 8px; padding: 8px 10px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-white); }
#ao-root .ao-owner-dot { width: 9px; height: 9px; flex: none; border-radius: 50%; background: var(--ao-line-strong); }
#ao-root .ao-owner-dot.ao-tone-success { background: var(--ao-success); }
#ao-root .ao-owner-dot.ao-tone-neutral { background: var(--ao-text-subtle); }
#ao-root .ao-owner-text { flex: 1; font-size: var(--ao-fs-md); }

/* Campi */
#ao-root .ao-field { display: grid; gap: 3px; min-width: 0; }
#ao-root .ao-label { color: var(--ao-text-muted); font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-medium); }
#ao-root .ao-input { width: 100%; min-height: var(--ao-control-height); padding: 5px 8px; border: 1px solid var(--ao-line-strong); border-radius: var(--ao-radius-base); background: var(--ao-white); color: var(--ao-text); font: inherit; }
#ao-root textarea.ao-input { resize: vertical; line-height: 1.4; }
#ao-root .ao-input:focus { border-color: var(--ao-accent); box-shadow: var(--ao-shadow-focus); outline: none; }
#ao-root .ao-grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
#ao-root .ao-favorites { display: grid; gap: 4px; }
#ao-root .ao-chips { display: flex; flex-wrap: wrap; gap: 5px; }
#ao-root .ao-fav { padding: 3px 9px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-pill); background: var(--ao-white); font-size: var(--ao-fs-sm); cursor: pointer; }
#ao-root .ao-fav:hover { border-color: var(--ao-accent-border); background: var(--ao-accent-bg); }
#ao-root .ao-fav.ao-active { border-color: var(--ao-accent); background: var(--ao-accent-bg); color: var(--ao-accent); font-weight: var(--ao-fw-medium); }
#ao-root .ao-fields { display: grid; grid-template-columns: auto 1fr; gap: 3px 10px; margin: 0; padding: 8px 10px; border-radius: var(--ao-radius-base); background: var(--ao-bg-soft); font-size: var(--ao-fs-md); }
#ao-root .ao-fields dt { color: var(--ao-text-muted); }
#ao-root .ao-fields dt:first-child { font-weight: var(--ao-fw-medium); color: var(--ao-text); }
#ao-root .ao-fields dd { margin: 0; overflow-wrap: anywhere; }

/* Ticket */
#ao-root .ao-ticket-head { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
#ao-root .ao-ticket-number { font-size: var(--ao-fs-lg); font-weight: var(--ao-fw-bold); color: var(--ao-accent); text-decoration: none; }
#ao-root .ao-ticket-number:hover { text-decoration: underline; }
#ao-root .ao-time { display: grid; gap: 4px; padding-top: 6px; border-top: 1px solid var(--ao-line-light); }
#ao-root .ao-time-row { display: flex; align-items: center; gap: 5px; }
#ao-root .ao-time-step { min-width: 40px; padding: 0 6px; }
#ao-root .ao-time-value { flex: 1; text-align: center; font-size: var(--ao-fs-lg); font-weight: var(--ao-fw-bold); }
#ao-root .ao-pending { display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 6px 10px; border-radius: var(--ao-radius-base); background: var(--ao-warning-bg); color: var(--ao-warning); font-size: var(--ao-fs-md); font-weight: var(--ao-fw-medium); }

/* Storico */
#ao-root .ao-history { display: grid; gap: 6px; margin: 0; padding: 0; list-style: none; }
#ao-root .ao-entry { padding: 7px 9px; border-left: 3px solid var(--ao-line-strong); border-radius: 0 var(--ao-radius-base) var(--ao-radius-base) 0; background: var(--ao-bg-soft); }
#ao-root .ao-entry-comment { border-left-color: var(--ao-accent); }
#ao-root .ao-entry-work_note { border-left-color: var(--ao-warning); background: var(--ao-warning-bg); }
#ao-root .ao-entry-head { display: flex; flex-wrap: wrap; gap: 6px; align-items: baseline; font-size: var(--ao-fs-sm); color: var(--ao-text-muted); }
#ao-root .ao-entry-kind { font-weight: var(--ao-fw-bold); color: var(--ao-text); }
#ao-root .ao-entry-text { margin-top: 2px; white-space: pre-wrap; overflow-wrap: anywhere; font-size: var(--ao-fs-md); }

/* Impostazioni */
#ao-root .ao-config-list { display: grid; gap: 4px; margin: 0; padding: 0; list-style: none; }
#ao-root .ao-config-item { display: flex; align-items: center; gap: 4px; padding: 6px 4px 6px 10px; border: 1px solid var(--ao-line-light); border-radius: var(--ao-radius-base); }
#ao-root .ao-config-item:hover { background: var(--ao-bg-soft); }
#ao-root .ao-config-text { flex: 1; display: grid; min-width: 0; }
#ao-root .ao-config-name { font-weight: var(--ao-fw-medium); }
#ao-root .ao-config-text .ao-hint { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
#ao-root .ao-editor { display: grid; gap: 8px; padding: 10px; border: 1px solid var(--ao-accent-border); border-radius: var(--ao-radius-base); background: var(--ao-accent-bg); }

/* Messaggi in fondo */
#ao-root .ao-status { position: sticky; bottom: 0; margin: 0 12px 12px; padding: 8px 10px; border-radius: var(--ao-radius-base); font-size: var(--ao-fs-md); font-weight: var(--ao-fw-medium); box-shadow: 0 2px 8px rgba(0,0,0,.12); }
#ao-root .ao-status.ao-tone-success { background: var(--ao-success-bg); color: var(--ao-success); }
#ao-root .ao-status.ao-tone-warning { background: var(--ao-warning-bg); color: var(--ao-warning); }
#ao-root .ao-status.ao-tone-danger { background: var(--ao-danger-bg); color: var(--ao-danger); }
`;
    function injectStyles() {
        if (document.getElementById('ao-styles')) return;
        document.head.append(h('style', { id: 'ao-styles' }, themeCssVariables() + STYLES));
    }

    // ======================================================================
    // 8. AVVIO E COLLEGAMENTO CON L'ADD-IN
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
        document.getElementById('open-amp-button')?.addEventListener('click', show);
        Outlook.onMailChanged(() => { handleMailChanged(); });
        Services.subscribe(() => { refreshRequests(); });
        state.mail = Outlook.currentMail();
        render();
        await refreshRequests();
        if (state.mail) await handleMailChanged();
    }
    function show() {
        initialize();
        const mainView = document.getElementById('main-view');
        const headerView = document.getElementById('header-view');
        const ampView = document.getElementById('amp-view');
        if (mainView) mainView.hidden = true;
        if (headerView) headerView.hidden = true;
        if (ampView) ampView.hidden = false;
        render();
    }

    window.AmpModule = Object.freeze({ initialize, show });
    const logInitError = error => console.error('AMP init error:', error);
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', () => initialize().catch(logInitError));
    else initialize().catch(logInitError);
})();
