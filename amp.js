/*
 * AMP - SNOW Assistant: area AMP dell'add-in Outlook.
 *
 * Nell'add-in i ticket si gestiscono da Outlook, senza aprire ServiceNow:
 *   - Coda         mail della casella condivisa che non hanno un ticket aperto;
 *   - Ticket       ticket aperti su ServiceNow, con l'assegnatario (Assigned to);
 *   - Mail         la mail aperta: ticket collegato (o da collegare / aprire), tempo,
 *                  storico di ServiceNow e azione da eseguire all'invio della risposta;
 *   - Impostazioni profilo, Agenzie (con i domini email), Attività, risposte salvate.
 *
 * Regole:
 *   - ServiceNow resta il riferimento; il pannello legge dai servizi la copia su SharePoint
 *     (stato, assegnatario = chi l'ha in carico, tempo, scadenze), allineata periodicamente.
 *     Il pannello non conserva dati dei ticket. Lo storico completo si legge su richiesta.
 *   - Ogni modifica va prima su ServiceNow; solo se riesce si aggiorna SharePoint
 *     (copia del ticket, thread in gestione / libero / scartato, registro dei movimenti).
 *   - In coda: un thread per conversazione, se non è in gestione. A ticket chiuso il thread
 *     torna libero e ricompare solo con un messaggio nuovo.
 *   - Mail e ticket si collegano per triangolazione: numero INC nel thread, collegamento
 *     registrato, oppure somiglianza (oggetto = Short description, arrivo = Reported) confermata
 *     dall'operatore. Lo stesso thread può avere più ticket nel tempo.
 *
 * Collegamento con l'add-in di jak (nessun suo file viene modificato):
 *   - espone window.AmpModule { initialize, show }; "Apri area AMP" chiama show();
 *   - la freccia in alto torna alla schermata principale (TestPluginAPI.showMainView);
 *   - legge la mail aperta da TestPluginAPI.getCurrentMessageContext / onMessageChanged.
 *
 * Tutto ciò che sta fuori dal pannello passa dal blocco 3 (Servizi). Oggi i servizi esistono
 * solo nell'anteprima (amp-preview-backend.js, simulati); nell'add-in vanno collegati alla
 * Function App (vedi amp-servizi.md).
 */
(function () {
    'use strict';

    // Indice dei blocchi
    //  0. Aspetto: tema, testi                 5. Interfaccia: struttura e navigazione
    //  1. Valori di ServiceNow                 6. Viste: Coda, Ticket, Mail, Impostazioni
    //  2. Configurazioni (salvate in locale)   7. Stili (usano i valori del blocco 0)
    //  3. Servizi e Outlook                    8. Avvio e collegamento con l'add-in
    //  4. Regole: agenzia suggerita, numeri INC, utente, date

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
            headerBackground: 'linear-gradient(135deg, #005a9e, #0078d4)',
            white: '#fff', // testo su sfondo blu (pulsanti, notifiche)
            surface: '#fff', // sfondo di card, campi, schede, barra di sistema
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
            lineLight: '#3b3a39',
            line: '#484644',
            lineStrong: '#605e5c',
            success: '#6ccb5f',
            successBg: '#1f3a1d',
            warning: '#f2c661',
            warningBg: '#3d3214',
            danger: '#ffb3b8',
            dangerBg: '#442726',
            hold: '#c3b2ff',
            holdBg: '#2f2a45'
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

    // Tutte le scritte del pannello. Le voci con (valore) sono funzioni.
    const TEXTS = {
        back: 'Torna all’add-in',
        system: {
            servicesOk: 'Servizi collegati',
            servicesMissing: 'Servizi non ancora collegati',
            serviceError: message => `Servizi non raggiungibili: ${message}`,
            synced: when => `Allineato ${when}`,
            syncHint: 'Ultimo allineamento della copia su SharePoint con ServiceNow'
        },
        tabs: { queue: 'In coda', tickets: 'Ticket', mail: 'Mail', settings: 'Impostazioni' },
        common: {
            me: 'me',
            back: view => `← ${view}`,
            loading: 'Caricamento…',
            cancel: 'Annulla',
            save: 'Salva',
            edit: 'Modifica',
            delete: 'Elimina',
            confirm: 'Confermi?',
            add: 'Aggiungi',
            none: '—',
            unassigned: 'Non assegnato',
            servicesMissing: 'I servizi non sono ancora collegati: qui compariranno i dati della casella condivisa e di ServiceNow.',
            serviceError: message => `Servizio non disponibile: ${message}`
        },
        queue: {
            empty: 'Nessuna richiesta in coda.',
            suggested: name => name,
            suggestedTitle: 'Agenzia suggerita',
            active: 'Attive',
            ignoredTab: 'Ignorate',
            ignoredEmpty: 'Nessun thread ignorato.',
            ignoredHint: 'Tornano da sole In coda quando arriva un messaggio nuovo nel thread; intanto puoi prenderle in carico da qui.',
            ignoredBy: (who, when) => `Ignorata da ${who} · ${when}`,

            takeCharge: 'Prendi in carico',
            ignore: 'Ignora'
        },
        tickets: {
            mine: 'Miei',
            all: 'Tutti',
            empty: 'Nessun ticket aperto.',
            noMail: 'senza mail',
            mailToConfirm: 'mail da confermare',
            newMessagesTitle: n => (n === 1 ? '1 messaggio nuovo nel thread' : `${n} messaggi nuovi nel thread`),
            threadGroup: n => `Stesso thread · ${n} ticket`,
            holdExpired: 'On Hold scaduto',
            holdUntil: date => `fino al ${date}`
        },
        mail: {
            takeBarWorking: 'Apro il ticket…',
            takeBarDone: 'Preso in carico',
            noMail: 'Apri una mail in Outlook, oppure scegli una richiesta dalla Coda o un ticket.',
            searching: 'Ricerca del ticket su ServiceNow…',
            openInOutlook: 'Apri la mail',
            // Collegamento con il ticket
            via: {
                inc: 'Collegato dal numero INC nel thread',
                link: 'Collegato (registro SharePoint)',
                confirmed: 'Collegato da un operatore'
            },
            ticketsOnThread: 'Ticket aperti relativi a questo thread',
            showDetails: 'Mostra dettagli',
            hideDetails: 'Nascondi dettagli',
            anotherTicket: '+ Apri un altro ticket su questo thread',
            candidatesTitle: 'Possibili ticket su ServiceNow',
            candidatesHint: 'Stesso oggetto e stessa ora di arrivo. Collega quello giusto, oppure apri un ticket nuovo qui sotto.',
            link: 'Collega',
            linked: number => `Mail collegata a ${number}.`,
            previousTitle: 'Ticket precedente in questo thread',
            previousHint: 'È arrivata una mail dopo la chiusura: è una richiesta nuova.',
            reportedAt: date => `Reported ${date}`,
            closedAt: date => `chiuso ${date}`,
            discardedBy: (who, when) => `Scartata da ${who} · ${when}`,
            restore: 'Rimetti in coda',
            // Presa in carico = Assigned to su ServiceNow
            assignedTo: name => `Assegnato a ${name}`,
            assignedToMe: name => `Assegnato a ${name} (me)`,
            notAssigned: 'Non assegnato',
            takeOver: 'Prendi in carico',
            taken: number => `${number} ora è assegnato a te.`,
            // Apertura del ticket
            newTicket: 'Presa in carico',
            newTicketHint: 'Apre il ticket su ServiceNow assegnato a te.',
            company: 'Agenzia',
            activity: 'Attività',
            choose: 'Seleziona…',
            suggestedFrom: reason => `Suggerita ${reason}`,
            byDomain: domain => `dal dominio @${domain}`,
            byText: 'dal testo della mail',
            favorites: 'Preferiti',
            shortDescription: 'Short description',
            fieldsPreview: 'Su ServiceNow',
            openTicket: 'Prendi in carico',
            discard: 'Ignora',
            profileMissing: 'Imposta prima il tuo profilo (Impostazioni → Profilo).',
            selectBoth: 'Scegli Agenzia e Attività.',
            ticketOpened: number => `Ticket ${number} aperto e assegnato a te.`,
            // Ticket
            openedAt: date => `Aperto ${date}`,
            openInServiceNow: 'Apri in ServiceNow',
            timeWorked: 'Tempo lavorato',
            minutes: n => `${n} min`,
            timeChanged: (delta, total) => `Tempo ${delta > 0 ? '+' : ''}${delta} min · totale ${total} min.`,
            closed: state => `Ticket ${state}: nessuna azione disponibile.`,
            noLinkedMail: 'Nessuna mail collegata: la risposta non è disponibile da qui.',
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
            reference: number => `Rif. ${number}`,
            includeReference: number => `Scrivi “Rif. ${number}” nella risposta`,
            prepareReply: 'Salva e rispondi',
            pending: label => `All’invio: ${label}`,
            clearPending: 'Togli',
            pendingSaved: 'Azione salvata: verrà eseguita all’invio della risposta.',
            replyOpened: 'Risposta aperta in Outlook.',
            // Storico
            history: 'Storico ServiceNow',
            loadHistory: 'Carica',
            loadMore: n => `Mostra precedenti (${n})`,
            historyEmpty: 'Nessuna attività.',
            kinds: { comment: 'Commento', work_note: 'Nota interna', system: 'Sistema' }
        },
        settings: {
            title: 'Impostazioni',
            tabs: { profile: 'Profilo', companies: 'Agenzie', activities: 'Attività', responses: 'Risposte' },
            profile: 'Profilo',
            profileHint: 'Il tuo nome come compare su ServiceNow: i ticket che apri vengono assegnati a te.',
            assignedTo: 'Assigned to',
            theme: 'Tema',
            themes: { auto: 'Automatico', light: 'Chiaro', dark: 'Scuro' },
            themeHint: 'Automatico segue il tema di Outlook. Chiaro e Scuro servono anche a verificare la leggibilità.',
            backup: 'Configurazioni',
            backupHint: 'Agenzie, Attività e risposte stanno in questo browser: esportale per averne una copia o portarle su un altro computer. Il file di Tampermonkey si importa senza modifiche.',
            companyName: 'Nome',
            company: 'Company',
            assignmentGroup: 'Assignment group',
            domains: 'Domini email',
            domainsHint: 'Separati da virgola, es. agenziaalfa.it, alfa.it',
            noDomains: 'nessuno: niente agenzia suggerita',
            category: 'Category',
            impactUrgency: 'Impact / Urgency',
            impact: 'Impact',
            urgency: 'Urgency',
            timeWorked: 'Tempo',
            timeWorkedMinutes: 'Minuti',
            prefix: 'Prefisso',
            prefixLong: 'Prefisso Short description',
            responseTitle: 'Titolo',
            response: 'Testo',
            empty: 'Nessuna voce.',
            toVerify: 'da verificare',
            toVerifyHint: 'Nome non ancora collegato a ServiceNow: verrà verificato dal servizio.',
            referenceHint: 'Il nome come compare su ServiceNow.',
            import: 'Importa JSON',
            export: 'Esporta JSON',
            imported: 'Configurazioni importate.',
            importInvalid: 'File JSON non valido.',
            saved: 'Salvato.',
            deleted: 'Eliminato.',
            nameRequired: 'Il nome è obbligatorio.',
            version: build => `Versione del pannello: ${build}`
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
    // Versione del pannello (si vede in Impostazioni → Profilo): serve a capire se il browser usa l'ultima.
    const BUILD = '2026-10-30 16:35';

    function createDefaultConfig() {
        const now = new Date().toISOString();
        return {
            version: STORAGE.version, created: now, updated: now,
            profile: { assignedTo: { value: '', displayValue: '' }, tag: '' },
            companies: {}, activities: {}, favorites: [], recent: [],
            combinationResponses: {}, usefulResponses: {},
            preferences: { favoritesGroupBy: 'none', theme: 'auto' }
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
    // Tema del pannello: auto (come Outlook), light, dark.
    function applyTheme(config = loadConfig()) {
        if (!root) return;
        const theme = config.preferences?.theme;
        if (theme === 'light' || theme === 'dark') root.dataset.theme = theme;
        else delete root.dataset.theme;
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
    // Unico punto in cui il pannello parla con l'esterno. Ogni servizio restituisce una Promise.
    //
    //   listQueue()                         → [{ conversationId, messageId, subject, sender, receivedAt,
    //                                            preview, previous: { number, state } | null, candidates: [numero] }]
    //   listIgnored()                       → [{ conversationId, messageId, subject, sender, receivedAt, discardedBy, discardedAt }]
    //   markThreadSeen(conversationId, utente) → i messaggi del thread non sono più "nuovi" per chi l'ha in carico
    //   listOpenTickets()                   → Ticket[]  (In Progress e On Hold del gruppo, copia su SharePoint)
    //   findTicketForMail(mail)             → { tickets: (Ticket & { via: 'inc'|'link'|'confirmed' })[], previous, candidates: Ticket[], discarded }
    //   getTicket(numero)                   → Ticket
    //   linkTicket(conversationId, numero, utente)  → come findTicketForMail (l'operatore conferma il collegamento)
    //   createTicket(mail, { companyId, companyName, activityId, activityName, fields }) → Ticket
    //   assignTicket(numero, utente)        → Ticket    (Assigned to = utente)
    //   addTime(numero, minuti, nome)       → Ticket    (minuti anche negativi)
    //   getTicketActivity(numero, { offset, limit }) → { entries: [{ at, author, kind, text }], total }
    //   setPendingAction(numero, azione | null)      → Ticket   (eseguita all'invio della risposta)
    //   discardMail(conversationId, utente) / restoreMail(conversationId)
    //   subscribe(listener)                 → funzione per smettere (facoltativo)
    //
    // mail   = { conversationId, itemId, subject, sender: { name, email }, receivedAt, body, incNumbers }
    // utente = { displayValue, value: sys_id }   (il profilo del pannello)
    // Ticket = { number, sysId, url, state, assignedTo: { name, sysId }, companyName, activityName,
    //            shortDescription, reportedOn, openedAt, closedAt, timeWorkedMinutes, holdReason,
    //            holdExpiration, conversationId, messageId, pendingAction, owner, newMessages,
    //            lastNewMessage: { from, at, preview } | null, syncedAt }
    // ======================================================================

    const SERVICE_METHODS = ['listQueue', 'listIgnored', 'markThreadSeen', 'listOpenTickets', 'findTicketForMail', 'getTicket', 'linkTicket', 'createTicket',
        'assignTicket', 'addTime', 'getTicketActivity', 'setPendingAction', 'discardMail', 'restoreMail'];

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
        currentMail() {
            const context = window.TestPluginAPI?.getCurrentMessageContext?.();
            if (!context) return null;
            const subject = String(context.subject || '');
            const body = String(context.bodyText || '');
            return {
                itemId: context.itemId || null,
                conversationId: context.normalizedConversationId || context.conversationId || null,
                subject,
                sender: { name: context.sender?.displayName || '', email: context.sender?.emailAddress || '' },
                receivedAt: context.dateTimeCreated || null,
                body,
                incNumbers: extractIncNumbers(`${subject}\n${body}`),
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
    // 4. REGOLE: AGENZIA SUGGERITA, NUMERI INC, UTENTE, DATE
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
    // Numeri di Incident citati in un testo, nell'ordine in cui compaiono.
    function extractIncNumbers(text) {
        return [...new Set((String(text || '').match(/\bINC\d{7}\b/gi) || []).map(item => item.toUpperCase()))];
    }
    // Utente del pannello: il profilo, altrimenti il nome dell'utente di Outlook.
    function currentUser(config = loadConfig()) {
        const profile = config.profile.assignedTo;
        if (profile.displayValue) return profile;
        const name = state.mail?.userName || window.TestPluginAPI?.getCurrentMessageContext?.()?.mailbox?.userDisplayName || '';
        return { displayValue: name, value: '' };
    }
    function isMine(ticket, user = currentUser()) {
        const assignee = ticket?.assignedTo;
        if (!assignee || !user.displayValue) return false;
        if (assignee.sysId && user.value) return assignee.sysId === user.value;
        return String(assignee.name || '').trim().toLowerCase() === user.displayValue.trim().toLowerCase();
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

    // Solo stato dell'interfaccia: i dati dei ticket arrivano ogni volta dai servizi.
    const state = {
        view: 'queue', // queue | tickets | mail | settings
        queue: [],
        tickets: [],
        loaded: false,
        loading: false,
        serviceError: '',
        ticketFilter: 'mine',
        mail: null, // mail aperta in Outlook
        target: null, // cosa mostra la vista Mail: { kind: 'mail', mail } | { kind: 'ticket', number }
        result: null, // triangolazione della mail mostrata
        searching: false,
        ticketCache: {}, // ultimo Ticket letto per numero (anche chiusi o non in elenco)
        ticketDrafts: {}, // per conversazione: Agenzia, Attività, Short description
        actionDrafts: {}, // per ticket: azione all'invio
        history: {}, // per ticket: storico caricato
        editing: null, // impostazioni: { kind, id }
        settingsTab: 'profile', // profile | companies | activities | responses
        backTo: null, // vista a cui torna il pulsante ← (Mail e Impostazioni aperte da un'altra vista)
        ignored: [], // archivio dei thread ignorati (In coda, in fondo)
        queueTab: 'active', // In coda: active (da prendere) | ignored (archivio)
        extraTicketFor: null, // thread su cui si sta aprendo un ticket in più
        expandedTickets: [], // ticket con i dettagli aperti (thread con più ticket)
        barOffset: 0, // di quanti px la barra secondaria è scivolata via (segue lo scorrimento)
        queueExpanded: null,
        swipeOpen: null, // conversazione In coda con il tasto Ignora scoperto (scheda spostata a sinistra)
        animateExpand: null, // modulo appena aperto, da animare
        holdRender: false, // true mentre una card esce con l'animazione // conversazione con il modulo di presa in carico aperto in In coda
        busy: false
    };
    let root = null;
    let statusTimer = null;
    let searchSequence = 0;

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
    // Solo la notifica: numero su sfondo blu.
    const newsBadge = (news, title = '') => h('span', { class: 'ao-news', title: title || undefined }, news);
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
    function checkboxField(label, checked, onchange) {
        return h('label', { class: 'ao-check' },
            h('input', { type: 'checkbox', checked, onchange: event => onchange(event.target.checked) }),
            h('span', {}, label));
    }
    // Pulsante che chiede conferma con un secondo clic (confirm() non esiste negli add-in Office).
    function confirmButton(label, className, action) {
        const button = h('button', { type: 'button', class: className }, label);
        let armed = null;
        button.addEventListener('click', () => {
            if (armed) {
                clearTimeout(armed);
                armed = null;
                action(button);
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
    // Un Ticket appena letto dai servizi aggiorna elenco, risultato della mail e cache.
    function rememberTicket(ticket) {
        if (!ticket?.number) return;
        state.ticketCache[ticket.number] = ticket;
        const index = state.tickets.findIndex(item => item.number === ticket.number);
        if (index >= 0) {
            if (OPEN_STATES.includes(ticket.state)) state.tickets[index] = ticket;
            else state.tickets.splice(index, 1);
        }
        const inResult = state.result?.tickets?.findIndex(item => item.number === ticket.number) ?? -1;
        if (inResult >= 0) state.result.tickets[inResult] = { ...ticket, via: state.result.tickets[inResult].via };
    }
    const ticketByNumber = number => state.tickets.find(item => item.number === number) || state.ticketCache[number] || null;

    // Elenchi di Coda e Ticket, sempre dai servizi.
    async function refreshLists() {
        if (!Services.available()) {
            state.loaded = true;
            render();
            return;
        }
        state.loading = !state.loaded;
        if (state.loading) render();
        try {
            const [queue, ignored, tickets] = await Promise.all([Services.listQueue(), Services.listIgnored(), Services.listOpenTickets()]);
            state.queue = queue;
            state.ignored = ignored;
            state.tickets = tickets;
            tickets.forEach(ticket => { state.ticketCache[ticket.number] = ticket; });
            state.serviceError = '';
        } catch (error) {
            state.serviceError = error?.message || String(error);
        }
        state.loaded = true;
        state.loading = false;
        render();
        await refreshTarget(true);
        // Lo storico aperto si ricarica per mostrare le novità (es. risposta appena inviata).
        Object.keys(state.history).forEach(number => {
            if (state.history[number].open) loadHistory(number, true);
        });
    }
    // Triangolazione per la mail mostrata nella vista Mail.
    async function refreshTarget(quiet = false) {
        if (state.target?.kind !== 'mail' || !Services.available()) return;
        const sequence = ++searchSequence;
        if (!quiet) {
            state.searching = true;
            render();
        }
        try {
            const result = await Services.findTicketForMail(state.target.mail);
            if (sequence !== searchSequence) return;
            state.result = result;
            [...(result.tickets || []), result.previous, ...(result.candidates || [])].forEach(ticket => { if (ticket) state.ticketCache[ticket.number] = ticket; });
        } catch (error) {
            if (sequence !== searchSequence) return;
            state.result = null;
            showStatus(error?.message || String(error), 'danger');
        }
        state.searching = false;
        render();
    }

    // Mail aperta in Outlook: la vista Mail cerca il suo ticket.
    function handleMailChanged() {
        state.mail = Outlook.currentMail();
        if (!state.mail?.conversationId) return render();
        showMailTarget(state.mail);
    }
    function showMailTarget(mail) {
        const sameConversation = state.target?.kind === 'mail' && state.target.mail.conversationId === mail.conversationId;
        state.target = { kind: 'mail', mail };
        if (!sameConversation) state.result = null;
        const entering = state.view !== 'settings' && state.view !== 'mail';
        if (entering) {
            state.backTo ||= state.view;
            state.view = 'mail';
        }
        render();
        if (entering) slideInView(1);
        refreshTarget(sameConversation);
    }
    // Dalla Coda: mostra la richiesta e apre la mail in Outlook.
    function openQueueItem(item) {
        state.backTo = 'queue';
        showMailTarget({
            conversationId: item.conversationId, itemId: item.messageId, subject: item.subject, sender: item.sender,
            receivedAt: item.receivedAt, body: item.preview, incNumbers: []
        });
        scrollToTop();
        Outlook.openMessage(item.messageId);
    }
    // Dai Ticket: mostra il ticket e, se ha una mail collegata, la apre.
    // Solo questo clic segna come lette le risposte nuove (aprire la mail in Outlook non basta).
    function openTicketItem(ticket) {
        if (ticket.newMessages && ticket.conversationId) Services.markThreadSeen(ticket.conversationId, currentUser()).catch(() => {});
        state.backTo = 'tickets';
        state.target = { kind: 'ticket', number: ticket.number };
        const entering = state.view !== 'mail';
        state.view = 'mail';
        render();
        scrollToTop();
        if (entering) slideInView(1);
        if (ticket.messageId) Outlook.openMessage(ticket.messageId);
    }
    // backTo: vista a cui tornare con il pulsante ← (le schede lo azzerano).
    function setView(view, backTo = null) {
        const from = state.view;
        state.barOffset = 0;
        state.view = view;
        state.backTo = backTo;
        state.editing = null;
        render();
        scrollToTop();
        if (from !== view) slideInView(VIEW_ORDER.indexOf(view) >= VIEW_ORDER.indexOf(from) ? 1 : -1);
    }
    // Cambio di scheda: il contenuto nuovo entra nella direzione dello spostamento
    // (da destra andando verso destra, da sinistra tornando indietro), così la navigazione si nota.
    const VIEW_ORDER = ['queue', 'tickets', 'mail', 'settings'];
    function slideInView(direction) {
        // Si anima il contenitore fisso (.ao-body): il contenuto può ridisegnarsi durante l'animazione.
        const body = root?.querySelector('.ao-body');
        if (!body?.animate || window.matchMedia?.('(prefers-reduced-motion: reduce)').matches) return;
        body.animate([
            { opacity: 0, transform: `translateX(${24 * direction}px)` },
            { opacity: 1, transform: 'none' }
        ], { duration: 280, easing: 'cubic-bezier(.2, .8, .2, 1)' });
    }
    // Cambiando vista si riparte dall'alto (se l'area è già scrollata oltre le schede).
    function scrollToTop() {
        const scroll = root?.querySelector('.ao-scroll');
        if (scroll) scroll.scrollTop = 0;
    }

    function buildShell(container) {
        // Una sola barra: freccia per tornare all'add-in, schede, ingranaggio delle impostazioni.
        root = h('div', { id: 'ao-root' },
            h('nav', { class: 'ao-tabs', role: 'tablist' },
                h('button', { type: 'button', class: 'ao-icon-button ao-back', title: TEXTS.back, 'aria-label': TEXTS.back, onclick: () => window.TestPluginAPI?.showMainView?.() }, '←'),
                // La vista di lavoro del ticket non è una scheda: si apre da sola (mail o ticket) e ha il suo "←".
                ['queue', 'tickets'].map(view => h('button', {
                    type: 'button', role: 'tab', class: 'ao-tab', dataset: { view },
                    onclick: () => setView(view)
                }, h('span', { class: 'ao-tab-label' }, TEXTS.tabs[view]), counter(0))),
                h('button', {
                    type: 'button', class: 'ao-icon-button ao-settings-button', dataset: { view: 'settings' },
                    title: TEXTS.tabs.settings, 'aria-label': TEXTS.tabs.settings, onclick: () => (state.view === 'settings' ? setView(state.backTo || 'queue') : setView('settings', state.view))
                }, '⚙')),
            h('div', { class: 'ao-body' }),
            h('div', { class: 'ao-status', role: 'status', 'aria-live': 'polite', hidden: true }),
            // Barra di sistema fissa in fondo: stato dei servizi e ultimo allineamento, sempre visibile.
            h('footer', { class: 'ao-statusbar', role: 'status', 'aria-live': 'polite' }));
        // Con il tasto Ignora scoperto, un clic altrove (anche sulla scheda) o Esc la rimette a posto.
        root.addEventListener('click', event => {
            if (!state.swipeOpen || event.target.closest('.ao-swipe-action, .ao-ignore')) return;
            const onCard = event.target.closest('.ao-queue-item.ao-swiped');
            closeSwipe();
            if (onCard) { event.preventDefault(); event.stopPropagation(); }
        }, true);
        root.addEventListener('keydown', event => { if (event.key === 'Escape') closeSwipe(); });
        container.replaceChildren(root);
    }

    function render() {
        if (!root || state.holdRender) return;
        const config = loadConfig();
        const user = currentUser(config);
        // Thread dei miei ticket con messaggi nuovi (uno per thread).
        const newReplies = new Set(state.tickets.filter(ticket => ticket.newMessages && isMine(ticket, user)).map(ticket => ticket.conversationId)).size;
        const counts = { queue: state.queue.length, tickets: state.tickets.filter(ticket => isMine(ticket, user)).length };
        root.querySelectorAll('.ao-tab').forEach(tab => {
            // Nella vista di lavoro resta evidenziata la scheda da cui si è arrivati.
            const active = tab.dataset.view === (state.view === 'mail' ? (state.backTo || 'tickets') : state.view);
            tab.classList.toggle('ao-active', active);
            tab.setAttribute('aria-selected', String(active));
            // Conteggio sempre; sfondo blu se c'è una notifica (Ticket: messaggi nuovi nei thread dei miei ticket).
            const notify = tab.dataset.view === 'tickets' && newReplies > 0;
            // Aggiornato sul posto (non ricreato), così segue lo zoom animato insieme alla scritta.
            const tabCounter = tab.querySelector('.ao-counter');
            const count = counts[tab.dataset.view];
            tabCounter.textContent = count || '';
            tabCounter.hidden = !count && !notify;
            tabCounter.classList.toggle('ao-counter-notify', notify);
            if (notify) tabCounter.title = TEXTS.tickets.newMessagesTitle(newReplies);
            else tabCounter.removeAttribute('title');
        });
        root.querySelector('.ao-settings-button').classList.toggle('ao-active', state.view === 'settings');
        renderStatusBar();
        const views = { queue: renderQueue, tickets: renderTickets, mail: renderMail, settings: renderSettings };
        const body = root.querySelector('.ao-body');
        // Chi resta fermo (barra secondaria, "←") sta in alto, separato da una linea; tutto il resto scorre
        // sotto, a tutta larghezza. La posizione di scorrimento resta uguale tra un aggiornamento e l'altro.
        const scrollKey = `${state.view}|${state.target?.kind || ''}|${state.target?.number || state.target?.mail?.conversationId || ''}|${state.queueTab}|${state.ticketFilter}|${state.settingsTab}`;
        const previousScroll = body.dataset.scrollKey === scrollKey ? body.querySelector('.ao-scroll')?.scrollTop || 0 : 0;
        const view = views[state.view](config, user);
        const children = [...view.children];
        const fixedCount = children.findIndex(child => !child.matches('.ao-segmented, .ao-back-row'));
        const fixed = fixedCount < 0 ? children : children.slice(0, fixedCount);
        const scroll = h('div', { class: 'ao-scroll' }, ...(fixedCount < 0 ? [] : children.slice(fixedCount)));
        if (!previousScroll) state.barOffset = 0;
        // La barra è il primo elemento dell'area che scorre: la sua "casa" è in cima all'elenco.
        const bar = fixed.length ? h('div', { class: 'ao-fixed' }, ...fixed) : null;
        if (bar) scroll.prepend(bar);
        view.replaceChildren(scroll);
        body.replaceChildren(view);
        body.dataset.scrollKey = scrollKey;
        scroll.scrollTop = previousScroll;
        if (bar) watchScroll(scroll, bar);
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

    // Gli avvisi sui servizi stanno nella barra di sistema in fondo: le viste non li ripetono.
    function serviceNotice() {
        return null;
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
        const syncedAt = state.tickets.find(ticket => ticket.syncedAt)?.syncedAt;
        bar.className = `ao-statusbar ao-statusbar-${tone}`;
        bar.replaceChildren(...[
            h('span', { class: 'ao-statusbar-dot', 'aria-hidden': 'true' }),
            h('span', { class: 'ao-statusbar-text', title: message }, message),
            syncedAt ? h('span', { class: 'ao-statusbar-sync', title: TEXTS.system.syncHint }, TEXTS.system.synced(formatWhen(syncedAt))) : null
        ].filter(Boolean));
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
            } else place(value, false);
            thumb.classList.add('ao-ready');
            // Se il riquadro cambia larghezza, la pillola segue la voce (senza animazione).
            if (window.ResizeObserver) new ResizeObserver(() => { if (!thumb.classList.contains('ao-moving')) place(value, false); }).observe(bar);
            thumb.addEventListener('transitionend', () => thumb.classList.remove('ao-moving'));
        });
        return bar;
    }
    // "←" della vista di lavoro: torna all'elenco da cui si è arrivati (di base Ticket).
    function backRow() {
        const target = state.backTo || 'tickets';
        return h('div', { class: 'ao-back-row' },
            h('button', { type: 'button', class: 'ao-link ao-back-link', onclick: () => setView(target) }, TEXTS.common.back(TEXTS.tabs[target])));
    }
    function clickable(className, onopen, ...children) {
        return h('li', { class: className, tabindex: '0', role: 'button', onclick: onopen, onkeydown: event => { if (event.key === 'Enter') onopen(); } }, ...children);
    }

    // ======================================================================
    // 6. VISTE
    // ======================================================================

    // ---------------------------- Coda ------------------------------------
    // Coda condivisa con barra secondaria: Attive (da prendere in carico) | Ignorate (archivio).
    function renderQueue(config) {
        const notice = serviceNotice();
        const active = state.queue.slice().sort((a, b) => String(b.receivedAt).localeCompare(String(a.receivedAt)));
        const showIgnored = state.queueTab === 'ignored';
        return h('div', { class: 'ao-view' },
            notice,
            filterBar([
                { value: 'active', label: TEXTS.queue.active, count: active.length },
                { value: 'ignored', label: TEXTS.queue.ignoredTab, count: state.ignored.length }
            ], showIgnored ? 'ignored' : 'active', value => { state.queueTab = value; state.queueExpanded = null; render(); }),
            showIgnored
                ? [
                    state.ignored.length ? h('ul', { class: 'ao-list' }, state.ignored.map(item => ignoredItem(item)))
                        : (servicesBlocked() ? null : h('p', { class: 'ao-empty' }, TEXTS.queue.ignoredEmpty)),
                    // Spiegazione in fondo, come nota.
                    state.ignored.length ? h('p', { class: 'ao-footnote' }, TEXTS.queue.ignoredHint) : null
                ]
                : (active.length ? h('ul', { class: 'ao-list' }, active.map(item => queueItem(item, config))) : (servicesBlocked() ? null : h('p', { class: 'ao-empty' }, TEXTS.queue.empty))));
    }
    // Thread ignorato: si può prendere in carico direttamente (stesso modulo compatto della coda).
    function ignoredItem(item) {
        const config = loadConfig();
        const mail = mailFromQueueItem(item);
        const expanded = state.queueExpanded === item.conversationId;
        return h('li', { class: `ao-item ao-queue-item${expanded ? ' ao-expanded' : ''}` },
            h('div', {
                class: 'ao-item-main', tabindex: '0', role: 'button',
                onclick: () => openQueueItem(item), onkeydown: event => { if (event.key === 'Enter') openQueueItem(item); }
            },
                // Stessa disposizione di Attive: oggetto e mittente a sinistra, ora a destra.
                h('div', { class: 'ao-item-left' },
                    h('span', { class: 'ao-item-title' }, item.subject || '(senza oggetto)'),
                    h('div', { class: 'ao-item-sub' }, item.sender?.name || item.sender?.email || '')),
                h('div', { class: 'ao-item-right' },
                    h('span', { class: 'ao-item-time' }, formatWhen(item.receivedAt)))),
            expanded
                ? h('div', { class: `ao-inline-form${state.animateExpand === item.conversationId ? ' ao-animate-expand' : ''}` }, ticketForm(mail, config, { compact: true }),
                    h('div', { class: 'ao-actions' },
                        openTicketButton(mail, config),
                        h('button', { type: 'button', class: 'ao-button ao-subtle', onclick: event => closeTakeCharge(event.currentTarget) }, TEXTS.common.cancel)))
                : h('div', { class: 'ao-item-actions' },
                    h('button', { type: 'button', class: 'ao-button ao-small ao-take', onclick: () => openTakeCharge(item.conversationId) }, TEXTS.queue.takeCharge),
                    // Chi l'ha ignorata, in basso a destra sulla stessa riga del pulsante.
                    h('span', { class: 'ao-hint ao-ignored-by' }, TEXTS.queue.ignoredBy(item.discardedBy || TEXTS.common.none, formatWhen(item.discardedAt)))));
    }
    // Prendi in carico: apre il modulo nella card con un'animazione (solo all'apertura, non a ogni aggiornamento).
    function openTakeCharge(conversationId) {
        state.queueExpanded = conversationId;
        state.animateExpand = conversationId;
        render();
        state.animateExpand = null;
        // L'animazione usa l'altezza vera del modulo, così la velocità è uniforme.
        const form = root.querySelector('.ao-animate-expand');
        if (form) form.style.setProperty('--ao-form-height', `${form.scrollHeight}px`);
    }
    // Annulla: il modulo si richiude con l'animazione inversa, poi la card torna normale.
    function closeTakeCharge(button) {
        const form = button.closest('.ao-inline-form');
        const finish = () => {
            if (!state.queueExpanded) return;
            state.queueExpanded = null;
            render();
        };
        if (!form || window.matchMedia?.('(prefers-reduced-motion: reduce)').matches) return finish();
        form.closest('.ao-expanded')?.classList.remove('ao-expanded'); // il bordo blu va via subito
        form.style.setProperty('--ao-form-height', `${form.scrollHeight}px`);
        form.classList.remove('ao-animate-expand');
        form.classList.add('ao-animate-collapse');
        form.addEventListener('animationend', finish, { once: true });
        setTimeout(finish, 300); // riserva, se l'animazione non parte
    }
    function mailFromQueueItem(item) {
        return {
            conversationId: item.conversationId, itemId: item.messageId, subject: item.subject, sender: item.sender,
            receivedAt: item.receivedAt, body: item.preview, incNumbers: []
        };
    }
    // Ogni riga: clic sul testo = apre la mail; Prendi in carico = modulo compatto; Ignora = scarta.
    function queueItem(item, config) {
        const suggestion = suggestCompany(config, item.sender, item.subject, item.preview);
        const company = suggestion && config.companies[suggestion.id];
        const expanded = state.queueExpanded === item.conversationId;
        const mail = mailFromQueueItem(item);
        const swiped = !expanded && state.swipeOpen === item.conversationId;
        return h('li', { class: `ao-item ao-queue-item${expanded ? ' ao-expanded' : ''}${swiped ? ' ao-swiped' : ''}`, dataset: { conversation: item.conversationId } },
            expanded ? null : swipeAction(mail),
            h('div', {
                class: 'ao-item-main', tabindex: '0', role: 'button',
                onclick: () => openQueueItem(item), onkeydown: event => { if (event.key === 'Enter') openQueueItem(item); }
            },
                // A sinistra oggetto e mittente; a destra ora di arrivo ed etichette, una sotto l'altra.
                h('div', { class: 'ao-item-left' },
                    h('span', { class: 'ao-item-title' }, item.subject || '(senza oggetto)'),
                    h('div', { class: 'ao-item-sub' }, item.sender?.name || item.sender?.email || '')),
                h('div', { class: 'ao-item-right' },
                    h('span', { class: 'ao-item-time' }, formatWhen(item.receivedAt)),
                    company ? chip(TEXTS.queue.suggested(company.name), 'accent', TEXTS.queue.suggestedTitle) : null)),
            expanded
                ? h('div', { class: `ao-inline-form${state.animateExpand === item.conversationId ? ' ao-animate-expand' : ''}` }, ticketForm(mail, config, { compact: true }),
                    h('div', { class: 'ao-actions' },
                        openTicketButton(mail, config),
                        h('button', { type: 'button', class: 'ao-button ao-subtle', onclick: event => closeTakeCharge(event.currentTarget) }, TEXTS.common.cancel)))
                : h('div', { class: 'ao-item-actions' },
                    h('button', { type: 'button', class: 'ao-button ao-small ao-take', onclick: () => openTakeCharge(item.conversationId) }, TEXTS.queue.takeCharge),
                    h('button', { type: 'button', class: 'ao-button ao-small ao-subtle ao-ignore', onclick: event => toggleSwipe(event.currentTarget.closest('.ao-queue-item'), item.conversationId) }, TEXTS.mail.discard)));
    }

    // ---------------------------- Ticket ----------------------------------
    // I ticket dello stesso thread stanno insieme; in cima i thread con messaggi nuovi.
    function renderTickets(config, user) {
        const notice = serviceNotice();
        const open = state.tickets.slice();
        const mine = open.filter(ticket => isMine(ticket, user));
        const items = state.ticketFilter === 'mine' ? mine : open;
        const groups = [];
        items.forEach(ticket => {
            const group = ticket.conversationId && groups.find(item => item.conversationId === ticket.conversationId);
            if (group) group.tickets.push(ticket);
            else groups.push({ conversationId: ticket.conversationId, tickets: [ticket] });
        });
        const rank = group => [
            group.tickets.some(ticket => ticket.newMessages) ? 1 : 0,
            group.tickets.some(isHoldExpired) ? 1 : 0,
            group.tickets.map(ticket => String(ticket.openedAt)).sort().slice(-1)[0]
        ];
        groups.sort((a, b) => {
            const [newA, holdA, dateA] = rank(a);
            const [newB, holdB, dateB] = rank(b);
            return newB - newA || holdB - holdA || dateB.localeCompare(dateA);
        });
        return h('div', { class: 'ao-view' },
            notice,
            filterBar([
                { value: 'mine', label: TEXTS.tickets.mine, count: mine.length },
                { value: 'all', label: TEXTS.tickets.all, count: open.length }
            ], state.ticketFilter, value => { state.ticketFilter = value; render(); }),
            groups.length ? h('ul', { class: 'ao-list' }, groups.map(group => (group.tickets.length > 1
                ? h('li', { class: `ao-group${group.tickets.some(ticket => ticket.newMessages) ? ' ao-group-new' : ''}` },
                    h('div', { class: 'ao-group-head' }, TEXTS.tickets.threadGroup(group.tickets.length)),
                    newMessagesPreview(group.tickets[0]),
                    h('ul', { class: 'ao-list' }, group.tickets.map(ticket => ticketItem(ticket, user, { inGroup: true }))))
                : ticketItem(group.tickets[0], user))))
                : (servicesBlocked() ? null : h('p', { class: 'ao-empty' }, TEXTS.tickets.empty)));
    }
    // Messaggi nuovi (numero su sfondo blu) e inizio dell'ultimo arrivato nel thread (chi scrive, quando, prime righe).
    function newMessagesPreview(ticket) {
        if (!ticket.newMessages) return null;
        const last = ticket.lastNewMessage;
        return h('div', { class: 'ao-new-preview', title: TEXTS.tickets.newMessagesTitle(ticket.newMessages) },
            h('div', { class: 'ao-new-head' },
                newsBadge(ticket.newMessages),
                last ? h('span', { class: 'ao-new-from' }, last.from) : null,
                last ? h('span', { class: 'ao-item-time' }, formatWhen(last.at)) : null),
            last?.preview ? h('div', { class: 'ao-new-text' }, last.preview) : null);
    }
    function ticketItem(ticket, user, { inGroup = false } = {}) {
        const mine = isMine(ticket, user);
        const holdDate = ticket.state === '3' ? parseDate(ticket.holdExpiration) : null;
        const highlight = ticket.newMessages && !inGroup ? ' ao-item-new' : '';
        return clickable(`ao-item${isHoldExpired(ticket) ? ' ao-item-alert' : ''}${highlight}`, () => openTicketItem(ticket),
            h('div', { class: 'ao-item-top' },
                h('span', { class: 'ao-item-number' }, ticket.number),
                stateChip(ticket)),
            inGroup ? null : newMessagesPreview(ticket),
            h('div', { class: 'ao-item-title' }, ticket.shortDescription),
            h('div', { class: 'ao-item-sub' }, [ticket.companyName, ticket.activityName].filter(Boolean).join(' · ')),
            h('div', { class: 'ao-item-meta' },
                chip(mine ? TEXTS.common.me : (ticket.assignedTo?.name || TEXTS.common.unassigned), mine ? 'success' : 'neutral'),
                h('span', { class: 'ao-item-time' }, TEXTS.mail.minutes(ticket.timeWorkedMinutes || 0)),
                holdDate ? h('span', { class: 'ao-item-time' }, TEXTS.tickets.holdUntil(formatWhen(holdDate))) : null,
                ticket.messageId ? (ticket.mailUnconfirmed ? chip(TEXTS.tickets.mailToConfirm, 'warning') : null) : chip(TEXTS.tickets.noMail, 'neutral'),
                ticket.pendingAction ? chip(pendingLabel(ticket.pendingAction), 'warning') : null));
    }

    // ---------------------------- Mail ------------------------------------
    function renderMail(config, user) {
        const notice = serviceNotice();
        const target = state.target;
        if (!target) return h('div', { class: 'ao-view' }, backRow(), notice, h('p', { class: 'ao-empty' }, TEXTS.mail.noMail));
        if (target.kind === 'ticket') {
            const ticket = ticketByNumber(target.number);
            return h('div', { class: 'ao-view' }, backRow(), notice,
                ticket ? ticketSection(ticket, config, user, { isCurrentMail: false, showOpen: true }) : h('p', { class: 'ao-muted' }, TEXTS.common.loading));
        }
        const mail = target.mail;
        const isCurrentMail = state.mail?.conversationId === mail.conversationId;
        const result = state.result;
        const content = [];
        if (!result) content.push(h('p', { class: 'ao-muted' }, Services.available() ? TEXTS.mail.searching : ''));
        else if (result.tickets?.length) {
            const openCount = result.tickets.filter(ticket => OPEN_STATES.includes(ticket.state)).length;
            if (openCount > 1) {
                // Più ticket sullo stesso thread: avviso sintetico e solo i ticket; i dettagli si aprono con un clic.
                content.push(h('div', { class: 'ao-multi' }, h('span', {}, TEXTS.mail.ticketsOnThread), h('b', { class: 'ao-multi-count' }, openCount)));
                result.tickets.forEach(found => {
                    const ticket = { ...found, ...(ticketByNumber(found.number) || {}) };
                    const expanded = state.expandedTickets.includes(ticket.number);
                    const toggle = () => {
                        state.expandedTickets = expanded ? state.expandedTickets.filter(number => number !== ticket.number) : [...state.expandedTickets, ticket.number];
                        render();
                    };
                    content.push(h('section', { class: `ao-ticket-compact${expanded ? ' ao-expanded' : ''}` },
                        h('div', { class: 'ao-ticket-compact-head', tabindex: '0', role: 'button', 'aria-expanded': String(expanded), onclick: toggle, onkeydown: event => { if (event.key === 'Enter') toggle(); } },
                            h('div', { class: 'ao-ticket-head' },
                                h('div', { class: 'ao-item-left' },
                                    h('span', { class: 'ao-item-number' }, ticket.number),
                                    h('div', { class: 'ao-item-title' }, ticket.shortDescription)),
                                ticketTags(ticket, user)),
                            h('div', { class: 'ao-item-meta' },
                                h('span', { class: 'ao-item-time' }, TEXTS.mail.minutes(ticket.timeWorkedMinutes || 0)),
                                ticket.pendingAction ? chip(pendingLabel(ticket.pendingAction), 'warning') : null,
                                h('span', { class: 'ao-toggle' }, expanded ? `${TEXTS.mail.hideDetails} ▴` : `${TEXTS.mail.showDetails} ▾`))),
                        expanded ? h('div', { class: 'ao-ticket-details' },
                            h('div', { class: 'ao-via' }, TEXTS.mail.via[found.via] || ''),
                            ticketSection(ticket, config, user, { isCurrentMail, inDetails: true })) : null));
                });
            } else {
                result.tickets.forEach(found => {
                    content.push(h('div', { class: 'ao-via' }, TEXTS.mail.via[found.via] || ''));
                    content.push(ticketSection({ ...found, ...(ticketByNumber(found.number) || {}) }, config, user, { isCurrentMail }));
                });
            }
            // Una stessa richiesta può richiedere un altro ticket, subito o in seguito.
            content.push(state.extraTicketFor === mail.conversationId
                ? newTicketCard(mail, config, user, { allowIgnore: false, onCancel: () => { state.extraTicketFor = null; render(); } })
                : h('button', { type: 'button', class: 'ao-button ao-subtle ao-add-ticket', onclick: () => { state.extraTicketFor = mail.conversationId; render(); } }, TEXTS.mail.anotherTicket));
        } else if (result.discarded) {
            content.push(discardedCard(mail, result.discarded));
        } else {
            if (result.previous) content.push(previousCard(result.previous));
            if (result.candidates?.length) content.push(candidatesCard(mail, result.candidates, user));
            content.push(newTicketCard(mail, config, user));
        }
        return h('div', { class: 'ao-view' },
            backRow(),
            notice,
            mailHeader(mail, !isCurrentMail),
            content);
    }
    function mailHeader(mail, showOpen) {
        return h('div', { class: 'ao-mail-head' },
            h('div', { class: 'ao-mail-subject' }, mail.subject || '(senza oggetto)'),
            h('div', { class: 'ao-mail-from' },
                h('span', {}, `${mail.sender?.name || ''}${mail.sender?.email ? ` <${mail.sender.email}>` : ''}`),
                h('span', { class: 'ao-item-time' }, formatWhen(mail.receivedAt))),
            showOpen ? h('button', { type: 'button', class: 'ao-link', onclick: () => Outlook.openMessage(mail.itemId) }, TEXTS.mail.openInOutlook) : null);
    }
    function discardedCard(mail, discarded) {
        return h('div', { class: 'ao-card' },
            h('p', { class: 'ao-muted' }, TEXTS.mail.discardedBy(discarded.by || TEXTS.common.none, formatWhen(discarded.at))),
            h('button', {
                type: 'button', class: 'ao-button',
                onclick: async () => {
                    if (await run(() => Services.restoreMail(mail.conversationId))) await refreshLists();
                }
            }, TEXTS.mail.restore));
    }
    function previousCard(ticket) {
        return h('section', { class: 'ao-card ao-card-quiet' },
            h('div', { class: 'ao-card-head' },
                h('h3', { class: 'ao-card-title' }, TEXTS.mail.previousTitle),
                stateChip(ticket)),
            h('div', {},
                h('a', { class: 'ao-ticket-number', href: ticket.url || '#', target: '_blank', rel: 'noopener', title: TEXTS.mail.openInServiceNow }, ticket.number),
                ' ', h('span', { class: 'ao-item-sub' }, ticket.shortDescription)),
            h('div', { class: 'ao-hint' }, [ticket.assignedTo?.name, ticket.closedAt ? TEXTS.mail.closedAt(formatWhen(ticket.closedAt)) : ''].filter(Boolean).join(' · ')),
            h('div', { class: 'ao-hint ao-hint-accent' }, TEXTS.mail.previousHint));
    }
    function candidatesCard(mail, candidates, user) {
        return h('section', { class: 'ao-card ao-card-warning' },
            h('h3', { class: 'ao-card-title' }, TEXTS.mail.candidatesTitle),
            h('p', { class: 'ao-hint' }, TEXTS.mail.candidatesHint),
            h('ul', { class: 'ao-list' }, candidates.map(ticket => h('li', { class: 'ao-candidate' },
                h('div', { class: 'ao-candidate-text' },
                    h('div', { class: 'ao-item-top' }, h('span', { class: 'ao-item-number' }, ticket.number), stateChip(ticket)),
                    h('div', { class: 'ao-item-sub' }, ticket.shortDescription),
                    h('div', { class: 'ao-hint' }, [ticket.assignedTo?.name || TEXTS.common.unassigned, TEXTS.mail.reportedAt(formatWhen(ticket.reportedOn))].join(' · '))),
                h('button', {
                    type: 'button', class: 'ao-button ao-small',
                    onclick: async () => {
                        const result = await run(() => Services.linkTicket(mail.conversationId, ticket.number, user), TEXTS.mail.linked(ticket.number));
                        if (!result) return;
                        state.result = result;
                        await refreshLists();
                    }
                }, TEXTS.mail.link)))));
    }

    // Apertura del ticket: Agenzia suggerita, Attività, anteprima dei campi.
    function ticketDraft(mail, config) {
        if (!state.ticketDrafts[mail.conversationId]) {
            const suggestion = suggestCompany(config, mail.sender, mail.subject, mail.body);
            state.ticketDrafts[mail.conversationId] = { companyId: suggestion?.id || '', activityId: '', subject: buildShortDescription(mail.subject, null), suggestion };
        }
        return state.ticketDrafts[mail.conversationId];
    }
    function newTicketCard(mail, config, user, { allowIgnore = true, onCancel = null } = {}) {
        return h('section', { class: 'ao-card' },
            h('h3', { class: 'ao-card-title' }, TEXTS.mail.newTicket),
            h('p', { class: 'ao-hint' }, TEXTS.mail.newTicketHint),
            ticketForm(mail, config, { user }),
            h('div', { class: 'ao-actions' },
                openTicketButton(mail, config),
                allowIgnore ? ignoreButton(mail, 'ao-button ao-subtle') : null,
                onCancel ? h('button', { type: 'button', class: 'ao-button ao-subtle', onclick: onCancel }, TEXTS.common.cancel) : null));
    }
    // Campi della presa in carico: Preferiti, Agenzia (suggerita), Attività, Short description.
    // Nella versione compatta (In coda) mancano l'anteprima dei campi su ServiceNow.
    function ticketForm(mail, config, { compact = false, user = currentUser(config) } = {}) {
        const draft = ticketDraft(mail, config);
        const companies = Object.entries(config.companies).sort(byName);
        const activities = Object.entries(config.activities).sort(byName);
        const company = config.companies[draft.companyId];
        const activity = config.activities[draft.activityId];
        const favorites = config.favorites.filter(item => config.companies[item.company] && config.activities[item.activity]);
        const rerender = () => render();
        return [
            favorites.length ? h('div', { class: 'ao-favorites' },
                compact ? null : h('span', { class: 'ao-label' }, TEXTS.mail.favorites),
                h('div', { class: 'ao-chips' }, favorites.map(item => h('button', {
                    type: 'button',
                    class: `ao-fav${item.company === draft.companyId && item.activity === draft.activityId ? ' ao-active' : ''}`,
                    onclick: () => { Object.assign(draft, { companyId: item.company, activityId: item.activity }); rerender(); }
                }, `${config.companies[item.company].name} · ${config.activities[item.activity].name}`)))) : null,
            selectField(TEXTS.mail.company, [{ value: '', label: TEXTS.mail.choose }, ...companies.map(([id, item]) => ({ value: id, label: item.name }))],
                draft.companyId, value => { draft.companyId = value; rerender(); }),
            // Riga sempre presente (anche vuota): se l'avviso sparisce, il modulo non cambia altezza.
            h('div', { class: 'ao-hint ao-hint-accent ao-hint-slot' },
                draft.suggestion && draft.suggestion.id === draft.companyId ? TEXTS.mail.suggestedFrom(draft.suggestion.reason) : ''),
            selectField(TEXTS.mail.activity, [{ value: '', label: TEXTS.mail.choose }, ...activities.map(([id, item]) => ({ value: id, label: item.name }))],
                draft.activityId, value => { draft.activityId = value; rerender(); }),
            textField(TEXTS.mail.shortDescription, draft.subject, value => {
                draft.subject = value;
                const preview = root.querySelector('[data-preview="short"]');
                if (preview) preview.textContent = buildShortDescription(value, activity);
            }, { maxlength: '160' }),
            !compact && company && activity ? h('dl', { class: 'ao-fields' },
                h('dt', {}, TEXTS.mail.fieldsPreview), h('dd', {}),
                fieldRow('Company', company.company?.displayValue || company.name),
                fieldRow('Assignment group', company.assignmentGroup?.displayValue),
                fieldRow('Assigned to', user.displayValue || TEXTS.common.none),
                fieldRow('Category', activity.category?.displayValue),
                fieldRow('Impact / Urgency', `${activity.impact || '?'} / ${activity.urgency || '?'}`),
                fieldRow('Time worked', TEXTS.mail.minutes(Number(activity.timeWorked) || 0)),
                fieldRow('Short description', buildShortDescription(draft.subject, activity), 'short'),
                fieldRow('Reported', formatWhen(mail.receivedAt))) : null
        ];
    }
    function openTicketButton(mail, config) {
        const draft = ticketDraft(mail, config);
        const ready = Boolean(config.companies[draft.companyId] && config.activities[draft.activityId]);
        return h('button', { type: 'button', class: 'ao-button ao-primary', disabled: !ready, onclick: event => openTicket(mail, event.currentTarget.closest('.ao-queue-item')) }, TEXTS.mail.openTicket);
    }
    function ignoreButton(mail, className) {
        return confirmButton(TEXTS.mail.discard, `${className} ao-ignore`, async button => {
            // Vista del ticket (in In coda si usa lo swipe: swipeAction).
            const ok = await run(() => Services.discardMail(mail.conversationId, currentUser()));
            if (ok) await refreshLists();
            render(); // se non è riuscito, la card ricompare
        });
    }
    // Ignora in In coda, come lo "swipe to delete" di Gmail / Mail di iPhone (ma con il clic):
    //  1. clic su Ignora → la scheda scorre a sinistra e scopre il tasto rosso con l'icona dell'archivio;
    //  2. clic sul tasto rosso → il tasto si allarga fino a occupare tutta la scheda, poi lo spazio si chiude;
    //  clic altrove (o Esc) → la scheda torna a posto.
    const ARCHIVE_ICON = '<svg viewBox="0 0 20 20" width="18" height="18" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="2.5" y="3.5" width="15" height="4" rx="1"/><path d="M4 7.5v8a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1v-8"/><path d="M8 11h4"/></svg>';
    function swipeAction(mail) {
        const button = h('button', {
            type: 'button', class: 'ao-swipe-action', title: TEXTS.mail.discard, 'aria-label': TEXTS.mail.discard, tabindex: '-1',
            onclick: event => commitSwipe(event.currentTarget.closest('.ao-queue-item'), mail)
        });
        button.innerHTML = `${ARCHIVE_ICON}<span>${TEXTS.common.confirm}</span>`;
        return button;
    }
    function toggleSwipe(card, conversationId) {
        const open = state.swipeOpen !== conversationId;
        closeSwipe();
        if (!open || !card) return;
        state.swipeOpen = conversationId;
        card.classList.add('ao-swiped');
        card.querySelector('.ao-swipe-action')?.setAttribute('tabindex', '0');
        swipeTimer = setTimeout(closeSwipe, SWIPE_TIMEOUT_MS); // senza conferma torna da sola a posto
    }
    const SWIPE_TIMEOUT_MS = 4000;
    let swipeTimer = null;
    function closeSwipe() {
        clearTimeout(swipeTimer);
        if (!state.swipeOpen) return;
        root?.querySelectorAll('.ao-queue-item.ao-swiped').forEach(card => {
            card.classList.remove('ao-swiped');
            card.querySelector('.ao-swipe-action')?.setAttribute('tabindex', '-1');
        });
        state.swipeOpen = null;
    }
    async function commitSwipe(card, mail) {
        clearTimeout(swipeTimer);
        state.swipeOpen = null;
        const reduced = window.matchMedia?.('(prefers-reduced-motion: reduce)').matches;
        state.holdRender = !reduced;
        const done = run(() => Services.discardMail(mail.conversationId, currentUser()));
        if (card && !reduced) {
            card.classList.add('ao-swipe-commit'); // il tasto si allarga su tutta la scheda
            await new Promise(next => setTimeout(next, SWIPE_EXPAND_MS));
            await animateOut(card, 'ao-swipe-collapse', 'ao-slide-collapse', 500); // poi lo spazio si chiude
        }
        state.holdRender = false;
        if (await done) await refreshLists();
        render(); // se non è riuscito, la scheda ricompare
    }
    const SWIPE_EXPAND_MS = 300;
    // Presa in carico (In coda), stesso concetto di Ignora ma dal basso:
    //  1. al clic su "Prendi in carico" nel modulo il blu sale dal basso e copre tutta la scheda in un solo
    //     movimento (il contenuto sale anche lui, più piano, e viene raggiunto), con "Apro il ticket…" (resta così mentre ServiceNow lavora);
    //  2. ticket aperto → "✓ Preso in carico" e lo spazio si chiude;
    //  3. errore → il tasto riscende e la scheda resta com'era.
    const CHECK_ICON = '<svg viewBox="0 0 20 20" width="18" height="18" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4.5 10.5l3.5 3.5 7.5-8"/></svg>';
    const TAKE_BAR_MS = 420;
    const TAKE_HOLD_MS = 220;
    const takeMotion = card => card && !window.matchMedia?.('(prefers-reduced-motion: reduce)').matches;
    function takeBarIn(card) {
        if (!takeMotion(card)) return;
        const bar = h('div', { class: 'ao-take-bar', role: 'status' });
        bar.innerHTML = `<span class="ao-take-spinner" aria-hidden="true"></span><span>${TEXTS.mail.takeBarWorking}</span>`;
        card.style.setProperty('--ao-card-height', `${card.offsetHeight}px`);
        bar.style.height = `${card.clientHeight}px`; // già alto quanto la scheda: sale e la copre tutta
        card.append(bar);
        bar.getBoundingClientRect(); // parte da sotto
        card.classList.add('ao-taking');
        card.dataset.takeStart = String(performance.now());
    }
    async function takeBarDone(card) {
        const bar = card?.querySelector('.ao-take-bar');
        if (!bar) return;
        bar.innerHTML = `${CHECK_ICON}<span>${TEXTS.mail.takeBarDone}</span>`;
        // La chiusura parte appena il blu ha finito di salire (se ServiceNow ha già risposto).
        const risen = performance.now() - Number(card.dataset.takeStart || 0);
        if (risen < TAKE_BAR_MS) await new Promise(next => setTimeout(next, TAKE_BAR_MS - risen));
        await new Promise(next => setTimeout(next, TAKE_HOLD_MS)); // un attimo per leggere la spunta
        await animateOut(card, 'ao-swipe-collapse', 'ao-slide-collapse', 500);
    }
    async function takeBarOut(card) {
        const bar = card?.querySelector('.ao-take-bar');
        if (!bar) return;
        card.classList.remove('ao-taking');
        await new Promise(next => setTimeout(next, TAKE_BAR_MS));
        bar.remove();
    }
    function animateOut(card, className, lastAnimation, fallbackMs) {
        if (!card || window.matchMedia?.('(prefers-reduced-motion: reduce)').matches) {
            state.holdRender = false;
            return Promise.resolve();
        }
        // Durante l'animazione il pannello non si ridisegna (altrimenti la card sparirebbe di colpo).
        state.holdRender = true;
        return new Promise(done => {
            const resolve = () => {
                if (!state.holdRender) return;
                state.holdRender = false;
                done();
            };
            card.style.setProperty('--ao-card-height', `${card.offsetHeight}px`);
            card.classList.add(className);
            card.addEventListener('animationend', event => { if (event.animationName === lastAnimation) resolve(); });
            setTimeout(resolve, fallbackMs); // riserva
        });
    }
    function fieldRow(label, value, previewKey) {
        return [h('dt', {}, label), h('dd', previewKey ? { dataset: { preview: previewKey } } : {}, value || TEXTS.common.none)];
    }
    function buildShortDescription(subject, activity) {
        const clean = String(subject || '').replace(/^\s*((re|r|fw|fwd|i)\s*:\s*)+/i, '').trim();
        const prefix = activity?.shortDescriptionPrefix || '';
        return prefix ? `${prefix} ${clean.replace(/^\s*\[[^\]]+\]\s*/, '')}`.trim() : clean;
    }
    // Apre l'Incident su ServiceNow (assegnato a chi lo apre); il servizio registra poi il collegamento.
    // card: la card di In coda da cui si prende in carico (si chiude con un'animazione se va tutto bene).
    async function openTicket(mail, card = null) {
        const config = loadConfig();
        const draft = ticketDraft(mail, config);
        const company = config.companies[draft.companyId];
        const activity = config.activities[draft.activityId];
        if (!company || !activity) return showStatus(TEXTS.mail.selectBoth, 'danger');
        if (!config.profile.assignedTo.displayValue) return showStatus(TEXTS.mail.profileMissing, 'danger');
        // Corpo completo se è la mail aperta in Outlook; altrimenti l'anteprima della Coda.
        const body = state.mail?.conversationId === mail.conversationId ? state.mail.body : mail.body;
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
            reportedOn: mail.receivedAt ? formatServiceNowDate(parseDate(mail.receivedAt)) : ''
        };
        if (card) state.holdRender = true; // la card resta ferma finché non parte l'animazione
        takeBarIn(card);
        const ticket = await run(() => Services.createTicket(mail, {
            companyId: draft.companyId, companyName: company.name,
            activityId: draft.activityId, activityName: activity.name, fields
        }), result => TEXTS.mail.ticketOpened(result.number));
        if (!ticket) {
            await takeBarOut(card);
            state.holdRender = false;
            render();
            return;
        }
        if (takeMotion(card)) await takeBarDone(card);
        state.holdRender = false;
        delete state.ticketDrafts[mail.conversationId];
        state.ticketCache[ticket.number] = ticket;
        if (state.queueExpanded === mail.conversationId) state.queueExpanded = null;
        if (state.extraTicketFor === mail.conversationId) state.extraTicketFor = null;
        config.recent = [{ company: draft.companyId, activity: draft.activityId, timestamp: new Date().toISOString() },
            ...config.recent.filter(item => !(item.company === draft.companyId && item.activity === draft.activityId))].slice(0, 10);
        saveConfig(config);
        render();
        await refreshLists();
    }

    // Ticket collegato: assegnatario, tempo, azione all'invio, storico.
    function ticketSection(ticket, config, user, { isCurrentMail, showOpen = false, inDetails = false }) {
        const open = OPEN_STATES.includes(ticket.state);
        return [
            ticketCard(ticket, showOpen, inDetails, user),
            open && ticket.messageId && !ticket.mailUnconfirmed ? onSendCard(ticket, config, isCurrentMail) : null,
            open && (!ticket.messageId || ticket.mailUnconfirmed) ? h('p', { class: 'ao-hint' }, TEXTS.mail.noLinkedMail) : null,
            historyCard(ticket)
        ];
    }
    // Assegnatario come etichetta (me / nome del collega) e, se è di un collega, "Prendi in carico".
    function assigneeTag(ticket, user) {
        const mine = isMine(ticket, user);
        return chip(mine ? TEXTS.common.me : (ticket.assignedTo?.name || TEXTS.common.unassigned), mine ? 'success' : 'person', TEXTS.mail.assignedTo(ticket.assignedTo?.name || TEXTS.common.unassigned));
    }
    // In alto a destra, uno sotto l'altro: stato, poi assegnatario.
    function ticketTags(ticket, user) {
        return h('div', { class: 'ao-tags' }, stateChip(ticket), assigneeTag(ticket, user));
    }
    function takeOverButton(ticket, user) {
        if (isMine(ticket, user) || !OPEN_STATES.includes(ticket.state)) return null;
        return confirmButton(TEXTS.mail.takeOver, 'ao-button ao-small ao-take', async () => {
            const config = loadConfig();
            if (!config.profile.assignedTo.displayValue) return showStatus(TEXTS.mail.profileMissing, 'danger');
            const updated = await run(() => Services.assignTicket(ticket.number, config.profile.assignedTo), TEXTS.mail.taken(ticket.number));
            if (!updated) return;
            rememberTicket(updated);
            render();
            refreshLists();
        });
    }
    // inDetails: dentro i dettagli di un ticket compatto, senza ripetere numero, stato e descrizione.
    function ticketCard(ticket, showOpen, inDetails = false, user = currentUser()) {
        const open = OPEN_STATES.includes(ticket.state);
        const takeOver = takeOverButton(ticket, user);
        return h('section', { class: 'ao-card' },
            inDetails ? null : h('div', { class: 'ao-ticket-head' },
                h('div', { class: 'ao-item-left' },
                    h('a', { class: 'ao-ticket-number', href: ticket.url || '#', target: '_blank', rel: 'noopener', title: TEXTS.mail.openInServiceNow }, ticket.number),
                    h('div', { class: 'ao-item-title' }, ticket.shortDescription)),
                ticketTags(ticket, user)),
            h('div', { class: 'ao-item-sub' }, [ticket.companyName, ticket.activityName].filter(Boolean).join(' · ')),
            // Assegnatario: etichetta (nei dettagli compatti c'è già nell'intestazione) ed eventuale presa in carico.
            takeOver ? h('div', { class: 'ao-item-meta' }, takeOver) : null,
            ticket.state === '3' && ticket.holdExpiration
                ? h('div', { class: `ao-hint${isHoldExpired(ticket) ? ' ao-hint-danger' : ''}` },
                    `${HOLD_REASONS.find(item => item.value === ticket.holdReason)?.label || 'On Hold'} · ${TEXTS.tickets.holdUntil(formatWhen(ticket.holdExpiration))}`) : null,
            h('div', { class: 'ao-hint' }, [TEXTS.mail.openedAt(formatWhen(ticket.openedAt)), TEXTS.mail.reportedAt(formatWhen(ticket.reportedOn))].join(' · ')),
            showOpen && ticket.messageId ? h('button', { type: 'button', class: 'ao-link', onclick: () => Outlook.openMessage(ticket.messageId) }, TEXTS.mail.openInOutlook) : null,
            h('div', { class: 'ao-time' },
                h('span', { class: 'ao-label' }, TEXTS.mail.timeWorked),
                h('div', { class: 'ao-time-row' },
                    TIME_STEPS.filter(step => step < 0).map(step => timeButton(ticket, step, open)),
                    h('span', { class: 'ao-time-value' }, TEXTS.mail.minutes(ticket.timeWorkedMinutes || 0)),
                    TIME_STEPS.filter(step => step > 0).map(step => timeButton(ticket, step, open)))),
            open ? null : h('p', { class: 'ao-muted' }, TEXTS.mail.closed(TICKET_STATES[ticket.state]?.label || ticket.state)));
    }
    function timeButton(ticket, step, enabled) {
        return h('button', {
            type: 'button', class: 'ao-button ao-small ao-time-step',
            disabled: !enabled || (step < 0 && (ticket.timeWorkedMinutes || 0) <= 0),
            onclick: async () => {
                const before = ticket.timeWorkedMinutes || 0;
                const updated = await run(() => Services.addTime(ticket.number, step, currentUser().displayValue));
                if (!updated) return;
                rememberTicket(updated);
                render();
                showStatus(TEXTS.mail.timeChanged(updated.timeWorkedMinutes - before, updated.timeWorkedMinutes));
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
    function actionDraft(ticket, config) {
        if (!state.actionDrafts[ticket.number]) {
            const saved = ticket.pendingAction || {};
            const key = combinationKey(ticket.companyId, ticket.activityId);
            state.actionDrafts[ticket.number] = {
                type: saved.type || 'comment',
                resolutionCode: saved.resolutionCode || RESOLUTION_CODES[0],
                holdReason: saved.holdReason || HOLD_REASONS[0].value,
                holdExpiration: saved.holdExpiration ? toDateTimeInput(parseDate(saved.holdExpiration)) : defaultHoldExpiration(),
                workNotes: saved.workNotes || '',
                template: config.combinationResponses[key]?.response ? 'combination' : '',
                includeReference: true
            };
        }
        return state.actionDrafts[ticket.number];
    }
    function onSendCard(ticket, config, isCurrentMail) {
        const draft = actionDraft(ticket, config);
        const rerender = () => render();
        const key = combinationKey(ticket.companyId, ticket.activityId);
        const templates = [
            { value: '', label: TEXTS.mail.replyEmpty },
            ...(config.combinationResponses[key]?.response ? [{ value: 'combination', label: TEXTS.mail.replyCombination }] : []),
            ...Object.entries(config.usefulResponses).sort(byName).map(([id, item]) => ({ value: `useful:${id}`, label: item.title }))
        ];
        return h('section', { class: 'ao-card' },
            h('h3', { class: 'ao-card-title' }, TEXTS.mail.onSend),
            h('p', { class: 'ao-hint' }, TEXTS.mail.onSendHint),
            ticket.pendingAction ? h('div', { class: 'ao-pending' },
                h('span', {}, TEXTS.mail.pending(pendingLabel(ticket.pendingAction))),
                h('button', { type: 'button', class: 'ao-link', onclick: () => savePendingAction(ticket, null) }, TEXTS.mail.clearPending)) : null,
            filterBar(Object.entries(TEXTS.mail.actions).map(([value, label]) => ({ value, label })), draft.type, value => { draft.type = value; rerender(); }),
            draft.type === 'resolve' ? selectField(TEXTS.mail.resolutionCode, RESOLUTION_CODES.map(code => ({ value: code, label: code })), draft.resolutionCode, value => { draft.resolutionCode = value; }) : null,
            draft.type === 'hold' ? [
                selectField(TEXTS.mail.holdReason, HOLD_REASONS, draft.holdReason, value => { draft.holdReason = value; }),
                textField(TEXTS.mail.holdExpiration, draft.holdExpiration, value => { draft.holdExpiration = value; }, { type: 'datetime-local' })
            ] : null,
            draft.type !== 'comment' ? textField(TEXTS.mail.workNotes, draft.workNotes, value => { draft.workNotes = value; }, { multiline: true, rows: '2' }) : null,
            selectField(TEXTS.mail.replyText, templates, draft.template, value => { draft.template = value; }),
            checkboxField(TEXTS.mail.includeReference(ticket.number), draft.includeReference, value => { draft.includeReference = value; }),
            h('div', { class: 'ao-actions' },
                h('button', { type: 'button', class: 'ao-button ao-primary', onclick: () => prepareReply(ticket, isCurrentMail) }, TEXTS.mail.prepareReply)));
    }
    function buildPendingAction(draft, by) {
        if (draft.type === 'comment') return null;
        const action = { type: draft.type, workNotes: draft.workNotes.trim(), by };
        if (draft.type === 'resolve') action.resolutionCode = draft.resolutionCode;
        if (draft.type === 'hold') {
            const date = parseDate(draft.holdExpiration.replace('T', ' '));
            action.holdReason = draft.holdReason;
            action.holdExpiration = date ? formatServiceNowDate(date) : '';
        }
        return action;
    }
    async function savePendingAction(ticket, action) {
        const updated = await run(() => Services.setPendingAction(ticket.number, action));
        if (!updated) return null;
        rememberTicket(updated);
        delete state.actionDrafts[ticket.number];
        render();
        return updated;
    }
    async function prepareReply(ticket, isCurrentMail) {
        const config = loadConfig();
        const draft = actionDraft(ticket, config);
        if (draft.type === 'hold' && !parseDate(draft.holdExpiration)) return showStatus(TEXTS.mail.holdExpiration, 'danger');
        const template = draft.template === 'combination'
            ? config.combinationResponses[combinationKey(ticket.companyId, ticket.activityId)]?.response
            : config.usefulResponses[draft.template.replace(/^useful:/, '')]?.response;
        // Il riferimento all'Incident nella risposta rende sicuro il collegamento mail ↔ ticket.
        const text = [draft.includeReference ? TEXTS.mail.reference(ticket.number) : '', template || ''].filter(Boolean).join('\n\n');
        const updated = await savePendingAction(ticket, buildPendingAction(draft, currentUser(config).displayValue));
        if (!updated) return;
        // In Outlook la risposta si apre solo sulla mail aperta: per un'altra mail
        // si apre la mail e l'azione resta salvata. L'anteprima le gestisce entrambe.
        let opened = false;
        if (isCurrentMail) opened = Outlook.openReply(text);
        else {
            Outlook.openMessage(ticket.messageId);
            if (window.AmpPreviewOutlook) opened = Outlook.openReply(text);
        }
        showStatus(opened ? TEXTS.mail.replyOpened : TEXTS.mail.pendingSaved, opened ? 'success' : 'warning');
    }

    // Storico di ServiceNow, caricato solo quando serve e a pagine.
    function historyCard(ticket) {
        const history = state.history[ticket.number];
        return h('section', { class: 'ao-card' },
            h('div', { class: 'ao-card-head' },
                h('h3', { class: 'ao-card-title' }, TEXTS.mail.history),
                history?.open ? null : h('button', { type: 'button', class: 'ao-button ao-small', onclick: () => loadHistory(ticket.number) }, TEXTS.mail.loadHistory)),
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
                    ? h('button', { type: 'button', class: 'ao-link', onclick: () => loadHistory(ticket.number, false, true) }, TEXTS.mail.loadMore(history.total - history.entries.length)) : null
            ] : null);
    }
    async function loadHistory(number, reload = false, more = false) {
        const current = state.history[number] || { entries: [], total: 0 };
        const limit = reload ? Math.max(HISTORY_PAGE, current.entries.length) : HISTORY_PAGE;
        const offset = more ? current.entries.length : 0;
        state.history[number] = { ...current, open: true, loading: true };
        if (!reload) render();
        try {
            const page = await Services.getTicketActivity(number, { offset, limit });
            state.history[number] = {
                open: true, loading: false, total: page.total,
                entries: more ? [...current.entries, ...page.entries] : page.entries
            };
        } catch (error) {
            state.history[number] = { ...current, open: false, loading: false };
            showStatus(error?.message || String(error), 'danger');
        }
        render();
    }

    // ---------------------------- Impostazioni ----------------------------
    // Profilo e configurazioni in schede separate (barra secondaria). Ogni configurazione mostra
    // subito tutti i suoi valori; la modifica avviene in linea. EDITORS descrive i campi.
    // I riferimenti a ServiceNow (Company, Assignment group, Category, Assigned to) si scrivono
    // per nome: il sys_id resta interno (arriva dall'import o dalla verifica del servizio) e non
    // si mostra né si modifica. Se il nome cambia, il sys_id si azzera e la voce è "da verificare".
    const EDITORS = {
        companies: {
            properties: item => [
                [TEXTS.settings.company, referenceValue(item.company)],
                [TEXTS.settings.assignmentGroup, referenceValue(item.assignmentGroup)],
                [TEXTS.settings.domains, (item.domains || []).length
                    ? (item.domains || []).map(domain => chip(`@${domain}`, 'neutral'))
                    : h('span', { class: 'ao-missing' }, TEXTS.settings.noDomains)]
            ],
            fields: [
                { key: 'name', label: TEXTS.settings.companyName, required: true },
                { key: 'company', label: TEXTS.settings.company, reference: true },
                { key: 'assignmentGroup', label: TEXTS.settings.assignmentGroup, reference: true },
                { key: 'domains', label: TEXTS.settings.domains, hint: TEXTS.settings.domainsHint, list: true }
            ]
        },
        activities: {
            properties: item => [
                [TEXTS.settings.category, referenceValue(item.category), { twoLines: true }],
                [TEXTS.settings.impactUrgency, `${SEVERITY_LABELS[item.impact] || '?'} / ${SEVERITY_LABELS[item.urgency] || '?'}`],
                [TEXTS.settings.timeWorked, TEXTS.mail.minutes(Number(item.timeWorked) || 0)],
                [TEXTS.settings.prefix, item.shortDescriptionPrefix || h('span', { class: 'ao-missing' }, TEXTS.common.none)]
            ],
            fields: [
                { key: 'name', label: TEXTS.settings.companyName, required: true },
                { key: 'category', label: TEXTS.settings.category, reference: true },
                { key: 'impact', label: TEXTS.settings.impact, options: ['1', '2', '3'] },
                { key: 'urgency', label: TEXTS.settings.urgency, options: ['1', '2', '3'] },
                { key: 'timeWorked', label: TEXTS.settings.timeWorkedMinutes, number: true },
                { key: 'shortDescriptionPrefix', label: TEXTS.settings.prefixLong }
            ]
        },
        responses: {
            collection: 'usefulResponses',
            nameKey: 'title',
            body: item => h('p', { class: 'ao-response-text' }, item.response || ''),
            fields: [
                { key: 'title', label: TEXTS.settings.responseTitle, required: true },
                { key: 'response', label: TEXTS.settings.response, multiline: true }
            ]
        }
    };
    const collectionOf = tab => EDITORS[tab].collection || tab;

    // Valore di un riferimento: il nome, con l'avviso se non è ancora collegato a ServiceNow.
    function referenceValue(reference) {
        const normalized = normalizeReference(reference);
        if (!normalized.displayValue) return h('span', { class: 'ao-missing' }, TEXTS.common.none);
        return [normalized.displayValue, normalized.value ? null : chip(TEXTS.settings.toVerify, 'warning', TEXTS.settings.toVerifyHint)];
    }
    // Il sys_id resta quello di prima solo se il nome non è cambiato.
    function keepReference(previous, displayValue) {
        const before = normalizeReference(previous);
        const name = String(displayValue || '').trim();
        return { displayValue: name, value: name && name === before.displayValue ? before.value : '' };
    }

    function renderSettings(config) {
        const tabs = Object.keys(TEXTS.settings.tabs);
        const tab = tabs.includes(state.settingsTab) ? state.settingsTab : 'profile';
        // Niente riga ←: dalle impostazioni si esce con le schede o con l'ingranaggio.
        return h('div', { class: 'ao-view' },
            filterBar(tabs.map(value => ({
                value,
                label: TEXTS.settings.tabs[value],
                count: value === 'profile' ? undefined : Object.keys(config[collectionOf(value)]).length
            })), tab, value => { state.settingsTab = value; state.editing = null; render(); }),
            tab === 'profile' ? profileSection(config) : collectionSection(tab, config));
    }
    function profileSection(config) {
        const draft = { name: config.profile.assignedTo.displayValue };
        const outlookName = state.mail?.userName || window.TestPluginAPI?.getCurrentMessageContext?.()?.mailbox?.userDisplayName || '';
        return [
            h('section', { class: 'ao-card' },
                h('h3', { class: 'ao-card-title' }, TEXTS.settings.profile),
                textField(TEXTS.settings.assignedTo, draft.name, value => { draft.name = value; }, { placeholder: outlookName, hint: TEXTS.settings.profileHint }),
                config.profile.assignedTo.displayValue && !config.profile.assignedTo.value
                    ? h('div', {}, chip(TEXTS.settings.toVerify, 'warning'), ' ', h('span', { class: 'ao-hint' }, TEXTS.settings.toVerifyHint)) : null,
                h('div', { class: 'ao-actions' },
                    h('button', {
                        type: 'button', class: 'ao-button ao-primary',
                        onclick: () => {
                            const current = loadConfig();
                            current.profile = { ...current.profile, assignedTo: keepReference(current.profile.assignedTo, draft.name) };
                            saveConfig(current);
                            render();
                            showStatus(TEXTS.settings.saved);
                        }
                    }, TEXTS.common.save))),
            h('section', { class: 'ao-card' },
                h('h3', { class: 'ao-card-title' }, TEXTS.settings.theme),
                filterBar(Object.entries(TEXTS.settings.themes).map(([value, label]) => ({ value, label })), config.preferences?.theme || 'auto', value => {
                    const current = loadConfig();
                    current.preferences = { ...(current.preferences || {}), theme: value };
                    saveConfig(current);
                    applyTheme(current);
                    render();
                }),
                h('p', { class: 'ao-hint' }, TEXTS.settings.themeHint)),
            h('section', { class: 'ao-card' },
                h('h3', { class: 'ao-card-title' }, TEXTS.settings.backup),
                h('p', { class: 'ao-hint' }, TEXTS.settings.backupHint),
                h('div', { class: 'ao-actions' },
                    h('button', { type: 'button', class: 'ao-button', onclick: importConfig }, TEXTS.settings.import),
                    h('button', { type: 'button', class: 'ao-button', onclick: exportConfig }, TEXTS.settings.export))),
            h('p', { class: 'ao-footnote' }, TEXTS.settings.version(BUILD))
        ];
    }
    function collectionSection(tab, config) {
        const editor = EDITORS[tab];
        const items = Object.entries(config[collectionOf(tab)]).sort(byName);
        const editing = state.editing?.kind === tab ? state.editing : null;
        return [
            editing?.id === null ? editorForm(tab, null, {}) : h('div', { class: 'ao-toolbar' },
                h('span', { class: 'ao-hint' }, `${items.length} ${TEXTS.settings.tabs[tab].toLowerCase()}`),
                h('button', { type: 'button', class: 'ao-button ao-small', onclick: () => { state.editing = { kind: tab, id: null }; render(); } }, `+ ${TEXTS.common.add}`)),
            items.length ? h('ul', { class: 'ao-config-list' }, items.map(([id, item]) => h('li', {}, editing?.id === id
                ? editorForm(tab, id, item)
                : h('div', { class: 'ao-config-card' },
                    h('div', { class: 'ao-config-card-head' },
                        h('span', { class: 'ao-config-name' }, item[editor.nameKey || 'name']),
                        h('button', { type: 'button', class: 'ao-icon-button', title: TEXTS.common.edit, 'aria-label': TEXTS.common.edit, onclick: () => { state.editing = { kind: tab, id }; render(); } }, '✎'),
                        confirmButton('🗑', 'ao-icon-button ao-danger', () => deleteItem(tab, id))),
                    editor.properties
                        ? h('dl', { class: 'ao-props' }, editor.properties(item).map(([label, value, options = {}]) => {
                            const className = options.twoLines ? 'ao-prop-two-lines' : undefined;
                            return [h('dt', { class: className }, label), h('dd', { class: className }, value)];
                        }))
                        : editor.body(item)))))
                : (editing ? null : h('p', { class: 'ao-muted' }, TEXTS.settings.empty))
        ];
    }
    function editorForm(tab, id, item) {
        const editor = EDITORS[tab];
        const values = {};
        const fields = editor.fields.map(field => {
            if (field.reference) {
                values[field.key] = normalizeReference(item[field.key]).displayValue;
                return textField(field.label, values[field.key], value => { values[field.key] = value; }, { hint: TEXTS.settings.referenceHint });
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
            const collection = collectionOf(tab);
            const targetId = id || createId(values[nameKey]);
            const previous = config[collection][targetId] || {};
            const next = { ...previous };
            editor.fields.forEach(field => {
                const value = values[field.key];
                if (field.reference) next[field.key] = keepReference(previous[field.key], value);
                else if (field.list) next[field.key] = normalizeDomains(value);
                else if (field.number) next[field.key] = Math.max(0, Number.parseInt(value, 10) || 0);
                else next[field.key] = field.multiline ? String(value) : String(value).trim();
            });
            if (tab === 'responses') {
                const now = new Date().toISOString();
                next.created ||= now;
                next.updated = now;
            }
            config[collection][targetId] = next;
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
    function deleteItem(tab, id) {
        const config = loadConfig();
        delete config[collectionOf(tab)][id];
        if (tab === 'companies' || tab === 'activities') {
            const key = tab === 'companies' ? 'company' : 'activity';
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
#ao-root .ao-tabs { flex: none; display: flex; align-items: center; gap: 2px; padding: 0 4px; background: var(--ao-surface); border-bottom: 1px solid var(--ao-line); }
#ao-root .ao-tabs .ao-icon-button { flex: none; font-size: var(--ao-fs-lg); }
#ao-root .ao-settings-button.ao-active { color: var(--ao-text-accent); background: var(--ao-accent-bg); }
#ao-root .ao-tab { flex: 1; display: flex; align-items: center; justify-content: center; gap: 5px; padding: 9px 4px 7px; border: 0; border-bottom: 2px solid transparent; background: none; color: var(--ao-text-muted); font-size: var(--ao-fs-md); font-weight: var(--ao-fw-normal); cursor: pointer; }
#ao-root .ao-tab:hover { color: var(--ao-text); background: var(--ao-bg-soft); }
#ao-root .ao-tab.ao-active { color: var(--ao-text-accent); border-bottom-color: var(--ao-accent); font-weight: var(--ao-fw-bold); }
/* Scheda aperta: leggero zoom della scritta (e del numero), che resta finché è selezionata */
#ao-root .ao-tab > * { transition: transform .35s ease-in-out; }
#ao-root .ao-tab.ao-active > * { transform: scale(1.04); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-tab > * { transition: none; } }
/* Schede: il conteggio è un numero pulito; con una notifica lo stesso numero ha lo sfondo blu. */
#ao-root .ao-counter { display: inline-block; min-width: 18px; height: 18px; padding: 0 4px; border-radius: var(--ao-radius-badge); color: var(--ao-text-muted); font-size: var(--ao-fs-md); font-weight: var(--ao-fw-medium); line-height: 18px; text-align: center; }
#ao-root .ao-tab.ao-active .ao-counter { color: var(--ao-text-accent); }
#ao-root .ao-counter-notify, #ao-root .ao-tab.ao-active .ao-counter-notify { background: var(--ao-accent); color: var(--ao-white); font-weight: var(--ao-fw-bold); }
/* Card: numero dei messaggi nuovi su sfondo blu. */
#ao-root .ao-news { display: inline-block; min-width: 18px; height: 18px; padding: 0 4px; border-radius: var(--ao-radius-badge); background: var(--ao-accent); color: var(--ao-white); font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-bold); line-height: 18px; text-align: center; }
#ao-root .ao-new-head .ao-news { align-self: center; }
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
#ao-root .ao-scroll { flex: 1; min-height: 0; display: grid; align-content: start; gap: 10px; padding: 0 12px 12px; scroll-padding-top: 60px; overflow-y: auto; scrollbar-gutter: stable; scrollbar-width: thin; scrollbar-color: var(--ao-line-strong) transparent; }
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
/* Prendi in carico (righe di In coda): azzurro leggero al passaggio del mouse */
#ao-root .ao-take:hover { color: var(--ao-text-accent); background: var(--ao-accent-bg); border-color: var(--ao-accent-border); }
/* Ignora: solo la scritta diventa rossa (niente sfondo né bordo), anche quando chiede conferma */
#ao-root .ao-ignore:hover { color: var(--ao-danger); background: transparent; border-color: transparent; }
#ao-root .ao-ignore.ao-armed, #ao-root .ao-ignore.ao-armed:hover { background: transparent !important; border-color: transparent !important; }
#ao-root .ao-armed, #ao-root .ao-armed:hover { background: var(--ao-danger-bg) !important; border-color: var(--ao-danger) !important; color: var(--ao-danger) !important; }
#ao-root .ao-icon-button { display: inline-flex; align-items: center; justify-content: center; min-width: 28px; height: 28px; padding: 0 6px; border: 1px solid transparent; border-radius: var(--ao-radius-base); background: none; color: var(--ao-text-muted); cursor: pointer; }
#ao-root .ao-icon-button:hover { background: var(--ao-bg-soft); border-color: var(--ao-line); }
#ao-root .ao-danger:hover { color: var(--ao-danger); }
#ao-root .ao-link { padding: 0; border: 0; background: none; color: var(--ao-text-accent); font-size: var(--ao-fs-md); cursor: pointer; text-align: left; }
#ao-root .ao-link:hover { text-decoration: underline; }
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
#ao-root .ao-card-title .ao-count { margin-left: 6px; }

/* Elenchi di Coda e Ticket */
#ao-root .ao-list { display: grid; gap: 6px; margin: 0; padding: 0; list-style: none; }
#ao-root .ao-item { display: grid; gap: 3px; padding: 9px 10px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); box-shadow: var(--ao-shadow-card); cursor: pointer; }
/* Passaggio del mouse, uguale in tutte le sezioni: solo il titolo diventa blu (sfondo e bordo invariati).
   Selezionata (modulo o dettagli aperti): bordo blu normale, sfondo bianco. */
#ao-root .ao-item:hover .ao-item-title, #ao-root .ao-ticket-compact-head:hover .ao-item-title { color: var(--ao-text-accent); }
#ao-root .ao-expanded, #ao-root .ao-expanded:hover, #ao-root .ao-ticket-compact.ao-expanded { background: var(--ao-surface); border-color: var(--ao-accent); }
#ao-root .ao-item-top { display: flex; align-items: baseline; justify-content: space-between; gap: 8px; }
#ao-root .ao-item-title { font-weight: var(--ao-fw-medium); overflow-wrap: anywhere; }
#ao-root .ao-item-number { font-weight: var(--ao-fw-bold); color: var(--ao-text-accent); }
#ao-root .ao-item-sub { color: var(--ao-text-muted); font-size: var(--ao-fs-md); overflow-wrap: anywhere; }
#ao-root .ao-item-time { color: var(--ao-text-subtle); font-size: var(--ao-fs-sm); white-space: nowrap; }
#ao-root .ao-item-meta { display: flex; flex-wrap: wrap; align-items: center; gap: 6px; margin-top: 3px; }
#ao-root .ao-empty, #ao-root .ao-footnote { margin: 8px 0; color: var(--ao-text-subtle); text-align: center; }
#ao-root .ao-footnote { font-size: var(--ao-fs-sm); }

/* Etichette colorate */
#ao-root .ao-chip { display: inline-flex; align-items: center; padding: 1px 8px; border-radius: var(--ao-radius-pill); font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-medium); white-space: nowrap; }
#ao-root .ao-chip.ao-tone-neutral { background: var(--ao-line-light); color: var(--ao-text-muted); }
#ao-root .ao-chip.ao-tone-person { background: var(--ao-surface); color: var(--ao-text); box-shadow: inset 0 0 0 1px var(--ao-text-subtle); }
#ao-root .ao-tags { display: flex; flex-direction: column; align-items: flex-end; gap: 4px; flex: none; }
#ao-root .ao-chip.ao-tone-accent { background: var(--ao-accent-bg); color: var(--ao-text-accent); }
#ao-root .ao-chip.ao-tone-success { background: var(--ao-success-bg); color: var(--ao-success); }
#ao-root .ao-chip.ao-tone-warning { background: var(--ao-warning-bg); color: var(--ao-warning); }
#ao-root .ao-chip.ao-tone-danger { background: var(--ao-danger-bg); color: var(--ao-danger); }
#ao-root .ao-chip.ao-tone-hold { background: var(--ao-hold-bg); color: var(--ao-hold); }

/* Riquadri */
#ao-root .ao-card { display: grid; gap: 8px; padding: 12px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); box-shadow: var(--ao-shadow-card); }
#ao-root .ao-card-head { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
#ao-root .ao-card-title { margin: 0; font-size: var(--ao-fs-lg); font-weight: var(--ao-fw-medium); }
#ao-root .ao-notice { padding: 8px 10px; border-radius: var(--ao-radius-base); background: var(--ao-accent-bg); color: var(--ao-text-muted); font-size: var(--ao-fs-md); }
#ao-root .ao-notice.ao-tone-danger { background: var(--ao-danger-bg); color: var(--ao-danger); }
#ao-root .ao-muted { margin: 0; color: var(--ao-text-muted); }
#ao-root .ao-hint { color: var(--ao-text-subtle); font-size: var(--ao-fs-sm); }
#ao-root .ao-hint-accent { margin-top: -4px; color: var(--ao-text-accent); }
#ao-root .ao-hint-slot { min-height: 1.4em; }
#ao-root .ao-hint-danger { color: var(--ao-danger); font-weight: var(--ao-fw-medium); }
#ao-root .ao-actions { display: flex; flex-wrap: wrap; gap: 8px; }
#ao-root .ao-actions .ao-primary { flex: 1; }

/* Mail aperta */
#ao-root .ao-mail-head { display: grid; gap: 2px; padding: 2px 2px 0; }
#ao-root .ao-mail-subject { font-size: var(--ao-fs-lg); font-weight: var(--ao-fw-bold); overflow-wrap: anywhere; }
#ao-root .ao-mail-from { display: flex; justify-content: space-between; gap: 8px; color: var(--ao-text-muted); font-size: var(--ao-fs-md); overflow-wrap: anywhere; }

/* Campi */
#ao-root .ao-field { display: grid; gap: 3px; min-width: 0; }
#ao-root .ao-label { color: var(--ao-text-muted); font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-medium); }
#ao-root .ao-input { width: 100%; min-height: var(--ao-control-height); padding: 5px 8px; border: 1px solid var(--ao-line-strong); border-radius: var(--ao-radius-base); background: var(--ao-surface); color: var(--ao-text); font: inherit; }
#ao-root textarea.ao-input { resize: vertical; line-height: 1.4; }
#ao-root .ao-input:focus { border-color: var(--ao-accent); box-shadow: var(--ao-shadow-focus); outline: none; }
#ao-root .ao-grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
#ao-root .ao-favorites { display: grid; gap: 4px; }
#ao-root .ao-chips { display: flex; flex-wrap: wrap; gap: 5px; }
#ao-root .ao-fav { padding: 3px 9px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-pill); background: var(--ao-surface); font-size: var(--ao-fs-sm); cursor: pointer; }
#ao-root .ao-fav:hover { border-color: var(--ao-accent-border); background: var(--ao-accent-bg); }
#ao-root .ao-fav.ao-active { border-color: var(--ao-accent); background: var(--ao-accent-bg); color: var(--ao-text-accent); font-weight: var(--ao-fw-medium); }
#ao-root .ao-fields { display: grid; grid-template-columns: auto 1fr; gap: 3px 10px; margin: 0; padding: 8px 10px; border-radius: var(--ao-radius-base); background: var(--ao-bg-soft); font-size: var(--ao-fs-md); }
#ao-root .ao-fields dt { color: var(--ao-text-muted); }
#ao-root .ao-fields dt:first-child { font-weight: var(--ao-fw-medium); color: var(--ao-text); }
#ao-root .ao-fields dd { margin: 0; overflow-wrap: anywhere; }

/* Ticket */
#ao-root .ao-ticket-head { display: flex; align-items: flex-start; justify-content: space-between; gap: 8px; }
#ao-root .ao-ticket-number { font-size: var(--ao-fs-lg); font-weight: var(--ao-fw-bold); color: var(--ao-text-accent); text-decoration: none; }
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

/* Torna all'elenco (vista di lavoro del ticket) */
#ao-root .ao-back-row { display: flex; align-items: center; min-height: 22px; }
#ao-root .ao-back-link { font-weight: var(--ao-fw-medium); }

/* Impostazioni */
#ao-root .ao-config-list { display: grid; gap: 4px; margin: 0; padding: 0; list-style: none; }
#ao-root .ao-config-item { display: flex; align-items: center; gap: 4px; padding: 6px 4px 6px 10px; border: 1px solid var(--ao-line-light); border-radius: var(--ao-radius-base); }
#ao-root .ao-config-item:hover { background: var(--ao-bg-soft); }
#ao-root .ao-config-text { flex: 1; display: grid; min-width: 0; }
#ao-root .ao-config-name { font-weight: var(--ao-fw-medium); }
#ao-root .ao-config-text .ao-hint { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
#ao-root .ao-editor { display: grid; gap: 8px; padding: 10px; border: 1px solid var(--ao-accent-border); border-radius: var(--ao-radius-base); background: var(--ao-accent-bg); }

/* In coda: riga con azioni e modulo compatto */
#ao-root .ao-queue-item { cursor: default; gap: 6px; }
#ao-root .ao-item-main { display: grid; grid-template-columns: minmax(0, 1fr) auto; gap: 3px 10px; cursor: pointer; border-radius: calc(var(--ao-radius-base) - 2px); }
#ao-root .ao-item-left { display: grid; gap: 3px; align-content: start; min-width: 0; }
#ao-root .ao-item-right { display: flex; flex-direction: column; align-items: flex-end; gap: 4px; max-width: 160px; }
/* Le etichette lunghe vanno a capo invece di essere tagliate */
#ao-root .ao-item-right .ao-chip { white-space: normal; text-align: right; overflow-wrap: anywhere; }
#ao-root .ao-item-actions { display: flex; align-items: center; gap: 6px; }
#ao-root .ao-ignored-by { margin-left: auto; text-align: right; }
#ao-root .ao-inline-form { display: grid; gap: 8px; padding-top: 8px; border-top: 1px solid var(--ao-line-light); }
/* Apertura del modulo Prendi in carico: si allunga e compare con dolcezza */
#ao-root .ao-animate-expand { overflow: hidden; animation: ao-expand .35s ease-in-out; }
@keyframes ao-expand { from { max-height: 0; opacity: 0; padding-top: 0; } to { max-height: var(--ao-form-height, 600px); opacity: 1; padding-top: 8px; } }
#ao-root .ao-animate-collapse { overflow: hidden; animation: ao-collapse .2s ease-in forwards; }
/* Ignora (In coda): swipe to delete col clic. La scheda scorre a sinistra e scopre il tasto rosso;
   confermando, il tasto si allarga su tutta la scheda e poi lo spazio si chiude. */
#ao-root .ao-queue-item { position: relative; overflow: hidden; }
#ao-root .ao-queue-item > :not(.ao-swipe-action, .ao-take-bar) { transition: transform .28s cubic-bezier(.2, .8, .2, 1), opacity .2s; }
#ao-root .ao-queue-item.ao-swiped > :not(.ao-swipe-action) { transform: translateX(calc(-1 * var(--ao-swipe-width))); }
#ao-root .ao-queue-item { --ao-swipe-width: 76px; }
#ao-root .ao-swipe-action {
    position: absolute; top: 0; right: 0; bottom: 0; z-index: 1; width: var(--ao-swipe-width);
    display: grid; place-content: center; justify-items: center; gap: 3px; padding: 0; border: 0;
    background: #d13438; color: #fff; font: inherit; font-size: var(--ao-fs-xs); font-weight: var(--ao-fw-medium); cursor: pointer;
    transform: translateX(100%); transition: transform .28s cubic-bezier(.2, .8, .2, 1), width .3s cubic-bezier(.4, 0, .2, 1), background-color .15s;
}
#ao-root .ao-swiped .ao-swipe-action { transform: none; }
/* Con il tasto laterale scoperto il pulsante Ignora non serve più: sparisce */
#ao-root .ao-queue-item .ao-ignore { transition: opacity .2s, visibility .2s; }
#ao-root .ao-swiped .ao-ignore { opacity: 0; visibility: hidden; }
#ao-root .ao-swipe-action:hover, #ao-root .ao-swipe-action:focus-visible { background: #a4262c; outline: none; }
#ao-root .ao-swipe-commit { pointer-events: none; }
#ao-root .ao-swipe-commit .ao-swipe-action { width: 100%; transform: none; }
/* alla conferma il contenuto continua a scorrere a sinistra, più piano del rosso, che lo raggiunge e lo copre */
#ao-root .ao-queue-item.ao-swipe-commit > :not(.ao-swipe-action) { transform: translateX(calc(-1 * var(--ao-swipe-width) - 35%)); transition: transform .3s cubic-bezier(.4, 0, .2, 1); }
#ao-root .ao-swipe-collapse { overflow: hidden; pointer-events: none; animation: ao-slide-collapse .25s ease-in-out forwards; }
/* Prendi in carico (In coda): il tasto blu sale dal basso ("Apro il ticket…"), poi si allarga su tutta la scheda */
#ao-root .ao-take-bar {
    position: absolute; left: 0; right: 0; bottom: 0; z-index: 2;
    display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 3px;
    background: var(--ao-accent); color: #fff; font-size: var(--ao-fs-xs); font-weight: var(--ao-fw-medium);
    transform: translateY(100%); transition: transform .42s cubic-bezier(.3, .8, .2, 1);
}
#ao-root .ao-taking { pointer-events: none; }
#ao-root .ao-taking .ao-take-bar { transform: none; }
/* il contenuto sale anche lui, ma più piano del blu, che lo raggiunge e lo copre (un solo movimento) */
#ao-root .ao-taking > :not(.ao-take-bar) { transform: translateY(calc(-.35 * var(--ao-card-height, 300px))); transition: transform .42s cubic-bezier(.3, .8, .2, 1); }
#ao-root .ao-take-spinner { width: 14px; height: 14px; border: 2px solid rgba(255,255,255,.45); border-top-color: #fff; border-radius: 50%; animation: ao-spin .8s linear infinite; }
@keyframes ao-spin { to { transform: rotate(360deg); } }
@keyframes ao-slide-collapse { from { max-height: var(--ao-card-height, 200px); } to { max-height: 0; padding-top: 0; padding-bottom: 0; margin-top: -6px; border-width: 0; } }
@keyframes ao-collapse { from { max-height: var(--ao-form-height, 600px); opacity: 1; padding-top: 8px; } to { max-height: 0; opacity: 0; padding-top: 0; border-top-color: transparent; } }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-animate-expand, #ao-root .ao-animate-collapse, #ao-root .ao-swipe-collapse { animation: none; } #ao-root .ao-queue-item > *, #ao-root .ao-swipe-action, #ao-root .ao-take-bar { transition: none; } }

/* Impostazioni */
#ao-root .ao-toolbar { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
#ao-root .ao-config-card { display: grid; gap: 6px; padding: 10px 12px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); }
#ao-root .ao-config-card-head { display: flex; align-items: center; gap: 4px; }
#ao-root .ao-config-card-head .ao-config-name { flex: 1; }
#ao-root .ao-props { display: grid; grid-template-columns: auto 1fr; gap: 7px 10px; margin: 0; font-size: var(--ao-fs-md); line-height: 1.35; }
#ao-root .ao-props dt { align-self: center; color: var(--ao-text-muted); }
/* Category: spazio per due righe, etichetta centrata in altezza */
#ao-root .ao-props dd.ao-prop-two-lines { min-height: calc(2em * 1.35); }
#ao-root .ao-props dd { margin: 0; display: flex; flex-wrap: wrap; align-items: center; gap: 4px 6px; overflow-wrap: anywhere; }
#ao-root .ao-props .ao-missing { color: var(--ao-text-subtle); font-style: italic; }
#ao-root .ao-response-text { margin: 0; color: var(--ao-text-muted); font-size: var(--ao-fs-md); white-space: pre-wrap; overflow-wrap: anywhere; display: -webkit-box; -webkit-line-clamp: 5; -webkit-box-orient: vertical; overflow: hidden; }

/* Risposte nuove e più ticket sullo stesso thread */
#ao-root .ao-item-new .ao-item-title { color: var(--ao-text-accent); }
#ao-root .ao-multi { display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 8px 10px; border-radius: var(--ao-radius-base); background: var(--ao-accent-bg); font-size: var(--ao-fs-md); }
#ao-root .ao-multi-count { font-size: var(--ao-fs-lg); color: var(--ao-text-accent); }
#ao-root .ao-ticket-compact { display: grid; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); box-shadow: var(--ao-shadow-card); }
#ao-root .ao-ticket-compact-head { display: grid; gap: 3px; padding: 9px 10px; cursor: pointer; border-radius: var(--ao-radius-base); }
#ao-root .ao-toggle { margin-left: auto; color: var(--ao-text-accent); font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-medium); }
#ao-root .ao-ticket-details { display: grid; gap: 10px; padding: 10px; border-top: 1px solid var(--ao-line-light); background: var(--ao-bg); border-radius: 0 0 var(--ao-radius-base) var(--ao-radius-base); }
#ao-root .ao-add-ticket { justify-self: start; color: var(--ao-text-accent); }

/* Archivio ignorate, gruppi di ticket, anteprima dei messaggi nuovi */
#ao-root .ao-group { display: grid; gap: 6px; padding: 8px; border: 1px dashed var(--ao-line-strong); border-radius: var(--ao-radius-base); }
#ao-root .ao-group-new { border-color: var(--ao-accent); border-style: solid; }
#ao-root .ao-group-head { color: var(--ao-text-muted); font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-medium); }
#ao-root .ao-new-preview { display: grid; gap: 2px; padding: 6px 8px; border-radius: var(--ao-radius-base); background: var(--ao-surface); border: 1px solid var(--ao-accent-border); }
#ao-root .ao-item .ao-new-preview { background: var(--ao-accent-bg); }
#ao-root .ao-new-head { display: flex; align-items: baseline; gap: 6px; font-size: var(--ao-fs-sm); }
#ao-root .ao-new-from { flex: 1; font-weight: var(--ao-fw-medium); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
#ao-root .ao-new-text { color: var(--ao-text-muted); font-size: var(--ao-fs-md); display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; }

/* Collegamento mail ↔ ticket */
#ao-root .ao-via { margin-bottom: -4px; color: var(--ao-text-subtle); font-size: var(--ao-fs-sm); }
#ao-root .ao-card-quiet { background: var(--ao-bg-soft); }
#ao-root .ao-card-warning { border-color: var(--ao-warning-bg); border-left: 3px solid var(--ao-warning); }
#ao-root .ao-candidate { display: flex; align-items: center; gap: 8px; padding: 8px 10px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); }
#ao-root .ao-candidate-text { flex: 1; display: grid; gap: 2px; min-width: 0; }
#ao-root .ao-check { display: flex; align-items: center; gap: 6px; color: var(--ao-text-muted); font-size: var(--ao-fs-md); cursor: pointer; }
#ao-root .ao-check input { margin: 0; accent-color: var(--ao-accent); }

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
        applyTheme();
        fillPane();
        // Il riquadro di Outlook si può ridimensionare: l'altezza dell'area si ricalcola.
        window.addEventListener('resize', fitHeight);
        if (window.ResizeObserver) new ResizeObserver(() => fitHeight()).observe(scrollContainer());
        document.getElementById('open-amp-button')?.addEventListener('click', show);
        Outlook.onMailChanged(() => handleMailChanged());
        Services.subscribe(() => { refreshLists(); });
        state.mail = Outlook.currentMail();
        render();
        await refreshLists();
        if (state.mail?.conversationId) showMailTarget(state.mail);
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
