/*
 * AMP - collegamento del pannello (amp.js) alla Function App di jak.
 *
 * Espone window.AmpBackend con i metodi del contratto di amp.js (blocco 3), tradotti nelle chiamate alle function:
 *   ListQueue, ListThreadMessages, GetMailBody, ManageThread  → coda, ignorate, messaggi, etichetta (tag), presa in carico, visto
 *   ManageTicket, SearchTicketRecords                          → ticket (lista Ticket su SharePoint)
 *   ForwardConversation, ReplyAll                              → inoltro ai gruppi esterni, risposta con la sigla
 * Token: window.WorkflowHealth.getAccessTokenSilently() (add-in di jak). Operatore = utente di Outlook.
 * Il thread si identifica con ParentItemId (MailEvents); il pannello usa conversationId: la corrispondenza
 * si tiene qui. I ticket di un thread non ancora visto in coda usano la chiave "parent:<ParentItemId>";
 * i ticket senza mail (SHP-M-<n>) non hanno thread.
 *
 * L'etichetta di chi gestisce un thread in coda è lo stato Working su MailEvents (ManageThread tag / untag):
 * presa in carico, inoltro e ignora la richiedono.
 * Non ancora disponibili (nessuna function): ServiceNow (numero INC, storico, ricerca su AMP), collegamento manuale
 * di un ticket a una mail.
 */
(function () {
    'use strict';

    const API = 'https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/';
    const CONFIG_KEY = 'test-plugin-operational-config';
    const NOTES_MAX = 2000;
    const STATE_CODES = { 'In Progress': '2', 'On Hold': '3', Resolved: '6', Canceled: '8' };
    const OPEN_STATUSES = ['In Progress', 'On Hold'];
    // Ticket su AMP (ServiceNow) dal numero INC: niente sys_id, l'istanza apre l'incident con quel numero.
    const ampUrl = number => (/^INC\d+$/i.test(String(number || '').trim())
        ? `https://amp.service-now.com/nav_to.do?uri=${encodeURIComponent(`incident.do?sysparm_query=number=${String(number).trim().toUpperCase()}`)}` : '');

    const threadOf = new Map(); // conversationId → ParentItemId
    const conversationOf = new Map(); // ParentItemId → conversationId
    const messageCounts = new Map(); // conversationId → mail del thread viste dal pannello (controllo delle azioni sul thread)

    // ------------------------------------------------------------------
    // Chiamate
    // ------------------------------------------------------------------

    // Token rifiutato dalla Function (401): si scarta e si riprova una volta con un token nuovo.
    async function call(name, body, method = 'POST', retried = false) {
        const log = window.AmpActionLog;
        const tokenStarted = Date.now();
        const authentication = await window.WorkflowHealth?.getAccessTokenSilently?.();
        log?.record('function', `${name}:token`, { ms: Date.now() - tokenStarted, mode: authentication?.acquisitionMode || null, ok: Boolean(authentication?.accessToken) });
        if (!authentication?.accessToken) throw new Error('Accesso non disponibile: riaprire il pannello.');
        rememberTokenUser(authentication.accessToken);
        const headers = { Authorization: `Bearer ${authentication.accessToken}`, Accept: 'application/json' };
        let url = API + name;
        const request = { method, headers };
        if (method === 'GET') url += `?${new URLSearchParams(body || {})}`;
        else {
            headers['Content-Type'] = 'application/json';
            request.body = JSON.stringify(body || {});
        }
        const fetchStarted = Date.now();
        log?.record('function', `${name}:fetch`, body?.action || null);
        const response = await fetch(url, request).catch((error) => {
            log?.record('function', `${name}:network`, { ms: Date.now() - fetchStarted, message: error?.message || String(error) });
            throw error;
        });
        log?.record('function', `${name}:${response.status}`, { ms: Date.now() - fetchStarted });
        if (response.status === 401 && !retried && window.WorkflowHealth?.invalidateToken) {
            window.WorkflowHealth.invalidateToken('functionResponse');
            return call(name, body, method, true);
        }
        const text = await response.text();
        let data = null;
        try { data = text ? JSON.parse(text) : null; } catch { /* risposta non JSON */ }
        if (!response.ok || data?.status === 'error') {
            const error = new Error(data?.message || `${name}: errore ${response.status}`);
            error.reason = data?.reason;
            error.data = data;
            throw error;
        }
        return data;
    }

    // Frequenze degli aggiornamenti: config.refresh (uguale per tutti), sostituibile sul proprio pannello dai Developer
    // Tools (salvato nel browser). Valori fuori dai limiti si riportano dentro.
    const REFRESH_DEFAULTS = { listsSeconds: 60, threadCacheSeconds: 15 };
    const REFRESH_LIMITS = { listsSeconds: [15, 900], threadCacheSeconds: [0, 300] };
    const REFRESH_OVERRIDE_KEY = 'amp_panel_refresh_override_v1';
    const refreshListeners = new Set();
    function refreshOverride() {
        try { return JSON.parse(localStorage.getItem(REFRESH_OVERRIDE_KEY) || 'null') || {}; } catch { return {}; }
    }
    function refreshSettings() {
        const config = pluginConfig().refresh || {};
        const override = refreshOverride();
        const pick = name => {
            const value = Number(override[name] ?? config[name] ?? REFRESH_DEFAULTS[name]);
            const [min, max] = REFRESH_LIMITS[name];
            return Number.isFinite(value) ? Math.min(max, Math.max(min, Math.round(value))) : REFRESH_DEFAULTS[name];
        };
        const source = Object.keys(override).length ? 'pannello' : Object.keys(config).length ? 'config' : 'predefiniti';
        return { listsSeconds: pick('listsSeconds'), threadCacheSeconds: pick('threadCacheSeconds'), source, defaults: REFRESH_DEFAULTS, limits: REFRESH_LIMITS };
    }
    // values null: si torna alla config. Il nuovo intervallo vale subito.
    function setRefreshOverride(values) {
        const clean = Object.fromEntries(Object.keys(REFRESH_DEFAULTS)
            .filter(name => values && values[name] !== '' && values[name] !== null && values[name] !== undefined && Number.isFinite(Number(values[name])))
            .map(name => [name, Number(values[name])]));
        try {
            if (Object.keys(clean).length) localStorage.setItem(REFRESH_OVERRIDE_KEY, JSON.stringify(clean));
            else localStorage.removeItem(REFRESH_OVERRIDE_KEY);
        } catch { /* storage non disponibile */ }
        window.AmpActionLog?.record('refresh', 'settings', refreshSettings());
        refreshListeners.forEach(restart => restart());
        return refreshSettings();
    }

    function pluginConfig() {
        try { return JSON.parse(localStorage.getItem(CONFIG_KEY) || 'null')?.payload || {}; }
        catch { return {}; }
    }
    // Chi ha fatto l'accesso all'add-in, letto dal token (ricordato per l'avvio, prima del primo token).
    const TOKEN_USER_KEY = 'amp_panel_token_user_v1';
    let tokenUser = '';
    try { tokenUser = localStorage.getItem(TOKEN_USER_KEY) || ''; } catch { /* storage non disponibile */ }
    function rememberTokenUser(accessToken) {
        try {
            const part = String(accessToken).split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
            const claims = JSON.parse(decodeURIComponent(escape(atob(part.padEnd(Math.ceil(part.length / 4) * 4, '=')))));
            const user = String(claims.preferred_username || claims.upn || claims.email || claims.unique_name || '').trim().toLowerCase();
            if (!user || user === tokenUser) return;
            tokenUser = user;
            try { localStorage.setItem(TOKEN_USER_KEY, user); } catch { /* storage non disponibile */ }
        } catch { /* token non leggibile: resta l'email di Outlook */ }
    }
    // Operatore che usa il pannello: chi ha fatto l'accesso, non la casella aperta in Outlook (su Outlook desktop una
    // casella condivisa aggiunta come account risulta come utente). Vale quello dei due che è tra gli operatori della config.
    function me() {
        const outlook = window.Office?.context?.mailbox?.userProfile?.emailAddress
            || window.TestPluginAPI?.getCurrentMessageContext?.()?.mailbox?.userEmailAddress || '';
        const operators = Object.values(pluginConfig().operators || {});
        const known = email => email && operators.some(operator => sameEmail(operator?.email, email));
        const email = [tokenUser, outlook].find(known) || tokenUser || outlook;
        if (!email) throw new Error('Utente di Outlook non disponibile.');
        return email;
    }
    const sameEmail = (a, b) => String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
    // Etichetta di un thread in coda: chi ha il tag (Working). I colleghi in grigio; la propria la colora il pannello.
    // Tag = chi ci lavora (AssignedTo) su un thread in coda (Working) o ignorato.
    function threadLabel(thread) {
        if (!['Working', 'Ignored'].includes(thread?.status) || !thread.assignedTo) return null;
        return { name: operatorName(thread.assignedTo), color: 'grafite', mine: sameEmail(thread.assignedTo, me()), at: thread.statusAt || null };
    }
    // Gestione di una mail (la sua riga su MailEvents), per gli esiti: in carico, ignorata, inoltrata, risolta.
    // Inoltrata: AssignedTo è il gruppo; in carico: chi ce l'ha; ignorata e risolta: chi l'ha decisa.
    function handledOf(message) {
        if (!['Acquired', 'Ignored', 'Forwarded', 'Solved'].includes(message?.status)) return null;
        const by = message.status === 'Forwarded' ? message.assignedTo || '' : operatorName(message.status === 'Acquired' ? message.assignedTo : message.statusBy);
        return { status: message.status, by, at: message.statusAt || null };
    }
    function operatorName(email) {
        const entry = Object.values(pluginConfig().operators || {})
            .find(operator => String(operator?.email || '').toLowerCase() === String(email || '').toLowerCase());
        return entry?.displayName || email || '';
    }
    // "Nome <email>" oppure solo l'email.
    function sender(from) {
        const text = String(from || '').trim();
        const match = text.match(/^(.*?)\s*<([^>]+)>$/);
        if (match) return { name: match[1].replace(/^"|"$/g, '') || match[2], email: match[2] };
        return { name: text, email: text.includes('@') ? text : '' };
    }

    // ------------------------------------------------------------------
    // Thread: conversationId ↔ ParentItemId
    // ------------------------------------------------------------------

    const normalize = id => String(id || '').replace(/\//g, '-').replace(/\+/g, '_');
    function remember(conversationId, parentItemId) {
        if (!conversationId || !parentItemId) return;
        threadOf.set(conversationId, Number(parentItemId));
        conversationOf.set(Number(parentItemId), conversationId);
    }
    // Ticket senza mail: nessun thread (ParentItemId vuoto), quindi nessuna conversazione.
    const keyOf = parentItemId => (parentItemId ? conversationOf.get(Number(parentItemId)) || `parent:${parentItemId}` : '');
    async function parentOf(conversationId) {
        if (String(conversationId).startsWith('parent:')) return Number(String(conversationId).slice(7));
        if (threadOf.has(conversationId)) return threadOf.get(conversationId);
        const result = await call('GetRecordByConversationId', { conversationId: normalize(conversationId), mode: 'first' }, 'GET');
        const parent = Number(result?.fields?.ParentItemId || result?.itemId);
        if (!result?.found || !parent) throw new Error('Thread non trovato.');
        remember(conversationId, parent);
        return parent;
    }

    // ------------------------------------------------------------------
    // Ticket della lista Ticket nel formato del pannello
    // ------------------------------------------------------------------

    function toTicket(record, extra = {}) {
        const closed = !OPEN_STATUSES.includes(record.status);
        return {
            number: record.ticketId,
            sysId: record.id,
            url: ampUrl(record.ampReference),
            state: STATE_CODES[record.status] || '',
            assignedTo: { name: operatorName(record.operator), sysId: record.operator || '' },
            companyName: record.agency || '',
            activityName: '',
            shortDescription: record.shortDescription || '',
            description: record.description || '',
            reportedOn: record.created || '',
            openedAt: record.created || '',
            closedAt: closed ? record.modified || '' : '',
            timeWorkedMinutes: Number(record.elapsedTime) || 0,
            holdReason: '',
            holdExpiration: '',
            resolutionNotes: record.resolutionNotes || '',
            notes: record.notes || '',
            ampReference: record.ampReference || '',
            conversationId: keyOf(record.parentItemId),
            parentItemId: record.parentItemId,
            messageId: record.messageId || null,
            owner: record.operator || null,
            newMessages: 0,
            lastNewMessage: null,
            lastMessageAt: null,
            forwardGroup: record.forwardGroup || '',
            repliedAt: record.repliedAt || '',
            ...extra
        };
    }
    async function searchTickets(filters) {
        const result = await call('SearchTicketRecords', filters);
        return result?.records || [];
    }
    async function ticketRecord(number) {
        const [record] = await searchTickets({ ticketId: number });
        if (!record) throw new Error(`Ticket ${number} non trovato.`);
        return record;
    }
    async function editTicket(number, changes, force = false) {
        const result = await call('ManageTicket', { action: 'editTicket', ticketId: number, caller: me(), force, ...changes });
        return toTicket(result.ticket);
    }
    // Note: si aggiunge in coda con data e autore, tenendo gli ultimi caratteri ammessi.
    function appendNote(current, text, author) {
        const stamp = new Date().toLocaleString('it-IT', { dateStyle: 'short', timeStyle: 'short' });
        const next = [current, `[${stamp} · ${author}] ${text}`].filter(Boolean).join('\n');
        return next.length > NOTES_MAX ? next.slice(next.length - NOTES_MAX) : next;
    }

    // ------------------------------------------------------------------
    // Coda
    // ------------------------------------------------------------------

    function queueItem(thread) {
        remember(thread.conversationId, thread.parentItemId);
        messageCounts.set(thread.conversationId, Number(thread.messageCount) || 0);
        const last = thread.lastMessage || {};
        return {
            conversationId: thread.conversationId,
            parentItemId: thread.parentItemId,
            messageId: last.messageId,
            subject: thread.subject || last.subject || '',
            mailbox: pluginConfig().ticketing?.email || '',
            sender: sender(last.from || thread.from),
            receivedAt: thread.lastReceived,
            preview: '',
            label: threadLabel(thread.thread),
            history: thread.previous || null
        };
    }
    // Con le azioni che cambiano il thread va il numero di mail che il pannello conosce: se nel frattempo ne è arrivata
    // un'altra, la function si ferma (thread_changed) e il pannello si aggiorna.
    const COUNTED_ACTIONS = ['tag', 'acquire', 'ignore', 'restore'];
    const threadAction = async (conversationId, action, extra = {}) => {
        const known = COUNTED_ACTIONS.includes(action) ? messageCounts.get(conversationId) : null;
        return call('ManageThread', { action, parentItemId: await parentOf(conversationId), caller: me(), ...(known ? { knownCount: known } : {}), ...extra });
    };

    // ------------------------------------------------------------------
    // Contratto del pannello
    // ------------------------------------------------------------------

    const AmpBackend = {
        mode: 'functions', // amp.js: pannello di produzione (scheda Debug al posto delle Impostazioni)
        async listQueue() {
            const result = await call('ListQueue', { section: 'active' }, 'GET');
            return (result?.threads || []).map(queueItem);
        },
        async listIgnored() {
            const result = await call('ListQueue', { section: 'ignored' }, 'GET');
            return (result?.threads || []).map(thread => ({
                ...queueItem(thread),
                discardedBy: operatorName(thread.thread?.statusBy),
                discardedAt: thread.thread?.statusAt || null
            }));
        },
        async markThreadSeen(conversationId) {
            await threadAction(conversationId, 'seen');
        },
        // Ticket aperti, con i messaggi non visti dei thread in carico a chi usa il pannello.
        async listOpenTickets() {
            const [records, mine] = await Promise.all([
                searchTickets({ statuses: OPEN_STATUSES }),
                call('ListQueue', { section: 'mine', caller: me() }, 'GET').catch(() => null)
            ]);
            const unseen = new Map();
            (mine?.threads || []).forEach(thread => {
                remember(thread.conversationId, thread.parentItemId);
                unseen.set(Number(thread.parentItemId), thread);
            });
            return records.map(record => {
                const thread = unseen.get(Number(record.parentItemId));
                const last = thread?.lastUnseen;
                return toTicket(record, thread ? {
                    newMessages: thread.unseenCount || 0,
                    lastNewMessage: last ? { from: sender(last.from).name, at: last.receivedDate, preview: last.subject || '' } : null,
                    lastMessageAt: thread.lastReceived || null
                } : {});
            });
        },
        // Mail selezionata in Outlook: stato del thread su MailEvents (ultima riga) e i suoi ticket.
        async findTicketForMail(mail) {
            const result = await call('GetRecordByConversationId', { conversationId: normalize(mail.conversationId), mode: 'all' }, 'GET');
            const rows = (result?.items || []).slice().sort((a, b) => Number(a.itemId) - Number(b.itemId));
            if (!result?.found || !rows.length) return { tickets: [], thread: { status: 'none' } };
            const parentItemId = Number(rows[0].fields?.ParentItemId || rows[0].itemId);
            remember(mail.conversationId, parentItemId);
            const last = rows[rows.length - 1].fields || {};
            const records = await searchTickets({ parentItemId });
            return {
                tickets: records.map(record => toTicket(record)),
                thread: {
                    status: last.Status || 'New',
                    assignedTo: last.AssignedTo || null,
                    assignedToName: last.AssignedTo ? operatorName(last.AssignedTo) : null,
                    statusBy: last.StatusBy || null,
                    statusByName: last.StatusBy ? operatorName(last.StatusBy) : null,
                    statusAt: last.StatusAt || null,
                    messageCount: rows.length,
                    parentItemId
                }
            };
        },
        // Presa in carico del thread, poi il ticket dalla mail scelta. knownCount: le mail viste in coda; se nel
        // frattempo ne è arrivata un'altra la presa in carico non riesce (thread_changed) e la coda si rilegge.
        // Senza mail (mail null): ticket SHP-M-<n>, senza thread da prendere in carico.
        async createTicket(mail, payload) {
            const parentItemId = mail ? await parentOf(mail.conversationId) : null;
            if (mail) await threadAction(mail.conversationId, 'acquire');
            const fields = payload?.fields || {};
            const result = await call('ManageTicket', {
                action: 'createTicket',
                ...(mail ? { parentItemId, messageId: mail.itemId || mail.messageId } : {}),
                operator: me(),
                caller: me(),
                status: 'In Progress',
                shortDescription: fields.shortDescription || mail?.subject || '',
                description: String(fields.description || mail?.body || '').slice(0, 4000),
                agency: payload?.companyName || '',
                elapsedTime: String(Number(fields.timeWorkedMinutes) || 0),
                ...(payload?.forwardGroup ? { forwardGroup: payload.forwardGroup } : {})
            });
            if (mail) remember(mail.conversationId, parentItemId);
            return toTicket(result.ticket);
        },
        async assignTicket(number) {
            return editTicket(number, { operator: me() }, true);
        },
        async discardMail(conversationId) {
            await threadAction(conversationId, 'ignore');
        },
        async restoreMail(conversationId) {
            await threadAction(conversationId, 'restore');
        },
        async forwardMail(conversationId, group) {
            const parentItemId = await parentOf(conversationId);
            await call('ForwardConversation', { conversationId: normalize(conversationId), parentItemId, groupName: group, caller: me() });
        },
        // Rispondi a tutti con la presa in carico; tag: sigla del gruppo nell'oggetto; ticket: dove segnare la risposta.
        async replyAll(conversationId, { tag, reference, ticket } = {}) {
            const parentItemId = await parentOf(conversationId);
            await call('ReplyAll', { parentItemId, caller: me(), tag, reference, ticketId: ticket });
        },
        // Attività della scheda Utility: la esegue RunUtility sul tenant dell'agenzia.
        async runUtility(operationId, agencyId, params) {
            return call('RunUtility', { operationId, agencyId, params, caller: me() });
        },
        // Testo e azione sul ticket (lista Ticket; ServiceNow non ancora collegato).
        async updateTicket(number, action, text, user) {
            const record = await ticketRecord(number);
            const author = user?.displayValue || operatorName(me());
            const changes = {};
            if (text) {
                if (action.type === 'resolve' || action.type === 'cancel') changes.resolutionNotes = text.slice(0, NOTES_MAX);
                else changes.notes = appendNote(record.notes, text, author);
            }
            if (action.type === 'resolve') changes.status = 'Resolved';
            if (action.type === 'cancel') changes.status = 'Canceled';
            if (action.type === 'hold') changes.status = 'On Hold';
            if (action.type === 'comment' && record.status === 'On Hold') changes.status = 'In Progress';
            if (action.addMinutes) changes.elapsedTime = String(Math.max(0, (Number(record.elapsedTime) || 0) + Number(action.addMinutes)));
            if (!Object.keys(changes).length) return toTicket(record);
            return editTicket(number, changes);
        },
        // Ricerca: per ora sulla lista Ticket (numero, riferimento AMP, short description, descrizione).
        async searchTickets(text) {
            const records = await searchTickets({ text, resultSize: 20 });
            return records.map(record => ({ ...toTicket(record), readOnly: true, lastWorkNote: null }));
        },
        async getMailBody(messageId) {
            const result = await call('GetMailBody', { messageId }, 'GET');
            return result?.newText || result?.text || '';
        },
        async listThreadMessages(conversationId) {
            const result = await call('ListThreadMessages', { parentItemId: await parentOf(conversationId) }, 'GET');
            const messages = result?.messages || [];
            messageCounts.set(conversationId, messages.length);
            return messages.map(message => ({
                messageId: message.messageId,
                sender: sender(message.from),
                receivedAt: message.receivedDate,
                preview: message.subject || '',
                handled: handledOf(message)
            }));
        },
        // Etichetta = tag sul thread (Working). Se c'è già quella di un collega non si scrive niente e si restituisce la sua;
        // force la sostituisce (subentro esplicito).
        async labelMail(conversationId, by, color, { force = false } = {}) {
            try {
                const result = await threadAction(conversationId, 'tag', force ? { force: true } : {});
                return { ok: true, label: { ...threadLabel(result.thread), name: by?.displayValue || operatorName(me()), color, mine: true } };
            } catch (error) {
                if (error.reason !== 'tagged_by_other') throw error;
                return { ok: false, label: threadLabel(error.data?.thread) };
            }
        },
        async unlabelMail(conversationId) {
            await threadAction(conversationId, 'untag');
            return { ok: true };
        },
        refreshSettings,
        operationalConfig: () => pluginConfig(),
        currentUserEmail: () => { try { return me(); } catch { return ''; } },
        // nome dell'operatore che usa il pannello, se è nella config (altrimenti vuoto)
        currentUserName: () => { try { const email = me(); const name = operatorName(email); return name === email ? '' : name; } catch { return ''; } }, // config dell'add-in (agenzie e operazioni abilitate per la scheda Utility)
        setRefreshOverride,
        // Aggiornamento: ogni refresh.listsSeconds e quando il pannello torna in primo piano.
        subscribe(listener) {
            const notify = () => { if (document.visibilityState === 'visible') listener(); };
            let timer = null;
            const schedule = () => {
                clearTimeout(timer);
                timer = setTimeout(() => { notify(); schedule(); }, refreshSettings().listsSeconds * 1000);
            };
            schedule();
            refreshListeners.add(schedule);
            document.addEventListener('visibilitychange', notify);
            return () => {
                clearTimeout(timer);
                refreshListeners.delete(schedule);
                document.removeEventListener('visibilitychange', notify);
            };
        }
    };

    window.AmpBackend = AmpBackend;
})();
