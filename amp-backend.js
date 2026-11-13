/*
 * AMP - collegamento del pannello (amp.js) alla Function App di jak.
 *
 * Espone window.AmpBackend con i metodi del contratto di amp.js (blocco 3), tradotti nelle chiamate alle function:
 *   ListQueue, ListThreadMessages, GetMailBody, ManageThread  → coda, ignorate, messaggi, presa in carico, visto
 *   ManageTicket, SearchTicketRecords                          → ticket (lista Ticket su SharePoint)
 *   ForwardConversation                                        → inoltro ai gruppi esterni
 * Token: window.WorkflowHealth.getAccessTokenSilently() (add-in di jak). Operatore = utente di Outlook.
 * Il thread si identifica con ParentItemId (MailEvents); il pannello usa conversationId: la corrispondenza
 * si tiene qui. I ticket di un thread non ancora visto in coda usano la chiave "parent:<ParentItemId>".
 *
 * Non ancora disponibili (nessuna function): ServiceNow (numero INC, storico, ricerca su AMP), collegamento manuale
 * di un ticket a una mail. Le etichette restano solo in questo pannello.
 */
(function () {
    'use strict';

    const API = 'https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/';
    const CONFIG_KEY = 'test-plugin-operational-config';
    const POLL_MS = 60 * 1000;
    const NOTES_MAX = 2000;
    const STATE_CODES = { 'In Progress': '2', 'On Hold': '3', Resolved: '6', Canceled: '8' };
    const OPEN_STATUSES = ['In Progress', 'On Hold'];

    const threadOf = new Map(); // conversationId → ParentItemId
    const conversationOf = new Map(); // ParentItemId → conversationId
    const labels = new Map(); // conversationId → { name, color, at } (solo locale)

    // ------------------------------------------------------------------
    // Chiamate
    // ------------------------------------------------------------------

    async function call(name, body, method = 'POST') {
        const authentication = await window.WorkflowHealth?.getAccessTokenSilently?.();
        if (!authentication?.accessToken) throw new Error('Accesso non disponibile: riaprire il pannello.');
        const headers = { Authorization: `Bearer ${authentication.accessToken}`, Accept: 'application/json' };
        let url = API + name;
        const request = { method, headers };
        if (method === 'GET') url += `?${new URLSearchParams(body || {})}`;
        else {
            headers['Content-Type'] = 'application/json';
            request.body = JSON.stringify(body || {});
        }
        const response = await fetch(url, request);
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

    function pluginConfig() {
        try { return JSON.parse(localStorage.getItem(CONFIG_KEY) || 'null')?.payload || {}; }
        catch { return {}; }
    }
    // Operatore che usa il pannello: l'email di Outlook.
    function me() {
        const email = window.Office?.context?.mailbox?.userProfile?.emailAddress
            || window.TestPluginAPI?.getCurrentMessageContext?.()?.mailbox?.userEmailAddress || '';
        if (!email) throw new Error('Utente di Outlook non disponibile.');
        return email;
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
    const keyOf = parentItemId => conversationOf.get(Number(parentItemId)) || `parent:${parentItemId}`;
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
            url: '',
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
            syncedAt: record.modified || null,
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
            previous: null,
            candidates: [],
            label: labels.get(thread.conversationId) || null,
            history: thread.previous || null
        };
    }
    const threadAction = async (conversationId, action, extra = {}) =>
        call('ManageThread', { action, parentItemId: await parentOf(conversationId), caller: me(), ...extra });

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
                discardedBy: operatorName(thread.thread?.assignedTo),
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
            if (!result?.found || !rows.length) return { tickets: [], previous: null, candidates: [], discarded: false, thread: { status: 'none' } };
            const parentItemId = Number(rows[0].fields?.ParentItemId || rows[0].itemId);
            remember(mail.conversationId, parentItemId);
            const last = rows[rows.length - 1].fields || {};
            const records = await searchTickets({ parentItemId });
            const open = records.filter(record => OPEN_STATUSES.includes(record.status));
            const closed = records.filter(record => !OPEN_STATUSES.includes(record.status));
            return {
                tickets: records.map(record => ({ ...toTicket(record), via: 'link' })),
                previous: closed.length && !open.length ? toTicket(closed[0]) : null,
                candidates: [],
                discarded: false,
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
        async getTicket(number) {
            return toTicket(await ticketRecord(number));
        },
        async linkTicket() {
            throw new Error('Collegamento manuale non ancora disponibile.');
        },
        // Presa in carico del thread, poi il ticket dalla mail scelta.
        async createTicket(mail, payload) {
            if (!mail) throw new Error('Ticket senza mail non ancora disponibile.');
            const parentItemId = await parentOf(mail.conversationId);
            await call('ManageThread', { action: 'acquire', parentItemId, caller: me() });
            const fields = payload?.fields || {};
            const result = await call('ManageTicket', {
                action: 'createTicket',
                parentItemId,
                messageId: mail.itemId || mail.messageId,
                operator: me(),
                caller: me(),
                status: 'In Progress',
                shortDescription: fields.shortDescription || mail.subject || '',
                description: String(fields.description || mail.body || '').slice(0, 4000),
                agency: payload?.companyName || '',
                elapsedTime: String(Number(fields.timeWorkedMinutes) || 0)
            });
            remember(mail.conversationId, parentItemId);
            return toTicket(result.ticket);
        },
        async assignTicket(number) {
            return editTicket(number, { operator: me() }, true);
        },
        // Storico: per ora solo le note del ticket.
        async getTicketActivity(number, { offset = 0, limit = 5 } = {}) {
            const record = await ticketRecord(number);
            const entries = String(record.notes || '').split('\n').filter(Boolean).reverse()
                .map(line => ({ at: record.modified, author: '', kind: 'comment', text: line }));
            return { entries: entries.slice(offset, offset + limit), total: entries.length };
        },
        async discardMail(conversationId) {
            await threadAction(conversationId, 'ignore');
            labels.delete(conversationId);
        },
        async restoreMail(conversationId) {
            await threadAction(conversationId, 'restore');
        },
        async forwardMail(conversationId, group) {
            const parentItemId = await parentOf(conversationId);
            await call('ForwardConversation', { conversationId: normalize(conversationId), parentItemId, groupName: group, caller: me() });
            labels.delete(conversationId);
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
            return (result?.messages || []).map(message => ({
                messageId: message.messageId,
                sender: sender(message.from),
                receivedAt: message.receivedDate,
                preview: message.subject || ''
            }));
        },
        // Etichette: solo in questo pannello (nessuna function).
        async labelMail(conversationId, by, color) {
            const label = { name: by?.displayValue || operatorName(me()), color, at: new Date().toISOString() };
            labels.set(conversationId, label);
            return { ok: true, label };
        },
        async unlabelMail(conversationId) {
            labels.delete(conversationId);
            return { ok: true };
        },
        // Aggiornamento: ogni minuto e quando il pannello torna in primo piano.
        subscribe(listener) {
            const notify = () => { if (document.visibilityState === 'visible') listener(); };
            const timer = setInterval(notify, POLL_MS);
            document.addEventListener('visibilitychange', notify);
            return () => {
                clearInterval(timer);
                document.removeEventListener('visibilitychange', notify);
            };
        }
    };

    window.AmpBackend = AmpBackend;
})();
