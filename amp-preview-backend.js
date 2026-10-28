/*
 * Servizi SIMULATI per l'anteprima dell'area AMP (amp-preview.html).
 * Non fa parte dell'add-in: taskpane.html non carica questo file.
 *
 * Simula, in memoria, le tre cose che nell'add-in vero staranno dietro la Function App:
 *   - la casella condivisa di Outlook (mail, conversazioni, tag = categorie);
 *   - la lista SharePoint "Richieste AMP" (una riga per conversazione: chi la gestisce,
 *     ticket collegato, azione da eseguire all'invio della risposta);
 *   - i ticket di ServiceNow con il loro storico (commenti e note interne).
 *
 * Espone window.AmpBackend con gli stessi metodi che amp.js si aspetta dal servizio vero
 * (vedi il blocco 3 di amp.js). Ogni metodo restituisce una Promise e copie dei dati.
 * In più, solo per l'anteprima, AmpBackend.mailbox permette alla finta Outlook di
 * leggere le mail e inviare risposte.
 */
(function () {
    'use strict';

    const STORAGE_KEY = 'amp_preview_backend_v2';
    const SERVICENOW_URL = 'https://amp.service-now.com/incident.do?sys_id=';
    const DELAY_MS = 220;
    const listeners = new Set();

    // ------------------------------------------------------------------
    // Dati di partenza (tutti inventati)
    // ------------------------------------------------------------------
    const MINUTE = 60 * 1000;
    const HOUR = 60 * MINUTE;
    const DAY = 24 * HOUR;
    const ago = ms => new Date(Date.now() - ms).toISOString();
    const pad = value => String(value).padStart(2, '0');
    function snDate(date) {
        return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
    }
    const SHARED_MAILBOX = { name: 'Supporto AMP', email: 'supporto.amp@example.com' };

    function seed() {
        const messages = [];
        const requests = [];
        const tickets = {};
        let messageCounter = 0;
        const message = (conversationId, from, subject, body, receivedAt, extra = {}) => {
            messageCounter += 1;
            const item = {
                id: `msg-${messageCounter}`,
                conversationId,
                from,
                to: [SHARED_MAILBOX],
                subject,
                body,
                receivedAt,
                categories: [],
                direction: 'in',
                ...extra
            };
            messages.push(item);
            return item;
        };
        const request = (conversation, extra = {}) => {
            const first = conversation[0];
            const record = {
                id: `req-${first.conversationId}`,
                conversationId: first.conversationId,
                messageId: conversation[conversation.length - 1].id,
                firstMessageId: first.id,
                subject: first.subject,
                sender: first.from,
                receivedAt: first.receivedAt,
                lastMessageAt: conversation[conversation.length - 1].receivedAt,
                preview: first.body.replace(/\s+/g, ' ').slice(0, 140),
                owner: '',
                status: 'queue',
                ticket: null,
                pendingAction: null,
                ...extra
            };
            requests.push(record);
            return record;
        };
        const tag = (conversation, owner) => conversation.forEach(item => { item.categories = [owner]; });

        // 1. Richiesta nuova, nessuno l'ha presa in carico
        const c1 = [message('conv-1', { name: 'Mario Bianchi', email: 'mario.bianchi@agenziaalfa.it' },
            'Casella di posta piena',
            'Buongiorno,\nla casella di posta dell\'ufficio risulta piena e non riceviamo più email dai clienti.\nPotete aiutarci?\n\nGrazie\nMario Bianchi\nAgenzia Alfa', ago(25 * MINUTE))];
        request(c1);

        // 2. Presa in carico da Marco, ticket ancora da aprire
        const c2 = [message('conv-2', { name: 'Ufficio Beta', email: 'ufficio@agenziabeta.it' },
            'Reset password utente Verdi',
            'Salve,\nil collega Luigi Verdi non riesce più ad accedere al gestionale: chiede il reset della password.\n\nCordiali saluti\nAgenzia Beta', ago(2 * HOUR))];
        tag(c2, 'Marco');
        request(c2, { owner: 'Marco' });

        // 3. Mail informativa: probabilmente non serve un ticket
        const c3 = [message('conv-3', { name: 'Fornitore Servizi', email: 'newsletter@fornitore.example' },
            'Aggiornamento condizioni di servizio',
            'Gentile cliente,\nla informiamo che dal prossimo mese entreranno in vigore le nuove condizioni di servizio.\nNessuna azione è richiesta.', ago(3 * HOUR))];
        request(c3);

        // 4. Ticket In Progress di Marco, con uno storico lungo
        const c4 = [
            message('conv-4', { name: 'Segreteria Comune Gamma', email: 'segreteria@comune-gamma.it' },
                'Stampante secondo piano non stampa',
                'Buongiorno,\nla stampante del secondo piano non stampa più da stamattina. Sul display compare "errore 50.4".\n\nSegreteria', ago(26 * HOUR)),
            message('conv-4', { name: 'Marco', email: SHARED_MAILBOX.email },
                'RE: Stampante secondo piano non stampa',
                'Buongiorno,\nabbiamo preso in carico la segnalazione. Potete spegnere e riaccendere la stampante e dirci se l\'errore resta?\n\nSaluti,\nSupporto AMP', ago(25 * HOUR), { direction: 'out', to: [{ name: 'Segreteria Comune Gamma', email: 'segreteria@comune-gamma.it' }] }),
            message('conv-4', { name: 'Segreteria Comune Gamma', email: 'segreteria@comune-gamma.it' },
                'RE: Stampante secondo piano non stampa',
                'L\'abbiamo riavviata ma l\'errore resta uguale.', ago(5 * HOUR))
        ];
        tag(c4, 'Marco');
        const r4 = request(c4, { owner: 'Marco', status: 'ticket' });
        r4.ticket = {
            number: 'INC0012101', sysId: 'sim-12101', url: `${SERVICENOW_URL}sim-12101`, state: '2',
            companyId: 'gamma', companyName: 'Comune Gamma', activityId: 'stampante', activityName: 'Stampante',
            shortDescription: 'Stampante secondo piano non stampa', timeWorkedMinutes: 30,
            holdReason: '', holdExpiration: '', openedBy: 'Marco', openedAt: ago(25.5 * HOUR)
        };
        tickets[r4.id] = [
            { at: ago(25.5 * HOUR), author: 'Marco', kind: 'system', text: 'Ticket aperto dall\'add-in (Comune Gamma · Stampante).' },
            { at: ago(25.5 * HOUR), author: 'Segreteria Comune Gamma', kind: 'comment', text: c4[0].body },
            { at: ago(25 * HOUR), author: 'Marco', kind: 'comment', text: c4[1].body },
            { at: ago(24 * HOUR), author: 'Marco', kind: 'work_note', text: 'Errore 50.4: probabile problema di alimentazione. Verificare con il fornitore se persiste.' },
            { at: ago(20 * HOUR), author: 'Marco', kind: 'system', text: 'Tempo lavorato: +15 minuti (totale 30).' },
            { at: ago(5 * HOUR), author: 'Segreteria Comune Gamma', kind: 'comment', text: c4[2].body },
            { at: ago(4.5 * HOUR), author: 'Marco', kind: 'work_note', text: 'Chiesto al fornitore un intervento in sede.' }
        ];

        // 5. Ticket On Hold di Giulia, scadenza passata
        const c5 = [message('conv-5', { name: 'Anna Neri', email: 'a.neri@agenziaalfa.it' },
            'Inoltro automatico non funziona',
            'Ciao,\nl\'inoltro automatico delle mail verso il cellulare non funziona più.\nAnna', ago(2 * DAY))];
        tag(c5, 'Giulia');
        const r5 = request(c5, { owner: 'Giulia', status: 'ticket' });
        r5.ticket = {
            number: 'INC0012095', sysId: 'sim-12095', url: `${SERVICENOW_URL}sim-12095`, state: '3',
            companyId: 'alfa', companyName: 'Agenzia Alfa', activityId: 'posta', activityName: 'Posta elettronica',
            shortDescription: '[FWD] Inoltro automatico non funziona', timeWorkedMinutes: 15,
            holdReason: '7', holdExpiration: snDate(new Date(Date.now() - 3 * HOUR)), openedBy: 'Giulia', openedAt: ago(2 * DAY - HOUR)
        };
        tickets[r5.id] = [
            { at: ago(2 * DAY - HOUR), author: 'Giulia', kind: 'system', text: 'Ticket aperto dall\'add-in (Agenzia Alfa · Posta elettronica).' },
            { at: ago(2 * DAY - HOUR), author: 'Anna Neri', kind: 'comment', text: c5[0].body },
            { at: ago(2 * DAY - 2 * HOUR), author: 'Giulia', kind: 'comment', text: 'Ciao Anna, ci indichi l\'indirizzo a cui inoltrare? Restiamo in attesa.' },
            { at: ago(2 * DAY - 2 * HOUR), author: 'Giulia', kind: 'system', text: 'Stato: On Hold (Awaiting Client Action/Update).' }
        ];

        // 6. Ticket In Progress di Luca
        const c6 = [message('conv-6', { name: 'Info Beta', email: 'info@agenziabeta.it' },
            'Accesso VPN da casa',
            'Buongiorno, da ieri la VPN non si collega dal portatile di casa. Errore di certificato.', ago(3 * DAY))];
        tag(c6, 'Luca');
        const r6 = request(c6, { owner: 'Luca', status: 'ticket' });
        r6.ticket = {
            number: 'INC0012090', sysId: 'sim-12090', url: `${SERVICENOW_URL}sim-12090`, state: '2',
            companyId: 'beta', companyName: 'Agenzia Beta', activityId: '', activityName: 'VPN',
            shortDescription: 'Accesso VPN da casa', timeWorkedMinutes: 45,
            holdReason: '', holdExpiration: '', openedBy: 'Luca', openedAt: ago(3 * DAY - HOUR)
        };
        tickets[r6.id] = [
            { at: ago(3 * DAY - HOUR), author: 'Luca', kind: 'system', text: 'Ticket aperto dall\'add-in (Agenzia Beta).' },
            { at: ago(3 * DAY - HOUR), author: 'Info Beta', kind: 'comment', text: c6[0].body }
        ];

        // 7. Ticket On Hold di Marco, scade tra due giorni
        const c7 = [message('conv-7', { name: 'Contabilità Beta', email: 'contabilita@agenziabeta.it' },
            'Firma digitale scaduta',
            'Salve, la firma digitale del responsabile è scaduta. Come possiamo rinnovarla?', ago(30 * HOUR))];
        tag(c7, 'Marco');
        const r7 = request(c7, { owner: 'Marco', status: 'ticket' });
        r7.ticket = {
            number: 'INC0012099', sysId: 'sim-12099', url: `${SERVICENOW_URL}sim-12099`, state: '3',
            companyId: 'beta', companyName: 'Agenzia Beta', activityId: 'password', activityName: 'Reset password',
            shortDescription: 'Firma digitale scaduta', timeWorkedMinutes: 10,
            holdReason: '8', holdExpiration: snDate(new Date(Date.now() + 2 * DAY)), openedBy: 'Marco', openedAt: ago(29 * HOUR)
        };
        tickets[r7.id] = [
            { at: ago(29 * HOUR), author: 'Marco', kind: 'system', text: 'Ticket aperto dall\'add-in (Agenzia Beta · Reset password).' },
            { at: ago(29 * HOUR), author: 'Contabilità Beta', kind: 'comment', text: c7[0].body },
            { at: ago(28 * HOUR), author: 'Marco', kind: 'comment', text: 'Vi abbiamo inviato la procedura di rinnovo: fateci sapere quando avete concluso.' },
            { at: ago(28 * HOUR), author: 'Marco', kind: 'system', text: 'Stato: On Hold (Awaiting Client Resolution).' }
        ];

        return { messages, requests, tickets, ticketCounter: 12101 };
    }

    // ------------------------------------------------------------------
    // Stato (sopravvive al ricaricamento della pagina nella stessa scheda)
    // ------------------------------------------------------------------
    let db = load();
    function load() {
        try {
            const stored = JSON.parse(sessionStorage.getItem(STORAGE_KEY) || 'null');
            if (stored?.requests) return stored;
        } catch (error) { /* si riparte dai dati di esempio */ }
        return seed();
    }
    function persist() {
        try { sessionStorage.setItem(STORAGE_KEY, JSON.stringify(db)); } catch (error) { /* solo anteprima */ }
        listeners.forEach(listener => { try { listener(); } catch (error) { console.error(error); } });
    }
    const copy = value => (value == null ? value : JSON.parse(JSON.stringify(value)));
    const reply = value => new Promise(resolve => setTimeout(() => resolve(copy(value)), DELAY_MS));
    function findRequest(id) {
        const request = db.requests.find(item => item.id === id);
        if (!request) throw new Error('Richiesta non trovata.');
        return request;
    }
    function addHistory(request, author, kind, text) {
        (db.tickets[request.id] ||= []).push({ at: new Date().toISOString(), author, kind, text });
    }
    function setConversationCategory(conversationId, owner) {
        db.messages.filter(item => item.conversationId === conversationId)
            .forEach(item => { item.categories = owner ? [owner] : []; });
    }
    const STATE_NAMES = { '2': 'In Progress', '3': 'On Hold', '6': 'Resolved', '8': 'Canceled' };

    // ------------------------------------------------------------------
    // Servizi usati dal pannello (stessi nomi del contratto in amp.js)
    // ------------------------------------------------------------------
    const AmpBackend = {
        mode: 'simulazione',

        listRequests() {
            return reply(db.requests);
        },
        getRequestByConversation(conversationId) {
            return reply(db.requests.find(item => item.conversationId === conversationId) || null);
        },
        registerRequest(context) {
            let request = db.requests.find(item => item.conversationId === context.conversationId);
            if (!request) {
                request = {
                    id: `req-${context.conversationId}`, conversationId: context.conversationId,
                    messageId: context.itemId, firstMessageId: context.itemId, subject: context.subject,
                    sender: context.sender, receivedAt: context.receivedAt, lastMessageAt: context.receivedAt,
                    preview: String(context.body || '').replace(/\s+/g, ' ').slice(0, 140),
                    owner: '', status: 'queue', ticket: null, pendingAction: null
                };
                db.requests.push(request);
                persist();
            }
            return reply(request);
        },
        takeCharge(requestId, owner) {
            const request = findRequest(requestId);
            request.owner = owner || '';
            setConversationCategory(request.conversationId, request.owner);
            if (request.ticket) addHistory(request, owner || 'Sistema', 'system', owner ? `Preso in carico da ${owner}.` : 'Rilasciato.');
            persist();
            return reply(request);
        },
        discardRequest(requestId, by) {
            const request = findRequest(requestId);
            Object.assign(request, { status: 'discarded', discardedBy: by, discardedAt: new Date().toISOString() });
            persist();
            return reply(request);
        },
        restoreRequest(requestId) {
            const request = findRequest(requestId);
            Object.assign(request, { status: 'queue', discardedBy: '', discardedAt: '' });
            persist();
            return reply(request);
        },
        createTicket(requestId, payload) {
            const request = findRequest(requestId);
            if (request.ticket) throw new Error('Questa richiesta ha già un ticket.');
            db.ticketCounter += 1;
            const number = `INC00${db.ticketCounter}`;
            const sysId = `sim-${db.ticketCounter}`;
            request.ticket = {
                number, sysId, url: `${SERVICENOW_URL}${sysId}`, state: '2',
                companyId: payload.companyId, companyName: payload.companyName,
                activityId: payload.activityId, activityName: payload.activityName,
                shortDescription: payload.fields.shortDescription, timeWorkedMinutes: Number(payload.fields.timeWorkedMinutes) || 0,
                holdReason: '', holdExpiration: '', openedBy: payload.openedBy, openedAt: new Date().toISOString(),
                fields: payload.fields
            };
            Object.assign(request, { status: 'ticket', owner: payload.openedBy });
            setConversationCategory(request.conversationId, payload.openedBy);
            addHistory(request, payload.openedBy, 'system', `Ticket aperto dall'add-in (${payload.companyName} · ${payload.activityName}).`);
            addHistory(request, request.sender?.name || 'Cliente', 'comment', payload.fields.description || '');
            persist();
            return reply(request);
        },
        addTime(requestId, deltaMinutes, by) {
            const request = findRequest(requestId);
            if (!request.ticket) throw new Error('Nessun ticket collegato.');
            const before = request.ticket.timeWorkedMinutes;
            request.ticket.timeWorkedMinutes = Math.max(0, before + Number(deltaMinutes));
            const applied = request.ticket.timeWorkedMinutes - before;
            addHistory(request, by, 'system', `Tempo lavorato: ${applied >= 0 ? '+' : ''}${applied} minuti (totale ${request.ticket.timeWorkedMinutes}).`);
            persist();
            return reply(request);
        },
        getTicketActivity(requestId, { offset = 0, limit = 5 } = {}) {
            const all = (db.tickets[requestId] || []).slice().reverse(); // dal più recente
            return reply({ entries: all.slice(offset, offset + limit), total: all.length });
        },
        setPendingAction(requestId, action) {
            const request = findRequest(requestId);
            request.pendingAction = action ? { ...action, at: new Date().toISOString() } : null;
            persist();
            return reply(request);
        },
        subscribe(listener) {
            listeners.add(listener);
            return () => listeners.delete(listener);
        },

        // --------------------------------------------------------------
        // Solo anteprima: la finta casella di Outlook
        // --------------------------------------------------------------
        mailbox: {
            address: SHARED_MAILBOX,
            listMessages() {
                return copy(db.messages);
            },
            getMessage(id) {
                return copy(db.messages.find(item => item.id === id) || null);
            },
            // Invio della risposta: qui avviene ciò che nell'add-in vero farà la Function App
            // quando vede la mail inviata (stessa conversazione) → aggiorna ServiceNow.
            sendReply(messageId, body, operator) {
                const original = db.messages.find(item => item.id === messageId);
                if (!original) throw new Error('Mail non trovata.');
                const sent = {
                    id: `msg-${db.messages.length + 1}-${Date.now().toString(36)}`,
                    conversationId: original.conversationId,
                    from: { name: operator, email: SHARED_MAILBOX.email },
                    to: [original.direction === 'in' ? original.from : original.to[0]],
                    subject: original.subject.startsWith('RE:') ? original.subject : `RE: ${original.subject}`,
                    body,
                    receivedAt: new Date().toISOString(),
                    categories: original.categories.slice(),
                    direction: 'out'
                };
                db.messages.push(sent);
                const request = db.requests.find(item => item.conversationId === original.conversationId);
                const applied = request ? applyReplyToTicket(request, body, operator) : null;
                if (request) Object.assign(request, { messageId: sent.id, lastMessageAt: sent.receivedAt });
                persist();
                return { sent: copy(sent), applied };
            },
            reset() {
                db = seed();
                persist();
            }
        }
    };

    // Cosa succede su ServiceNow quando parte la risposta (logica "lato server").
    function applyReplyToTicket(request, body, operator) {
        const ticket = request.ticket;
        if (!ticket || ['6', '8'].includes(ticket.state)) return null;
        const action = request.pendingAction || { type: 'comment' };
        addHistory(request, operator, 'comment', body);
        if (action.workNotes) addHistory(request, operator, 'work_note', action.workNotes);
        if (action.type === 'resolve' || action.type === 'cancel') {
            ticket.state = action.type === 'resolve' ? '6' : '8';
            Object.assign(ticket, { resolutionCode: action.resolutionCode || '', resolutionNotes: body, holdReason: '', holdExpiration: '' });
            addHistory(request, operator, 'system', `Stato: ${STATE_NAMES[ticket.state]}${action.resolutionCode ? ` (${action.resolutionCode})` : ''}. Note di risoluzione = testo della mail.`);
        } else if (action.type === 'hold') {
            Object.assign(ticket, { state: '3', holdReason: action.holdReason, holdExpiration: action.holdExpiration });
            addHistory(request, operator, 'system', `Stato: On Hold fino al ${action.holdExpiration}.`);
        } else if (ticket.state === '3') {
            // Una risposta su un ticket in attesa lo riporta In Progress.
            Object.assign(ticket, { state: '2', holdReason: '', holdExpiration: '' });
            addHistory(request, operator, 'system', 'Stato: In Progress.');
        }
        request.pendingAction = null;
        return { type: action.type, state: ticket.state };
    }

    window.AmpBackend = AmpBackend;
})();
