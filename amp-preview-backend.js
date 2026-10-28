/*
 * Servizi SIMULATI per l'anteprima dell'area AMP (amp-preview.html).
 * Non fa parte dell'add-in: taskpane.html non carica questo file.
 *
 * Simula in memoria quello che nell'add-in vero starà dietro la Function App:
 *   - ServiceNow: gli Incident (UNICA fonte per stato, assegnatario, tempo, scadenza On Hold, storico);
 *   - SharePoint: solo i collegamenti conversazione ↔ ticket, le mail scartate e l'azione
 *     da eseguire all'invio della risposta. Si scrive solo DOPO un esito positivo su ServiceNow;
 *   - la casella condivisa di Outlook (solo per la finta Outlook dell'anteprima).
 *
 * Collegamento mail ↔ ticket (triangolazione), nell'ordine:
 *   1. numero INC citato nel thread (l'ultimo citato: lo stesso thread si riusa per richieste nuove);
 *   2. collegamento registrato su SharePoint;
 *   3. somiglianza su ServiceNow: Short description = oggetto del thread e Reported = arrivo
 *      della mail (± 15 minuti). Sono solo proposte: il collegamento lo conferma l'operatore.
 * Se il ticket trovato è chiuso e nel thread è arrivata una mail dopo la chiusura, è una
 * richiesta nuova: il ticket chiuso diventa "ticket precedente".
 *
 * Espone window.AmpBackend con i metodi del contratto descritto in amp.js (blocco 3) e
 * in amp-servizi.md. Ogni metodo restituisce una Promise con copie dei dati.
 */
(function () {
    'use strict';

    const STORAGE_KEY = 'amp_preview_backend_v3';
    const SERVICENOW_URL = 'https://amp.service-now.com/incident.do?sys_id=';
    const DELAY_MS = 220;
    const SIMILAR_MINUTES = 15;
    const INC_PATTERN = /\bINC\d{7}\b/gi;
    const OPEN_STATES = ['2', '3'];
    const CLOSED_STATES = ['6', '7', '8'];
    const STATE_NAMES = { '2': 'In Progress', '3': 'On Hold', '6': 'Resolved', '7': 'Closed', '8': 'Canceled' };
    const listeners = new Set();

    const MINUTE = 60 * 1000;
    const HOUR = 60 * MINUTE;
    const DAY = 24 * HOUR;
    const pad = value => String(value).padStart(2, '0');
    const at = ms => new Date(Date.now() - ms);
    const iso = ms => at(ms).toISOString();
    function snDate(date) {
        return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
    }
    function parseDate(value) {
        if (!value) return null;
        const text = String(value);
        const date = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}/.test(text) ? new Date(text.replace(' ', 'T')) : new Date(text);
        return Number.isNaN(date.getTime()) ? null : date;
    }
    const SHARED_MAILBOX = { name: 'Supporto AMP', email: 'supporto.amp@example.com' };
    const OPERATORS = {
        marco: { name: 'Marco Rossi', sysId: 'sys_user_1' },
        giulia: { name: 'Giulia Verdi', sysId: 'sys_user_2' },
        luca: { name: 'Luca Neri', sysId: 'sys_user_3' }
    };

    // ------------------------------------------------------------------
    // Dati di partenza (tutti inventati)
    // ------------------------------------------------------------------
    function seed() {
        const messages = [];
        const incidents = {};
        const history = {};
        const links = [];
        let counter = 0;
        const mail = (conversationId, from, subject, body, ms, extra = {}) => {
            counter += 1;
            messages.push({
                id: `msg-${counter}`, conversationId, from, to: [SHARED_MAILBOX], subject, body,
                receivedAt: iso(ms), categories: [], direction: 'in', ...extra
            });
        };
        const reply = (conversationId, operator, to, subject, body, ms) => mail(conversationId,
            { name: operator.name, email: SHARED_MAILBOX.email }, `RE: ${subject}`, body, ms, { direction: 'out', to: [to] });
        const incident = (number, data, entries = []) => {
            const sysId = `sim-${number.slice(3)}`;
            incidents[number] = {
                number, sysId, url: `${SERVICENOW_URL}${sysId}`, holdReason: '', holdExpiration: '', closedAt: '',
                resolutionCode: '', ...data
            };
            history[number] = entries;
        };
        const entry = (ms, author, kind, text) => ({ at: iso(ms), author, kind, text });

        const alfa = { name: 'Mario Bianchi', email: 'mario.bianchi@agenziaalfa.it' };
        const beta = { name: 'Ufficio Beta', email: 'ufficio@agenziabeta.it' };
        const gamma = { name: 'Segreteria Comune Gamma', email: 'segreteria@comune-gamma.it' };

        // 1-3. Richieste nuove, nessun ticket
        mail('conv-1', alfa, 'Casella di posta piena',
            'Buongiorno,\nla casella di posta dell\'ufficio risulta piena e non riceviamo più email dai clienti.\nPotete aiutarci?\n\nGrazie\nMario Bianchi\nAgenzia Alfa', 25 * MINUTE);
        mail('conv-2', beta, 'Reset password utente Verdi',
            'Salve,\nil collega Luigi Verdi non riesce più ad accedere al gestionale: chiede il reset della password.\n\nCordiali saluti\nAgenzia Beta', 2 * HOUR);
        mail('conv-3', { name: 'Fornitore Servizi', email: 'newsletter@fornitore.example' }, 'Aggiornamento condizioni di servizio',
            'Gentile cliente,\nla informiamo che dal prossimo mese entreranno in vigore le nuove condizioni di servizio.\nNessuna azione è richiesta.', 3 * HOUR);

        // 4. Ticket di Marco, collegato dal numero INC scritto nella presa in carico
        const s4 = 'Stampante secondo piano non stampa';
        mail('conv-4', gamma, s4, 'Buongiorno,\nla stampante del secondo piano non stampa più da stamattina. Sul display compare "errore 50.4".\n\nSegreteria', 26 * HOUR);
        reply('conv-4', OPERATORS.marco, gamma, s4, 'Rif. INC0012101\n\nBuongiorno,\nabbiamo preso in carico la segnalazione. Potete spegnere e riaccendere la stampante e dirci se l\'errore resta?\n\nSaluti,\nSupporto AMP', 25 * HOUR);
        mail('conv-4', gamma, `RE: ${s4}`, 'L\'abbiamo riavviata ma l\'errore resta uguale.', 5 * HOUR);
        incident('INC0012101', {
            state: '2', assignedTo: OPERATORS.marco, companyId: 'gamma', companyName: 'Comune Gamma', activityId: 'stampante', activityName: 'Stampante',
            shortDescription: s4, reportedOn: snDate(at(26 * HOUR)), openedAt: iso(25.5 * HOUR), timeWorkedMinutes: 30
        }, [
            entry(25.5 * HOUR, 'Marco Rossi', 'system', 'Ticket aperto dall\'add-in (Comune Gamma · Stampante).'),
            entry(25.5 * HOUR, gamma.name, 'comment', messages[3].body),
            entry(25 * HOUR, 'Marco Rossi', 'comment', messages[4].body),
            entry(24 * HOUR, 'Marco Rossi', 'work_note', 'Errore 50.4: probabile problema di alimentazione. Verificare con il fornitore se persiste.'),
            entry(20 * HOUR, 'Marco Rossi', 'system', 'Tempo lavorato: +15 minuti (totale 30).'),
            entry(5 * HOUR, gamma.name, 'comment', messages[5].body),
            entry(4.5 * HOUR, 'Marco Rossi', 'work_note', 'Chiesto al fornitore un intervento in sede.')
        ]);

        // 5. Ticket di Giulia, On Hold scaduto, collegato solo su SharePoint
        const neri = { name: 'Anna Neri', email: 'a.neri@agenziaalfa.it' };
        mail('conv-5', neri, 'Inoltro automatico non funziona', 'Ciao,\nl\'inoltro automatico delle mail verso il cellulare non funziona più.\nAnna', 2 * DAY);
        incident('INC0012095', {
            state: '3', assignedTo: OPERATORS.giulia, companyId: 'alfa', companyName: 'Agenzia Alfa', activityId: 'posta', activityName: 'Posta elettronica',
            shortDescription: '[FWD] Inoltro automatico non funziona', reportedOn: snDate(at(2 * DAY)), openedAt: iso(2 * DAY - HOUR), timeWorkedMinutes: 15,
            holdReason: '7', holdExpiration: snDate(at(3 * HOUR))
        }, [
            entry(2 * DAY - HOUR, 'Giulia Verdi', 'system', 'Ticket aperto dall\'add-in (Agenzia Alfa · Posta elettronica).'),
            entry(2 * DAY - HOUR, neri.name, 'comment', 'Ciao,\nl\'inoltro automatico delle mail verso il cellulare non funziona più.\nAnna'),
            entry(2 * DAY - 2 * HOUR, 'Giulia Verdi', 'system', 'Stato: On Hold (Awaiting Client Action/Update).')
        ]);
        links.push({ conversationId: 'conv-5', number: 'INC0012095', linkedAt: iso(2 * DAY - HOUR), linkedBy: 'Giulia Verdi' });

        // 6. Ticket di Luca, collegato su SharePoint
        const info = { name: 'Info Beta', email: 'info@agenziabeta.it' };
        mail('conv-6', info, 'Accesso VPN da casa', 'Buongiorno, da ieri la VPN non si collega dal portatile di casa. Errore di certificato.', 3 * DAY);
        incident('INC0012090', {
            state: '2', assignedTo: OPERATORS.luca, companyId: 'beta', companyName: 'Agenzia Beta', activityId: '', activityName: 'VPN',
            shortDescription: 'Accesso VPN da casa', reportedOn: snDate(at(3 * DAY)), openedAt: iso(3 * DAY - HOUR), timeWorkedMinutes: 45
        }, [entry(3 * DAY - HOUR, 'Luca Neri', 'system', 'Ticket aperto dall\'add-in (Agenzia Beta).')]);
        links.push({ conversationId: 'conv-6', number: 'INC0012090', linkedAt: iso(3 * DAY - HOUR), linkedBy: 'Luca Neri' });

        // 7. Ticket di Marco, On Hold fino a dopodomani, INC nella risposta
        const contab = { name: 'Contabilità Beta', email: 'contabilita@agenziabeta.it' };
        const s7 = 'Firma digitale scaduta';
        mail('conv-7', contab, s7, 'Salve, la firma digitale del responsabile è scaduta. Come possiamo rinnovarla?', 30 * HOUR);
        reply('conv-7', OPERATORS.marco, contab, s7, 'Rif. INC0012099\n\nVi abbiamo inviato la procedura di rinnovo: fateci sapere quando avete concluso.\n\nSaluti,\nSupporto AMP', 28 * HOUR);
        incident('INC0012099', {
            state: '3', assignedTo: OPERATORS.marco, companyId: 'beta', companyName: 'Agenzia Beta', activityId: 'password', activityName: 'Reset password',
            shortDescription: s7, reportedOn: snDate(at(30 * HOUR)), openedAt: iso(29 * HOUR), timeWorkedMinutes: 10,
            holdReason: '8', holdExpiration: snDate(new Date(Date.now() + 2 * DAY))
        }, [
            entry(29 * HOUR, 'Marco Rossi', 'system', 'Ticket aperto dall\'add-in (Agenzia Beta · Reset password).'),
            entry(28 * HOUR, 'Marco Rossi', 'comment', 'Vi abbiamo inviato la procedura di rinnovo: fateci sapere quando avete concluso.'),
            entry(28 * HOUR, 'Marco Rossi', 'system', 'Stato: On Hold (Awaiting Client Resolution).')
        ]);

        // 8. Thread riusato: il ticket di 10 giorni fa è chiuso, oggi è arrivata una richiesta nuova
        const s8 = 'Richiesta nuova utenza gestionale';
        mail('conv-8', beta, s8, 'Buongiorno, ci serve un\'utenza per la nuova collega Sara Bassi.\nGrazie', 10 * DAY);
        reply('conv-8', OPERATORS.marco, beta, s8, 'Rif. INC0012050\n\nBuongiorno, l\'utenza è stata creata: le credenziali arrivano a parte.\n\nSaluti,\nSupporto AMP', 9 * DAY);
        mail('conv-8', beta, `RE: ${s8}`, 'Buongiorno, ci servirebbe un\'altra utenza, per il collega Paolo Riva.\nGrazie mille', 40 * MINUTE);
        incident('INC0012050', {
            state: '6', assignedTo: OPERATORS.marco, companyId: 'beta', companyName: 'Agenzia Beta', activityId: 'password', activityName: 'Reset password',
            shortDescription: s8, reportedOn: snDate(at(10 * DAY)), openedAt: iso(10 * DAY - HOUR), closedAt: iso(9 * DAY), timeWorkedMinutes: 20,
            resolutionCode: 'Solved (Permanently)'
        }, [entry(9 * DAY, 'Marco Rossi', 'system', 'Stato: Resolved (Solved (Permanently)).')]);

        // 9. Ticket aperto su ServiceNow senza passare dall'add-in: lo si trova per somiglianza
        const s9 = 'Errore stampa PDF gestionale';
        mail('conv-9', alfa, s9, 'Buongiorno, dal gestionale non riusciamo a stampare in PDF: compare "modulo non trovato".', 20 * HOUR);
        incident('INC0012097', {
            state: '2', assignedTo: OPERATORS.giulia, companyId: 'alfa', companyName: 'Agenzia Alfa', activityId: '', activityName: 'Gestionale',
            shortDescription: s9, reportedOn: snDate(at(20 * HOUR - 2 * MINUTE)), openedAt: iso(19 * HOUR), timeWorkedMinutes: 10
        }, [entry(19 * HOUR, 'Giulia Verdi', 'system', 'Ticket aperto da ServiceNow.')]);

        // 10. Ticket di Marco senza mail (aperto direttamente su AMP)
        incident('INC0012098', {
            state: '2', assignedTo: OPERATORS.marco, companyId: 'gamma', companyName: 'Comune Gamma', activityId: '', activityName: 'Manutenzione',
            shortDescription: 'Aggiornamento antivirus postazioni', reportedOn: snDate(at(2 * DAY)), openedAt: iso(2 * DAY), timeWorkedMinutes: 60
        }, [entry(2 * DAY, 'Marco Rossi', 'system', 'Ticket aperto da ServiceNow.')]);

        return { messages, incidents, history, links, discarded: {}, pending: {}, ticketCounter: 12101 };
    }

    // ------------------------------------------------------------------
    // Stato (sopravvive al ricaricamento della pagina nella stessa scheda)
    // ------------------------------------------------------------------
    let db = load();
    function load() {
        try {
            const stored = JSON.parse(sessionStorage.getItem(STORAGE_KEY) || 'null');
            if (stored?.incidents) return stored;
        } catch (error) { /* si riparte dai dati di esempio */ }
        return seed();
    }
    function persist() {
        try { sessionStorage.setItem(STORAGE_KEY, JSON.stringify(db)); } catch (error) { /* solo anteprima */ }
        listeners.forEach(listener => { try { listener(); } catch (error) { console.error(error); } });
    }
    const copy = value => (value == null ? value : JSON.parse(JSON.stringify(value)));
    const respond = value => new Promise(resolve => setTimeout(() => resolve(copy(value)), DELAY_MS));
    const fail = message => new Promise((resolve, reject) => setTimeout(() => reject(new Error(message)), DELAY_MS));

    // ------------------------------------------------------------------
    // Triangolazione
    // ------------------------------------------------------------------
    const conversationMessages = conversationId => db.messages
        .filter(item => item.conversationId === conversationId)
        .sort((a, b) => a.receivedAt.localeCompare(b.receivedAt));
    function normalizeSubject(subject) {
        return String(subject || '').toLowerCase()
            .replace(/^\s*((re|r|i|fw|fwd|inoltro)\s*:\s*)+/i, '')
            .replace(/\[[^\]]*\]/g, ' ')
            .replace(/\s+/g, ' ').trim();
    }
    // Numeri INC citati nel thread, dal più recente.
    function incNumbersIn(messages, extra = []) {
        return [...new Set([...incMentions(messages).map(item => item.number), ...extra.map(number => String(number).toUpperCase())])];
    }
    // Citazioni INC con la data della mail che le contiene, dalla più recente.
    function incMentions(messages) {
        const mentions = [];
        [...messages].reverse().forEach(message => {
            `${message.subject}\n${message.body}`.match(INC_PATTERN)?.forEach(number => mentions.push({ number: number.toUpperCase(), when: message.receivedAt }));
        });
        return mentions;
    }
    function similarIncidents(subject, times, excluded = []) {
        const wanted = normalizeSubject(subject);
        if (!wanted) return [];
        return Object.values(db.incidents).filter(incident => {
            if (excluded.includes(incident.number)) return false;
            const description = normalizeSubject(incident.shortDescription);
            if (!description || !(description.includes(wanted) || wanted.includes(description))) return false;
            const reported = parseDate(incident.reportedOn);
            return reported && times.some(time => Math.abs(reported - time) <= SIMILAR_MINUTES * MINUTE);
        });
    }
    // Risultato per una conversazione: { ticket, via, previous, candidates }.
    function resolveConversation(conversationId, extraNumbers = []) {
        const messages = conversationMessages(conversationId);
        const incoming = messages.filter(item => item.direction === 'in');
        const lastIncoming = incoming[incoming.length - 1] || null;
        // Vale l'indizio più recente: un INC citato in una mail o un collegamento registrato.
        const evidence = [
            ...incMentions(messages).map(item => ({ ...item, via: 'inc' })),
            // Numeri letti dalla mail aperta: solo come ultima riserva.
            ...extraNumbers.map(number => ({ number: String(number).toUpperCase(), when: '', via: 'inc' })),
            ...db.links.filter(item => item.conversationId === conversationId)
                .map(item => ({ number: item.number, when: item.linkedAt, via: item.confirmedBy ? 'confirmed' : 'link' }))
        ].filter(item => db.incidents[item.number]).sort((a, b) => b.when.localeCompare(a.when));
        let found = evidence[0] ? db.incidents[evidence[0].number] : null;
        const via = evidence[0]?.via || null;
        let previous = null;
        if (found && CLOSED_STATES.includes(found.state) && lastIncoming && found.closedAt && lastIncoming.receivedAt > found.closedAt) {
            previous = found;
            found = null;
        }
        let candidates = [];
        if (!found && messages.length) {
            const times = (previous ? [lastIncoming] : [incoming[0], lastIncoming]).filter(Boolean).map(item => new Date(item.receivedAt));
            candidates = similarIncidents(messages[0].subject, times, previous ? [previous.number] : []);
        }
        return { ticket: found ? decorate(found, conversationId) : null, via: found ? via : null, previous: previous ? decorate(previous, conversationId) : null, candidates: candidates.map(item => decorate(item)) };
    }
    // Ticket con i dati di SharePoint (mail collegata, azione all'invio).
    // mailUnconfirmed: la mail è stata trovata solo per somiglianza e va confermata dall'operatore.
    function decorate(incident, conversationId) {
        const linked = conversationId || conversationOf(incident.number);
        const similar = linked ? null : similarConversationOf(incident);
        const conversation = linked || similar;
        const messages = conversation ? conversationMessages(conversation) : [];
        return {
            ...incident,
            conversationId: conversation || null,
            messageId: messages.length ? messages[messages.length - 1].id : null,
            mailUnconfirmed: Boolean(similar),
            pendingAction: db.pending[incident.number] || null
        };
    }
    function similarConversationOf(incident) {
        const conversations = [...new Set(db.messages.map(item => item.conversationId))];
        return conversations.find(id => {
            const messages = conversationMessages(id);
            const incoming = messages.filter(item => item.direction === 'in');
            const times = incoming.map(item => new Date(item.receivedAt));
            return similarIncidents(messages[0].subject, times).some(item => item.number === incident.number);
        }) || null;
    }
    function conversationOf(number) {
        const conversations = [...new Set(db.messages.map(item => item.conversationId))];
        return conversations.find(id => incNumbersIn(conversationMessages(id)).includes(number))
            || db.links.filter(item => item.number === number).slice(-1)[0]?.conversationId || null;
    }
    function requireIncident(number) {
        const incident = db.incidents[number];
        if (!incident) throw new Error(`Ticket ${number} non trovato su ServiceNow.`);
        return incident;
    }
    function addHistory(number, author, kind, text) {
        (db.history[number] ||= []).push({ at: new Date().toISOString(), author, kind, text });
    }
    // Riga della coda per una conversazione senza ticket aperto (o null).
    function queueItem(conversationId) {
        const messages = conversationMessages(conversationId);
        const incoming = messages.filter(item => item.direction === 'in');
        const lastIncoming = incoming[incoming.length - 1];
        if (!lastIncoming) return null;
        const discarded = db.discarded[conversationId];
        if (discarded && discarded.at > lastIncoming.receivedAt) return null;
        // In coda solo le mail che non hanno ancora un INC su ServiceNow: niente ticket aperto
        // collegato e nessun ticket simile (quello va confermato dalla vista Mail o da Ticket).
        const result = resolveConversation(conversationId);
        if (result.ticket || result.candidates.length) return null;
        return {
            conversationId,
            messageId: messages[messages.length - 1].id,
            subject: messages[0].subject,
            sender: lastIncoming.from,
            receivedAt: lastIncoming.receivedAt,
            preview: lastIncoming.body.replace(/\s+/g, ' ').slice(0, 160),
            previous: result.previous ? { number: result.previous.number, state: result.previous.state } : null,
            candidates: result.candidates.map(item => item.number)
        };
    }

    // ------------------------------------------------------------------
    // Servizi usati dal pannello (contratto: amp.js blocco 3 e amp-servizi.md)
    // ------------------------------------------------------------------
    const AmpBackend = {
        mode: 'simulazione',

        // Mail della casella senza ticket aperto (Graph + triangolazione).
        listQueue() {
            const conversations = [...new Set(db.messages.map(item => item.conversationId))];
            return respond(conversations.map(queueItem).filter(Boolean));
        },
        // Ticket aperti della coda su ServiceNow (In Progress e On Hold), con la mail se collegata.
        listOpenTickets() {
            return respond(Object.values(db.incidents).filter(item => OPEN_STATES.includes(item.state)).map(item => decorate(item)));
        },
        // Triangolazione per la mail aperta: { ticket, via, previous, candidates, discarded }.
        findTicketForMail(mail) {
            const result = resolveConversation(mail.conversationId, mail.incNumbers || []);
            const lastIncoming = conversationMessages(mail.conversationId).filter(item => item.direction === 'in').slice(-1)[0];
            const discarded = db.discarded[mail.conversationId];
            result.discarded = discarded && (!lastIncoming || discarded.at > lastIncoming.receivedAt) ? discarded : null;
            return respond(result);
        },
        getTicket(number) {
            try { return respond(decorate(requireIncident(number))); } catch (error) { return fail(error.message); }
        },
        // L'operatore conferma che la mail appartiene a un ticket esistente → SharePoint.
        linkTicket(conversationId, number, by) {
            requireIncident(number);
            db.links.push({ conversationId, number, linkedAt: new Date().toISOString(), linkedBy: by?.displayValue || '', confirmedBy: by?.displayValue || '' });
            persist();
            return respond(resolveConversation(conversationId));
        },
        // Apre l'Incident su ServiceNow; solo se riesce registra il collegamento su SharePoint.
        createTicket(mail, payload) {
            db.ticketCounter += 1;
            const number = `INC00${db.ticketCounter}`;
            const assignee = payload.fields.assignedTo;
            db.incidents[number] = {
                number, sysId: `sim-${db.ticketCounter}`, url: `${SERVICENOW_URL}sim-${db.ticketCounter}`, state: '2',
                assignedTo: { name: assignee.displayValue, sysId: assignee.value },
                companyId: payload.companyId, companyName: payload.companyName,
                activityId: payload.activityId, activityName: payload.activityName,
                shortDescription: payload.fields.shortDescription, reportedOn: payload.fields.reportedOn,
                openedAt: new Date().toISOString(), timeWorkedMinutes: Number(payload.fields.timeWorkedMinutes) || 0,
                holdReason: '', holdExpiration: '', closedAt: '', resolutionCode: '', fields: payload.fields
            };
            db.history[number] = [];
            addHistory(number, assignee.displayValue, 'system', `Ticket aperto dall'add-in (${payload.companyName} · ${payload.activityName}).`);
            addHistory(number, mail.sender?.name || 'Cliente', 'comment', payload.fields.description || '');
            db.links.push({ conversationId: mail.conversationId, number, linkedAt: new Date().toISOString(), linkedBy: assignee.displayValue });
            delete db.discarded[mail.conversationId];
            persist();
            return respond(decorate(db.incidents[number], mail.conversationId));
        },
        // Prende in carico un ticket di un collega: cambia Assigned to su ServiceNow.
        assignTicket(number, assignee) {
            try {
                const incident = requireIncident(number);
                incident.assignedTo = { name: assignee.displayValue, sysId: assignee.value };
                addHistory(number, assignee.displayValue, 'system', `Assegnato a ${assignee.displayValue}.`);
                persist();
                return respond(decorate(incident));
            } catch (error) { return fail(error.message); }
        },
        addTime(number, deltaMinutes, by) {
            try {
                const incident = requireIncident(number);
                const before = incident.timeWorkedMinutes;
                incident.timeWorkedMinutes = Math.max(0, before + Number(deltaMinutes));
                const applied = incident.timeWorkedMinutes - before;
                addHistory(number, by, 'system', `Tempo lavorato: ${applied >= 0 ? '+' : ''}${applied} minuti (totale ${incident.timeWorkedMinutes}).`);
                persist();
                return respond(decorate(incident));
            } catch (error) { return fail(error.message); }
        },
        getTicketActivity(number, { offset = 0, limit = 5 } = {}) {
            const all = (db.history[number] || []).slice().reverse();
            return respond({ entries: all.slice(offset, offset + limit), total: all.length });
        },
        // Azione da eseguire all'invio della risposta (SharePoint).
        setPendingAction(number, action) {
            try {
                const incident = requireIncident(number);
                if (action) db.pending[number] = { ...action, at: new Date().toISOString() };
                else delete db.pending[number];
                persist();
                return respond(decorate(incident));
            } catch (error) { return fail(error.message); }
        },
        // Mail che non richiede un ticket (SharePoint). Torna in coda se arriva una mail nuova nel thread.
        discardMail(conversationId, by) {
            db.discarded[conversationId] = { by: by?.displayValue || '', at: new Date().toISOString() };
            persist();
            return respond(db.discarded[conversationId]);
        },
        restoreMail(conversationId) {
            delete db.discarded[conversationId];
            persist();
            return respond(true);
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
            operators: OPERATORS,
            listMessages() {
                return copy(db.messages);
            },
            getMessage(id) {
                return copy(db.messages.find(item => item.id === id) || null);
            },
            // Invio della risposta: qui avviene ciò che nell'add-in vero farà la Function App
            // quando vede la mail inviata nella conversazione → aggiorna ServiceNow.
            sendReply(messageId, body, operatorName) {
                const original = db.messages.find(item => item.id === messageId);
                if (!original) throw new Error('Mail non trovata.');
                const thread = conversationMessages(original.conversationId);
                const firstIncoming = thread.find(item => item.direction === 'in') || original;
                db.messages.push({
                    id: `msg-${db.messages.length + 1}-${Date.now().toString(36)}`,
                    conversationId: original.conversationId,
                    from: { name: operatorName, email: SHARED_MAILBOX.email },
                    to: [firstIncoming.from],
                    subject: thread[0].subject.startsWith('RE:') ? thread[0].subject : `RE: ${thread[0].subject}`,
                    body, receivedAt: new Date().toISOString(), categories: [], direction: 'out'
                });
                const { ticket } = resolveConversation(original.conversationId);
                const applied = ticket && OPEN_STATES.includes(ticket.state) ? applyReplyToTicket(db.incidents[ticket.number], body, operatorName) : null;
                persist();
                return { applied };
            },
            reset() {
                db = seed();
                persist();
            }
        }
    };

    // Cosa succede su ServiceNow quando parte la risposta (logica "lato server").
    function applyReplyToTicket(incident, body, operator) {
        const number = incident.number;
        const action = db.pending[number] || { type: 'comment' };
        addHistory(number, operator, 'comment', body);
        if (action.workNotes) addHistory(number, operator, 'work_note', action.workNotes);
        if (action.type === 'resolve' || action.type === 'cancel') {
            Object.assign(incident, {
                state: action.type === 'resolve' ? '6' : '8', resolutionCode: action.resolutionCode || '',
                resolutionNotes: body, holdReason: '', holdExpiration: '', closedAt: new Date().toISOString()
            });
            addHistory(number, operator, 'system', `Stato: ${STATE_NAMES[incident.state]}${action.resolutionCode ? ` (${action.resolutionCode})` : ''}. Resolution notes = testo della mail.`);
        } else if (action.type === 'hold') {
            Object.assign(incident, { state: '3', holdReason: action.holdReason, holdExpiration: action.holdExpiration });
            addHistory(number, operator, 'system', `Stato: On Hold fino al ${action.holdExpiration}.`);
        } else if (incident.state === '3') {
            Object.assign(incident, { state: '2', holdReason: '', holdExpiration: '' });
            addHistory(number, operator, 'system', 'Stato: In Progress.');
        }
        delete db.pending[number];
        return { type: action.type, state: incident.state, number };
    }

    window.AmpBackend = AmpBackend;
})();
