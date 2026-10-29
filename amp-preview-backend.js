/*
 * Servizi SIMULATI per l'anteprima dell'area AMP (amp-preview.html).
 * Non fa parte dell'add-in: taskpane.html non carica questo file.
 *
 * Simula in memoria quello che nell'add-in vero starà dietro la Function App:
 *   - ServiceNow: gli Incident e il loro storico. Resta il riferimento: ogni modifica va prima qui;
 *   - SharePoint: la copia che il pannello legge, così non serve interrogare ServiceNow:
 *       Thread AMP     una riga per conversazione (libero / in gestione / scartato, ticket);
 *       Ticket AMP     copia dei valori dei ticket (stato, assegnatario, tempo, scadenze…);
 *       Movimenti AMP  una riga per ogni azione (apertura, tempo, risposta, allineamento…);
 *   - un allineamento periodico (sync) che riporta su SharePoint le modifiche fatte su AMP
 *     fuori dal pannello;
 *   - la casella condivisa di Outlook (solo per la finta Outlook dell'anteprima).
 *
 * In coda: un thread per conversationId, se non è in gestione. Quando il ticket si chiude il
 * thread torna libero subito, ma ricompare In coda solo se arriva un messaggio nuovo.
 *
 * Collegamento mail ↔ ticket (triangolazione): vale l'indizio più recente tra il numero INC
 * citato nel thread e il ticket registrato sul thread; altrimenti ticket simili (stesso oggetto,
 * Reported vicino all'arrivo) da confermare. Ticket chiuso + mail successiva = richiesta nuova.
 *
 * Espone window.AmpBackend con i metodi del contratto descritto in amp.js (blocco 3) e
 * in amp-servizi.md. Ogni metodo restituisce una Promise con copie dei dati.
 */
(function () {
    'use strict';

    const STORAGE_KEY = 'amp_preview_backend_v5';
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

        // 11. Una richiesta, due ticket aperti subito (creazione DL e aggiunta membri); poi una risposta nuova
        const acquisti = { name: 'Ufficio Acquisti Beta', email: 'acquisti@agenziabeta.it' };
        const s11 = 'Nuova lista di distribuzione Acquisti';
        mail('conv-10', acquisti, s11, 'Buongiorno,\nci serve una lista di distribuzione acquisti@agenziabeta.it con dentro Rossi, Bianchi e Verdi.\nGrazie', 6 * HOUR);
        reply('conv-10', OPERATORS.marco, acquisti, s11, 'Rif. INC0012091, INC0012092\n\nBuongiorno, abbiamo preso in carico la creazione della lista e l\'aggiunta dei membri.\n\nSaluti,\nSupporto AMP', 5.5 * HOUR);
        mail('conv-10', acquisti, `RE: ${s11}`, 'Potete aggiungere anche la collega Neri? Grazie', 30 * MINUTE);
        incident('INC0012091', {
            state: '2', assignedTo: OPERATORS.marco, companyId: 'beta', companyName: 'Agenzia Beta', activityId: 'posta', activityName: 'Posta elettronica',
            shortDescription: '[FWD] Creazione DL acquisti@agenziabeta.it', reportedOn: snDate(at(6 * HOUR)), openedAt: iso(5.6 * HOUR), timeWorkedMinutes: 15
        }, [entry(5.6 * HOUR, 'Marco Rossi', 'system', 'Ticket aperto dall\'add-in (Agenzia Beta · Posta elettronica).')]);
        incident('INC0012092', {
            state: '2', assignedTo: OPERATORS.marco, companyId: 'beta', companyName: 'Agenzia Beta', activityId: 'posta', activityName: 'Posta elettronica',
            shortDescription: '[FWD] Aggiunta membri DL acquisti@agenziabeta.it', reportedOn: snDate(at(6 * HOUR)), openedAt: iso(5.6 * HOUR), timeWorkedMinutes: 15
        }, [entry(5.6 * HOUR, 'Marco Rossi', 'system', 'Ticket aperto dall\'add-in (Agenzia Beta · Posta elettronica).')]);

        return { messages, incidents, history, links, discarded: {}, pending: {}, ticketCounter: 12101 };
    }

    // ------------------------------------------------------------------
    // Stato (sopravvive al ricaricamento della pagina nella stessa scheda)
    //   db.incidents, db.history  → ServiceNow (riferimento)
    //   db.sp.threads             → SharePoint "Thread AMP": una riga per conversazione
    //   db.sp.tickets             → SharePoint "Ticket AMP": copia dei ticket letta dal pannello
    //   db.sp.movements           → SharePoint "Movimenti AMP": una riga per ogni azione
    // ------------------------------------------------------------------
    let db = null; // caricato più sotto, dopo le funzioni che lo usano
    function load() {
        try {
            const stored = JSON.parse(sessionStorage.getItem(STORAGE_KEY) || 'null');
            if (stored?.sp) return stored;
        } catch (error) { /* si riparte dai dati di esempio */ }
        return createDatabase();
    }
    function createDatabase() {
        const base = seed();
        const data = {
            messages: base.messages, incidents: base.incidents, history: base.history, ticketCounter: base.ticketCounter,
            sp: { threads: {}, tickets: {}, movements: [], syncedAt: '' }
        };
        db = data;
        // Situazione di partenza su SharePoint: copia di tutti i ticket e thread già collegati.
        Object.keys(data.incidents).forEach(number => mirror(number));
        base.links.forEach(link => registerThread(link.conversationId, link.number, link.linkedBy, link.linkedAt));
        conversationIds().forEach(id => {
            const numbers = [...new Set(incMentions(conversationMessages(id)).map(item => item.number))].filter(item => data.incidents[item]);
            numbers.reverse().forEach(number => registerThread(id, number, data.incidents[number].assignedTo.name, data.incidents[number].openedAt));
        });
        // Chi ha in carico il thread l'ha "visto" quando l'ha preso (le mail arrivate dopo sono nuove).
        Object.values(data.sp.threads).forEach(thread => { thread.ownerSeenAt = thread.links[0]?.linkedAt || ''; });
        data.sp.syncedAt = new Date().toISOString();
        return data;
    }
    function persist() {
        try { sessionStorage.setItem(STORAGE_KEY, JSON.stringify(db)); } catch (error) { /* solo anteprima */ }
        listeners.forEach(listener => { try { listener(); } catch (error) { console.error(error); } });
    }
    const copy = value => (value == null ? value : JSON.parse(JSON.stringify(value)));
    const respond = value => new Promise(resolve => setTimeout(() => resolve(copy(value)), DELAY_MS));
    const fail = message => new Promise((resolve, reject) => setTimeout(() => reject(new Error(message)), DELAY_MS));
    const now = () => new Date().toISOString();

    // ------------------------------------------------------------------
    // SharePoint: copia dei ticket, thread, movimenti
    // ------------------------------------------------------------------
    const SNOW_FIELDS = ['number', 'sysId', 'url', 'state', 'assignedTo', 'companyId', 'companyName', 'activityId', 'activityName',
        'shortDescription', 'reportedOn', 'openedAt', 'closedAt', 'timeWorkedMinutes', 'holdReason', 'holdExpiration', 'resolutionCode'];
    const snowView = incident => Object.fromEntries(SNOW_FIELDS.map(key => [key, incident[key] ?? '']));
    // Copia su SharePoint i valori del ticket letti da ServiceNow (l'azione all'invio resta).
    function mirror(number) {
        const previous = db.sp.tickets[number] || {};
        db.sp.tickets[number] = { ...snowView(db.incidents[number]), pendingAction: previous.pendingAction || null, updatedAt: now() };
        return db.sp.tickets[number];
    }
    function record(type, number, by, details = '', conversationId = null) {
        db.sp.movements.push({ at: now(), type, number: number || '', by: by || '', details, conversationId });
    }
    // Il thread registra il ticket indicato (un thread può avere più ticket, anche aperti insieme).
    // Chi lo prende in carico per primo ne diventa il proprietario finché resta in gestione.
    function registerThread(conversationId, number, by, when = now(), confirmed = false) {
        const thread = db.sp.threads[conversationId] || { conversationId, links: [], owner: '', ownerSeenAt: '' };
        if (!thread.links.some(link => link.number === number)) thread.links.push({ number, linkedAt: when, linkedBy: by || '', confirmedBy: confirmed ? by : '' });
        if (thread.status !== 'managed' || !thread.owner) thread.owner = by || '';
        Object.assign(thread, { discardedBy: '', discardedAt: '' });
        db.sp.threads[conversationId] = thread;
        refreshThreadStatus(thread);
        return thread;
    }
    const threadNumbers = thread => (thread?.links || []).map(link => link.number);
    // In gestione finché almeno un ticket del thread è aperto; poi libero (subito).
    function refreshThreadStatus(thread) {
        const tickets = threadNumbers(thread).map(number => db.sp.tickets[number]).filter(Boolean);
        if (!tickets.length) return;
        const open = tickets.some(ticket => OPEN_STATES.includes(ticket.state));
        if (open) Object.assign(thread, { status: 'managed', freedAt: '' });
        else if (thread.status !== 'free' && thread.status !== 'discarded') {
            const closedAt = tickets.map(ticket => ticket.closedAt).filter(Boolean).sort().slice(-1)[0];
            Object.assign(thread, { status: 'free', freedAt: closedAt || now() });
        }
        thread.updatedAt = now();
    }
    function updateThreadForTicket(number) {
        Object.values(db.sp.threads).filter(thread => threadNumbers(thread).includes(number)).forEach(refreshThreadStatus);
    }
    // Messaggi arrivati dopo l'ultima risposta o l'ultima visione di chi ha il thread in carico.
    function newMessagesOf(conversationId) {
        const thread = db.sp.threads[conversationId];
        if (thread?.status !== 'managed') return 0;
        const messages = conversationMessages(conversationId);
        const lastOut = messages.filter(item => item.direction === 'out').slice(-1)[0]?.receivedAt || '';
        const since = [thread.ownerSeenAt || '', lastOut].sort().slice(-1)[0];
        return messages.filter(item => item.direction === 'in' && item.receivedAt > since).length;
    }
    // Ogni movimento: prima ServiceNow (change), poi copia su SharePoint e registro.
    function applyOnServiceNow(number, by, type, details, change) {
        const incident = requireIncident(number);
        change(incident);
        mirror(number);
        updateThreadForTicket(number);
        record(type, number, by, details);
        return incident;
    }

    // ------------------------------------------------------------------
    // Triangolazione (tutta su SharePoint e sulle mail: nessuna chiamata a ServiceNow)
    // ------------------------------------------------------------------
    const conversationIds = () => [...new Set(db.messages.map(item => item.conversationId))];
    const conversationMessages = conversationId => db.messages
        .filter(item => item.conversationId === conversationId)
        .sort((a, b) => a.receivedAt.localeCompare(b.receivedAt));
    const lastIncomingOf = conversationId => conversationMessages(conversationId).filter(item => item.direction === 'in').slice(-1)[0] || null;
    function normalizeSubject(subject) {
        return String(subject || '').toLowerCase()
            .replace(/^\s*((re|r|i|fw|fwd|inoltro)\s*:\s*)+/i, '')
            .replace(/\[[^\]]*\]/g, ' ')
            .replace(/\s+/g, ' ').trim();
    }
    // Citazioni INC con la data della mail che le contiene, dalla più recente.
    function incMentions(messages) {
        const mentions = [];
        [...messages].reverse().forEach(message => {
            `${message.subject}\n${message.body}`.match(INC_PATTERN)?.forEach(number => mentions.push({ number: number.toUpperCase(), when: message.receivedAt }));
        });
        return mentions;
    }
    // Ticket simili nella copia su SharePoint: stesso oggetto e Reported vicino all'arrivo della mail.
    function similarTickets(subject, times, excluded = []) {
        const wanted = normalizeSubject(subject);
        if (!wanted) return [];
        return Object.values(db.sp.tickets).filter(ticket => {
            if (excluded.includes(ticket.number)) return false;
            if (Object.values(db.sp.threads).some(thread => threadNumbers(thread).includes(ticket.number))) return false;
            const description = normalizeSubject(ticket.shortDescription);
            if (!description || !(description.includes(wanted) || wanted.includes(description))) return false;
            const reported = parseDate(ticket.reportedOn);
            return reported && times.some(time => Math.abs(reported - time) <= SIMILAR_MINUTES * MINUTE);
        });
    }
    // Risultato per una conversazione: { tickets, previous, candidates, thread }.
    // tickets: i ticket aperti del thread (anche più di uno), ognuno con il suo "via";
    // se non ce ne sono, l'ultimo ticket chiuso (o "previous" se dopo è arrivata una mail nuova).
    function resolveConversation(conversationId, extraNumbers = []) {
        const messages = conversationMessages(conversationId);
        const incoming = messages.filter(item => item.direction === 'in');
        const lastIncoming = incoming[incoming.length - 1] || null;
        const thread = db.sp.threads[conversationId] || null;
        const evidence = [
            ...incMentions(messages).map(item => ({ ...item, via: 'inc' })),
            ...extraNumbers.map(number => ({ number: String(number).toUpperCase(), when: '', via: 'inc' })),
            ...(thread?.links || []).map(link => ({ number: link.number, when: link.linkedAt, via: link.confirmedBy ? 'confirmed' : 'link' }))
        ].filter(item => db.sp.tickets[item.number]).sort((a, b) => b.when.localeCompare(a.when));
        const viaOf = {};
        evidence.forEach(item => { viaOf[item.number] ||= item.via; });
        const known = Object.keys(viaOf).map(number => db.sp.tickets[number]);
        const open = known.filter(ticket => OPEN_STATES.includes(ticket.state)).sort((a, b) => String(a.openedAt).localeCompare(String(b.openedAt)));
        let tickets = open;
        let previous = null;
        if (!open.length && evidence.length) {
            const latest = db.sp.tickets[evidence[0].number];
            if (lastIncoming && latest.closedAt && lastIncoming.receivedAt > latest.closedAt) previous = latest;
            else tickets = [latest];
        }
        let candidates = [];
        if (!tickets.length && messages.length) {
            const times = (previous ? [lastIncoming] : [incoming[0], lastIncoming]).filter(Boolean).map(item => new Date(item.receivedAt));
            candidates = similarTickets(messages[0].subject, times, previous ? [previous.number] : []);
        }
        return {
            tickets: tickets.map(ticket => ({ ...decorate(ticket, conversationId), via: viaOf[ticket.number] })),
            previous: previous ? decorate(previous, conversationId) : null,
            candidates: candidates.map(item => decorate(item)),
            thread
        };
    }
    // Ticket (copia SharePoint) con la mail collegata.
    // mailUnconfirmed: la mail è stata trovata solo per somiglianza e va confermata dall'operatore.
    function decorate(ticket, conversationId) {
        const linked = conversationId || conversationOf(ticket.number);
        const similar = linked ? null : similarConversationOf(ticket);
        const conversation = linked || similar;
        const messages = conversation ? conversationMessages(conversation) : [];
        return {
            ...ticket,
            conversationId: conversation || null,
            messageId: messages.length ? messages[messages.length - 1].id : null,
            mailUnconfirmed: Boolean(similar),
            owner: conversation ? db.sp.threads[conversation]?.owner || '' : '',
            newMessages: conversation ? newMessagesOf(conversation) : 0,
            syncedAt: db.sp.syncedAt
        };
    }
    function similarConversationOf(ticket) {
        return conversationIds().find(id => {
            const messages = conversationMessages(id);
            const times = messages.filter(item => item.direction === 'in').map(item => new Date(item.receivedAt));
            return similarTickets(messages[0].subject, times).some(item => item.number === ticket.number);
        }) || null;
    }
    function conversationOf(number) {
        return Object.values(db.sp.threads).find(thread => threadNumbers(thread).includes(number))?.conversationId
            || conversationIds().find(id => incMentions(conversationMessages(id)).some(item => item.number === number)) || null;
    }
    function requireIncident(number) {
        const incident = db.incidents[number];
        if (!incident) throw new Error(`Ticket ${number} non trovato su ServiceNow.`);
        return incident;
    }
    function addHistory(number, author, kind, text) {
        (db.history[number] ||= []).push({ at: now(), author, kind, text });
    }
    // Riga In coda per un thread (o null). Regole:
    //   - thread in gestione → no;
    //   - thread scartato → solo se dopo è arrivato un messaggio nuovo;
    //   - thread libero (ticket chiuso) → solo se dopo la chiusura è arrivato un messaggio nuovo;
    //   - ticket simile già su ServiceNow → no (si conferma dalla vista Mail o da Ticket).
    function queueItem(conversationId) {
        const messages = conversationMessages(conversationId);
        const lastIncoming = lastIncomingOf(conversationId);
        if (!lastIncoming) return null;
        const thread = db.sp.threads[conversationId];
        if (thread?.status === 'managed') return null;
        if (thread?.status === 'discarded' && thread.discardedAt > lastIncoming.receivedAt) return null;
        if (thread?.status === 'free' && thread.freedAt && lastIncoming.receivedAt <= thread.freedAt) return null;
        const result = resolveConversation(conversationId);
        if (result.tickets.length || result.candidates.length) return null;
        return {
            conversationId,
            messageId: messages[messages.length - 1].id,
            subject: messages[0].subject,
            sender: lastIncoming.from,
            receivedAt: lastIncoming.receivedAt,
            preview: lastIncoming.body.replace(/\s+/g, ' ').slice(0, 160),
            previous: result.previous ? { number: result.previous.number, state: result.previous.state } : null
        };
    }
    const decoratedTicket = number => decorate(db.sp.tickets[number]);

    db = load();

    // ------------------------------------------------------------------
    // Servizi usati dal pannello (contratto: amp.js blocco 3 e amp-servizi.md).
    // Le letture usano solo SharePoint e le mail; le scritture vanno prima su ServiceNow.
    // ------------------------------------------------------------------
    const AmpBackend = {
        mode: 'simulazione',

        // Thread della casella da prendere in carico (uno per conversationId).
        listQueue() {
            return respond(conversationIds().map(queueItem).filter(Boolean));
        },
        // Coda personale: i thread in gestione di un operatore, con i messaggi nuovi da leggere.
        listMyThreads(user) {
            const name = user?.displayValue || '';
            const mine = Object.values(db.sp.threads).filter(thread => thread.status === 'managed' && thread.owner === name);
            return respond(mine.map(thread => {
                const messages = conversationMessages(thread.conversationId);
                const lastIncoming = lastIncomingOf(thread.conversationId);
                return {
                    conversationId: thread.conversationId,
                    messageId: messages[messages.length - 1]?.id || null,
                    subject: messages[0]?.subject || '',
                    sender: lastIncoming?.from || null,
                    receivedAt: messages[messages.length - 1]?.receivedAt || '',
                    tickets: threadNumbers(thread).map(number => db.sp.tickets[number])
                        .filter(ticket => ticket && OPEN_STATES.includes(ticket.state)).map(ticket => ({ number: ticket.number, state: ticket.state, holdExpiration: ticket.holdExpiration })),
                    newMessages: newMessagesOf(thread.conversationId)
                };
            }));
        },
        // Chi ha il thread in carico l'ha letto dal pannello: i messaggi fin qui non sono più "nuovi".
        markThreadSeen(conversationId, user) {
            const thread = db.sp.threads[conversationId];
            if (!thread || thread.owner !== user?.displayValue || !newMessagesOf(conversationId)) return respond(false);
            thread.ownerSeenAt = now();
            persist();
            return respond(true);
        },
        // Ticket aperti dalla copia su SharePoint (In Progress e On Hold), con la mail se collegata.
        listOpenTickets() {
            return respond(Object.values(db.sp.tickets).filter(item => OPEN_STATES.includes(item.state)).map(item => decorate(item)));
        },
        // Triangolazione per la mail aperta: { tickets, previous, candidates, discarded }.
        findTicketForMail(mail) {
            const result = resolveConversation(mail.conversationId, mail.incNumbers || []);
            const lastIncoming = lastIncomingOf(mail.conversationId);
            const thread = result.thread;
            result.discarded = thread?.status === 'discarded' && (!lastIncoming || thread.discardedAt > lastIncoming.receivedAt)
                ? { by: thread.discardedBy, at: thread.discardedAt } : null;
            delete result.thread;
            return respond(result);
        },
        getTicket(number) {
            return db.sp.tickets[number] ? respond(decoratedTicket(number)) : fail(`Ticket ${number} non trovato.`);
        },
        // L'operatore conferma che la mail appartiene a un ticket esistente → SharePoint.
        linkTicket(conversationId, number, by) {
            if (!db.sp.tickets[number]) return fail(`Ticket ${number} non trovato.`);
            registerThread(conversationId, number, by?.displayValue, now(), true);
            record('collegamento', number, by?.displayValue, 'Mail collegata dall\'operatore', conversationId);
            persist();
            const result = resolveConversation(conversationId);
            delete result.thread;
            return respond(result);
        },
        // Apre l'Incident su ServiceNow; solo se riesce aggiorna SharePoint (copia, thread, movimento).
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
                openedAt: now(), timeWorkedMinutes: Number(payload.fields.timeWorkedMinutes) || 0,
                holdReason: '', holdExpiration: '', closedAt: '', resolutionCode: '', fields: payload.fields
            };
            db.history[number] = [];
            addHistory(number, assignee.displayValue, 'system', `Ticket aperto dall'add-in (${payload.companyName} · ${payload.activityName}).`);
            addHistory(number, mail.sender?.name || 'Cliente', 'comment', payload.fields.description || '');
            mirror(number);
            registerThread(mail.conversationId, number, assignee.displayValue);
            record('apertura', number, assignee.displayValue, `${payload.companyName} · ${payload.activityName}`, mail.conversationId);
            persist();
            return respond(decorate(db.sp.tickets[number], mail.conversationId));
        },
        // Prende in carico un ticket di un collega: cambia Assigned to su ServiceNow.
        assignTicket(number, assignee) {
            try {
                applyOnServiceNow(number, assignee.displayValue, 'assegnazione', `Assegnato a ${assignee.displayValue}`, incident => {
                    incident.assignedTo = { name: assignee.displayValue, sysId: assignee.value };
                    addHistory(number, assignee.displayValue, 'system', `Assegnato a ${assignee.displayValue}.`);
                });
                // Il thread passa a chi prende il ticket.
            Object.values(db.sp.threads).filter(thread => threadNumbers(thread).includes(number)).forEach(thread => { thread.owner = assignee.displayValue; });
                persist();
                return respond(decoratedTicket(number));
            } catch (error) { return fail(error.message); }
        },
        addTime(number, deltaMinutes, by) {
            try {
                let applied = 0;
                applyOnServiceNow(number, by, 'tempo', `${deltaMinutes > 0 ? '+' : ''}${deltaMinutes} minuti`, incident => {
                    const before = incident.timeWorkedMinutes;
                    incident.timeWorkedMinutes = Math.max(0, before + Number(deltaMinutes));
                    applied = incident.timeWorkedMinutes - before;
                    addHistory(number, by, 'system', `Tempo lavorato: ${applied >= 0 ? '+' : ''}${applied} minuti (totale ${incident.timeWorkedMinutes}).`);
                });
                persist();
                return respond(decoratedTicket(number));
            } catch (error) { return fail(error.message); }
        },
        // Storico completo di ServiceNow (commenti di tutti): letto solo su richiesta.
        getTicketActivity(number, { offset = 0, limit = 5 } = {}) {
            const all = (db.history[number] || []).slice().reverse();
            return respond({ entries: all.slice(offset, offset + limit), total: all.length });
        },
        // Azione da eseguire all'invio della risposta (solo SharePoint).
        setPendingAction(number, action) {
            const ticket = db.sp.tickets[number];
            if (!ticket) return fail(`Ticket ${number} non trovato.`);
            ticket.pendingAction = action ? { ...action, at: now() } : null;
            record('azione all\'invio', number, action?.by || '', action ? action.type : 'tolta');
            persist();
            return respond(decoratedTicket(number));
        },
        // Thread che non richiede un ticket (SharePoint). Torna In coda con un messaggio nuovo.
        discardMail(conversationId, by) {
            const thread = db.sp.threads[conversationId] || { conversationId, links: [], owner: '', ownerSeenAt: '' };
            Object.assign(thread, { status: 'discarded', discardedBy: by?.displayValue || '', discardedAt: now(), updatedAt: now() });
            db.sp.threads[conversationId] = thread;
            record('ignorata', '', by?.displayValue, '', conversationId);
            persist();
            return respond({ by: thread.discardedBy, at: thread.discardedAt });
        },
        restoreMail(conversationId) {
            const thread = db.sp.threads[conversationId];
            if (thread) Object.assign(thread, { status: 'free', freedAt: '', discardedBy: '', discardedAt: '', updatedAt: now() });
            record('rimessa in coda', '', '', '', conversationId);
            persist();
            return respond(true);
        },
        subscribe(listener) {
            listeners.add(listener);
            return () => listeners.delete(listener);
        },

        // --------------------------------------------------------------
        // Allineamento periodico (nella Function App: un job ogni 10-15 minuti).
        // Una sola lettura su ServiceNow: i ticket del gruppo modificati dall'ultimo
        // allineamento. Aggiorna la copia su SharePoint e libera / riprende i thread.
        // --------------------------------------------------------------
        sync() {
            let changed = 0;
            Object.keys(db.incidents).forEach(number => {
                const current = JSON.stringify(snowView(db.incidents[number]));
                const stored = db.sp.tickets[number] ? JSON.stringify(snowView(db.sp.tickets[number])) : '';
                if (current === stored) return;
                mirror(number);
                updateThreadForTicket(number);
                record('allineamento', number, 'ServiceNow', 'Modifica fatta fuori dal pannello');
                changed += 1;
            });
            db.sp.syncedAt = now();
            persist();
            return respond({ changed, syncedAt: db.sp.syncedAt });
        },

        // --------------------------------------------------------------
        // Solo anteprima: la finta casella di Outlook e le prove
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
            // quando vede la mail inviata nella conversazione → ServiceNow, poi SharePoint.
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
                    body, receivedAt: now(), categories: [], direction: 'out'
                });
                // Ticket aggiornati dalla risposta: quelli con un'azione salvata e quelli citati nel
                // testo (Rif. INC…); se nessuno, il più recente dei ticket aperti del thread.
                const open = resolveConversation(original.conversationId).tickets.filter(ticket => OPEN_STATES.includes(ticket.state));
                const cited = (body.match(INC_PATTERN) || []).map(number => number.toUpperCase());
                let targets = open.filter(ticket => db.sp.tickets[ticket.number].pendingAction || cited.includes(ticket.number));
                if (!targets.length && open.length) targets = [open[open.length - 1]];
                const applied = targets.map(ticket => applyReply(ticket.number, body, operatorName));
                const conversationThread = db.sp.threads[original.conversationId];
                if (conversationThread?.owner === operatorName) conversationThread.ownerSeenAt = now();
                persist();
                return { applied };
            },
            // Un collega lavora direttamente su AMP: ServiceNow cambia, SharePoint no (fino all'allineamento).
            simulateExternalChange() {
                const changes = [];
                const luca = db.incidents.INC0012090;
                if (luca && OPEN_STATES.includes(luca.state)) {
                    Object.assign(luca, { state: '6', resolutionCode: 'Solved (Permanently)', closedAt: now() });
                    addHistory('INC0012090', 'Luca Neri', 'system', 'Stato: Resolved (modificato su AMP).');
                    changes.push('INC0012090 risolto da Luca su AMP');
                }
                const printer = db.incidents.INC0012101;
                if (printer) {
                    printer.timeWorkedMinutes += 15;
                    addHistory('INC0012101', 'Giulia Verdi', 'system', `Tempo lavorato: +15 minuti (totale ${printer.timeWorkedMinutes}) su AMP.`);
                    changes.push('INC0012101: +15 minuti da Giulia su AMP');
                }
                try { sessionStorage.setItem(STORAGE_KEY, JSON.stringify(db)); } catch (error) { /* solo anteprima */ }
                return changes;
            },
            movements() {
                return copy(db.sp.movements);
            },
            reset() {
                db = createDatabase();
                persist();
            }
        }
    };

    // Cosa succede quando parte la risposta (logica "lato server"): ServiceNow, poi SharePoint.
    function applyReply(number, body, operator) {
        const action = db.sp.tickets[number]?.pendingAction || { type: 'comment' };
        const incident = applyOnServiceNow(number, operator, 'risposta', action.type, item => {
            addHistory(number, operator, 'comment', body);
            if (action.workNotes) addHistory(number, operator, 'work_note', action.workNotes);
            if (action.type === 'resolve' || action.type === 'cancel') {
                Object.assign(item, {
                    state: action.type === 'resolve' ? '6' : '8', resolutionCode: action.resolutionCode || '',
                    resolutionNotes: body, holdReason: '', holdExpiration: '', closedAt: now()
                });
                addHistory(number, operator, 'system', `Stato: ${STATE_NAMES[item.state]}${action.resolutionCode ? ` (${action.resolutionCode})` : ''}. Resolution notes = testo della mail.`);
            } else if (action.type === 'hold') {
                Object.assign(item, { state: '3', holdReason: action.holdReason, holdExpiration: action.holdExpiration });
                addHistory(number, operator, 'system', `Stato: On Hold fino al ${action.holdExpiration}.`);
            } else if (item.state === '3') {
                Object.assign(item, { state: '2', holdReason: '', holdExpiration: '' });
                addHistory(number, operator, 'system', 'Stato: In Progress.');
            }
        });
        db.sp.tickets[number].pendingAction = null;
        return { type: action.type, state: incident.state, number };
    }

    window.AmpBackend = AmpBackend;
})();
