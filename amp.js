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
            hold: '#8a6d00', // On Hold: citrino
            holdBg: '#fbf3cc',
            forward: '#7c5bbf', // Inoltra (In coda): pulsante ed esito
            // Selettore della scheda aperta (Inoltra, Ignora): pillola e scritta della voce scelta
            choiceForwardBg: '#f1ecf9',
            choiceForwardText: '#6b47b0',
            choiceDangerBg: '#fde7e9',
            choiceDangerText: '#a4262c',
            choiceResolveBg: '#e3f1e9', // Risolvi: giada
            choiceResolveText: '#22693f',
            choiceHoldBg: '#fbf3cc', // On Hold: citrino
            choiceHoldText: '#8a6d00',
            // Stile: valori dello stile flat. Lo stile vetro ridefinisce gli stessi nomi in glassColors e glassDarkColors;
            // gli elementi usano solo questi nomi, così flat/vetro e chiaro/scuro cambiano tutto insieme.
            backdrop: '#f5f5f5', // sfondo dell'intero pannello
            blur: 'none', // sfocatura di quello che sta dietro alle superfici
            navBlur: 'none', // sfocatura della fascia blu (intestazione e schede)
            hoverFilter: 'none', // superficie al passaggio del mouse: si accende
            hoverInnerFilter: 'none', // etichette e notifiche dentro la scheda: si accendono meno della scheda
            chipLine: 'transparent', // bordino di etichette e stati dentro le schede (nel vetro chiaro serve a staccarli dal bianco)
            hoverGlow: '0 0 transparent', // bagliore intorno alla superficie al passaggio del mouse (nessuno)
            segmentTrack: 'var(--ao-surface)', // fondo dei selettori a segmenti
            segmentHover: 'transparent', // voce non scelta di un selettore al passaggio del mouse
            choiceTrack: 'var(--ao-bg)', // fondo del selettore delle azioni (Prendi in carico / Inoltra / Ignora)
            selection: 'var(--ao-accent-bg)', // voce scelta nei selettori a segmenti
            selectionLine: 'transparent', // bordo della voce scelta
            selectionShadow: '0 0 transparent', // ombra della voce scelta
            selectionTone: '100%', // quanto colore dell'azione nella voce scelta del selettore delle azioni
            selectionMotion: '.32s cubic-bezier(.3, .8, .2, 1)', // come scivola la voce scelta
            menuBg: 'var(--ao-surface)', // elenchi aperti: sempre pieni, si leggono sopra a tutto
            // + Nuovo ticket: nel flat bordo tratteggiato; nel vetro card come le altre con un bagliore fisso che "chiama",
            // un po' più forte al passaggio del mouse; aperto tiene il bagliore
            callBorder: '1px dashed var(--ao-accent)',
            callOpenBorder: 'var(--ao-accent)',
            callGlow: '0 0 transparent',
            callGlowHover: '0 0 transparent',
            searchBg: 'var(--ao-surface)', // ricerca aperta: campo
            searchLensBg: 'var(--ao-accent-bg)', // ricerca aperta: lente
            searchLine: 'var(--ao-accent-border)',
            searchFocusLine: 'var(--ao-accent)',
            notifyInk: 'var(--ao-surface)', // numero del segnalino degli aggiornamenti: colore dello sfondo
            fieldFocusLine: 'var(--ao-accent)', // campo che si sta compilando
            fieldFocusGlow: 'var(--ao-shadow-focus)',
            searchLensLine: 'var(--ao-accent-border)',
            // Colori delle azioni (pulsanti, esiti): nel vetro diventano pietre preziose
            toneTake: 'var(--ao-accent)', // Prendi in carico, Aggiorna, Salva…
            toneForward: 'var(--ao-forward)', // Inoltra
            toneIgnore: '#d13438', // Ignora, Annulla
            toneResolve: '#2e8b57', // Risolvi: giada
            toneHold: '#d4ac0d', // On Hold: citrino
            // Pulsanti pieni del colore dell'azione (nel vetro appena trasparenti), scritta bianca
            toneFill: '100%', // quanto colore nello sfondo
            toneSolid: '100%', // quanto colore nella scheda a esito concluso
            toneLine: 'transparent', // bordo dei pulsanti
            toneHoverFilter: 'brightness(.92)',
            btnHighlight: '0 0 transparent', // filo di luce sul bordo alto dei pulsanti
            spinnerTrack: '45%',
            // Etichette: quanto colore nel fondo (su che base) e nella scritta (su che base), bordo
            tagFill: '14%',
            tagBase: '#fff',
            tagInk: '80%',
            tagInkBase: '#000',
            tagLine: 'transparent'
        },
        // Tema scuro: segue Outlook (prefers-color-scheme) oppure la scelta in Impostazioni → Profilo.
        // Sfondo, riquadri, bordi e testi come l'add-in di jak in tema scuro.
        darkColors: {
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
            hold: '#f0d257',
            holdBg: '#3d3410',
            // Tema scuro: viola e rosso accesi quanto il blu (prima risultavano spenti)
            choiceForwardBg: 'color-mix(in srgb, #7c5bbf 26%, #2d2c2b)',
            choiceForwardText: '#b79bff',
            choiceDangerBg: 'color-mix(in srgb, #d13438 24%, #2d2c2b)',
            choiceDangerText: '#ff8f94',
            choiceResolveBg: 'color-mix(in srgb, #2e8b57 26%, #2d2c2b)',
            choiceResolveText: '#7fd1a3',
            choiceHoldBg: 'color-mix(in srgb, #d4ac0d 24%, #2d2c2b)',
            choiceHoldText: '#f0d257',
            backdrop: '#201f1e',
            tagFill: '30%',
            tagBase: '#2d2c2b',
            tagInk: '45%',
            tagInkBase: '#fff'
        },
        // Stile vetro (Impostazioni → Profilo → Stile): superfici semitrasparenti e sfocate sopra uno sfondo sfumato;
        // al passaggio del mouse si accendono, con un bagliore leggero. Gli elementi usano solo questi nomi (superficie,
        // sfondo, bordo, passaggio…): cambiando stile cambiano tutti insieme, impostazioni comprese.
        glassColors: {
            backdrop: 'transparent', // lo sfondo sfumato lo dipinge il contenitore dell'intestazione (glassBackdrop), dietro a tutto
            surface: 'rgba(255, 255, 255, .82)',
            hover: 'rgba(255, 255, 255, .9)',
            bg: 'rgba(110, 140, 190, .1)', // zone incassate (campi raggruppati, dettagli)
            bgSoft: 'rgba(255, 255, 255, .55)',
            line: 'rgba(255, 255, 255, .75)',
            lineLight: 'rgba(0, 0, 0, .07)',
            lineStrong: 'rgba(0, 0, 0, .14)',
            blur: 'blur(14px) saturate(1.2)',
            navBlur: 'blur(16px) saturate(1.3)',
            hoverFilter: 'brightness(1.06)',
            hoverInnerFilter: 'brightness(.96)', // solo nel chiaro: etichette e notifiche non passano sul bianco
            chipLine: 'rgba(0, 50, 110, .16)',
            notifyInk: '#fff', // vetro: superficie trasparente, il numero va su un colore pieno
            hoverGlow: '0 0 12px rgba(120, 170, 255, .2)',
            segmentTrack: 'rgba(214, 224, 240, .6)', // un po' più scuro delle card: la voce scelta si stacca
            segmentHover: 'rgba(255, 255, 255, .4)',
            choiceTrack: 'var(--ao-segment-track)',
            selection: 'var(--ao-surface)', // voce scelta: come una card a riposo
            selectionLine: 'var(--ao-line)',
            selectionShadow: '0 1px 2px rgba(0, 0, 0, .06)',
            selectionTone: '0%',
            selectionMotion: '.44s cubic-bezier(.32, 1.12, .5, 1)', // scivola con un leggero rimbalzo finale, come iOS
            menuBg: '#fbfcfe',
            callBorder: '1px solid var(--ao-line)',
            callOpenBorder: 'var(--ao-line)',
            callGlow: '0 0 18px rgba(80, 150, 255, .5), 0 0 4px rgba(80, 150, 255, .25)',
            callGlowHover: '0 0 21px rgba(80, 150, 255, .58), 0 0 5px rgba(80, 150, 255, .3)',
            searchBg: 'var(--ao-hover)', // ricerca aperta: accesa come una card al passaggio del mouse, senza contorno
            searchLensBg: 'var(--ao-hover)',
            searchLine: 'transparent',
            searchFocusLine: 'transparent',
            fieldFocusLine: 'rgba(120, 170, 255, .45)', // vetro: un bagliore soffuso al posto dell'anello
            fieldFocusGlow: '0 0 12px rgba(120, 170, 255, .35)',
            searchLensLine: 'transparent',
            toneTake: '#0f5fbf', // zaffiro
            toneForward: '#7b3fbf', // ametista
            toneIgnore: '#c4183c', // rubino
            toneFill: '86%', // vetro colorato
            toneSolid: '0%',
            toneLine: 'rgba(255, 255, 255, .22)',
            toneHoverFilter: 'brightness(1.08)',
            btnHighlight: 'inset 0 1px 0 rgba(255, 255, 255, .22)',
            spinnerTrack: '35%',
            tagFill: '22%',
            tagBase: 'transparent',
            tagInk: '80%',
            tagInkBase: '#000',
            tagLine: 'rgba(0, 50, 110, .16)' // come chipLine: un bordo bianco sparirebbe sulle schede bianche
        },
        glassDarkColors: {
            surface: 'rgba(52, 56, 66, .62)',
            hover: 'rgba(70, 77, 92, .68)',
            bg: 'rgba(0, 0, 0, .2)',
            bgSoft: 'rgba(40, 44, 52, .5)',
            line: 'rgba(255, 255, 255, .09)',
            lineLight: 'rgba(255, 255, 255, .07)',
            lineStrong: 'rgba(255, 255, 255, .16)',
            hoverFilter: 'brightness(1.22)',
            hoverInnerFilter: 'none', // nello scuro resta com'è
            chipLine: 'transparent',
            notifyInk: '#23262e',
            hoverGlow: '0 0 12px rgba(90, 150, 255, .18)',
            segmentTrack: 'rgba(28, 31, 38, .62)',
            segmentHover: 'rgba(255, 255, 255, .06)',
            selectionShadow: '0 1px 2px rgba(0, 0, 0, .3)',
            menuBg: '#2b2e36',
            fieldFocusLine: 'rgba(90, 150, 255, .4)',
            fieldFocusGlow: '0 0 12px rgba(90, 150, 255, .3)',
            callGlow: '0 0 18px rgba(70, 140, 255, .5), 0 0 4px rgba(70, 140, 255, .25)',
            callGlowHover: '0 0 21px rgba(70, 140, 255, .58), 0 0 5px rgba(70, 140, 255, .3)',
            // Scuro: le stesse pietre, scurite come l'intestazione; niente riflesso sul bordo
            toneTake: 'color-mix(in srgb, #0f5fbf 82%, #000)',
            toneForward: 'color-mix(in srgb, #7b3fbf 82%, #000)',
            toneIgnore: 'color-mix(in srgb, #c4183c 82%, #000)',
            toneResolve: 'color-mix(in srgb, #2e8b57 82%, #000)',
            toneHold: 'color-mix(in srgb, #d4ac0d 82%, #000)',
            toneLine: 'rgba(255, 255, 255, .14)',
            btnHighlight: '0 0 transparent',
            tagFill: '28%',
            tagInk: '45%',
            tagInkBase: '#fff',
            tagLine: 'rgba(255, 255, 255, .12)'
        },
        // Etichette degli operatori (segnaposto su In coda, ticket senza mail): un colore a testa, scelto in Profilo.
        // Ogni colore si adatta allo stile: fondo e scritta si ricavano dal colore con i valori label* qui sotto.
        labels: {
            rosso: '#d13438', arancio: '#ca5010', ambra: '#c19c00', verde: '#107c10', acqua: '#038387',
            azzurro: '#0078d4', indaco: '#4f6bed', viola: '#8764b8', rosa: '#c239b3', grafite: '#69797e'
        },
        // Stile vetro fuori dall'area AMP: sfondo sfumato (dietro intestazione e pannello) e intestazione dell'add-in
        glassBackdrop: 'radial-gradient(ellipse 260px 220px at 8% 4%, rgba(121, 184, 255, .45), transparent 70%), radial-gradient(ellipse 240px 220px at 95% 40%, rgba(185, 168, 245, .32), transparent 70%), linear-gradient(180deg, #e9f1fb, #eef0fa) fixed',
        glassDarkBackdrop: 'radial-gradient(ellipse 260px 220px at 8% 4%, rgba(40, 100, 190, .45), transparent 70%), radial-gradient(ellipse 240px 220px at 95% 40%, rgba(110, 80, 190, .3), transparent 70%), linear-gradient(180deg, #1b1e26, #1d1b24) fixed',
        // Intestazione dell'add-in nel vetro, per modo: occupa tutta la larghezza, come prima, ed è dello stesso vetro
        // del pannello, tinto di blu (riflesso diagonale, filo di luce in alto). La barra delle schede la prolunga.
        glassHeader: {
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
        back: 'Torna all’add-in',
        system: {
            servicesOk: 'Servizi collegati',
            servicesMissing: 'Servizi non ancora collegati',
            serviceError: message => `Servizi non raggiungibili: ${message}`,
            synced: when => `Allineato ${when}`,
            syncHint: 'Ultimo allineamento della copia su SharePoint con ServiceNow'
        },
        tabs: { queue: 'In coda', tickets: 'Ticket', mail: 'Mail', settings: 'Settings' },
        common: {
            me: 'me',
            back: view => `← ${view}`,
            loading: 'Caricamento…',
            cancel: 'Annulla',
            save: 'Salva',
            edit: 'Modifica',
            delete: 'Elimina',
            confirm: 'Confermi?',
            close: 'Chiudi',
            add: 'Aggiungi',
            none: '—',
            unassigned: 'Non assegnato',
            servicesMissing: 'I servizi non sono ancora collegati: qui compariranno i dati della casella condivisa e di ServiceNow.',
            serviceError: message => `Servizio non disponibile: ${message}`
        },
        queue: {
            labelAdd: 'Metti la tua etichetta',
            labelRemove: 'Togli la tua etichetta',
            labelTake: name => `${name}: clic per prenderla tu`,
            labelTakeConfirm: 'Clic di nuovo per conferma',
            labeledBy: name => `Già segnata da ${name}`,
            ignoredDone: 'Ignorata',
            forward: 'Inoltra',
            forwardWorking: 'Inoltro…',
            forwardDone: group => `Inoltrata a ${group}`,
            goToTicket: 'Vai al ticket',
            empty: 'Nessuna richiesta in coda.',
            active: 'Attive',
            ignoredTab: 'Ignorate',
            ignoredEmpty: 'Nessun thread ignorato.',
            ignoredHint: 'Tornano da sole In coda quando arriva un messaggio nuovo nel thread; intanto puoi prenderle in carico da qui.',
            ignoredBy: (who, when) => `Ignorata da ${who} · ${when}`,

            takeCharge: 'Prendi in carico',
            ignore: 'Ignora'
        },
        tickets: {
            newTicket: 'Nuovo ticket',
            shortDescription: 'Short description',
            description: 'Description',
            receivedDate: 'Ricevuta il',
            receivedTime: 'alle (hh:mm)',
            openNew: 'Apri ticket',
            missingShort: 'Scrivi la Short description.',
            badTime: "Scrivi l'ora come hh:mm (per esempio 09:30).",
            mine: 'Miei',
            others: 'Altri',
            fromMail: 'Da mail',
            withoutMail: 'Senza mail',
            allOperators: 'Tutti',
            status: 'Stato',
            operator: 'Operatore',
            statuses: { all: 'Tutti', open: 'Aperti', hold: 'In attesa' },
            defaultFilter: 'predefinito',
            clearFilter: 'Togli il filtro',
            filters: 'Filtri',
            search: 'Cerca',
            searchPlaceholder: 'INC o oggetto',
            noResults: 'Nessun risultato.',
            empty: 'Nessun ticket aperto.',
            noMail: 'senza mail',
            mailToConfirm: 'mail da confermare',
            newMessagesTitle: n => (n === 1 ? '1 messaggio nuovo nel thread' : `${n} messaggi nuovi nel thread`),
            threadGroup: n => `Stesso thread · ${n} ticket`,
            holdExpired: 'On Hold scaduto',
            holdUntil: date => `fino al ${date}`
        },
        mail: {
            outcomeWorking: 'Apro il ticket…',
            outcomeTaken: number => `${number} preso in carico`,
            noMail: 'Apri una mail in Outlook, oppure scegli una richiesta dalla Coda o un ticket.',
            searching: 'Ricerca del ticket su ServiceNow…',
            openInOutlook: 'Apri la mail',
            openFirst: 'Apri la mail in Outlook per prenderla in carico o ignorarla.',
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
            reply: 'Risposta',
            updateText: 'Testo',
            updateButtons: { comment: 'Aggiorna', resolve: 'Risolvi', hold: 'Metti in On Hold', cancel: 'Annulla il ticket' },
            updated: number => `${number} aggiornato`,
            closedTicket: number => `${number} chiuso`,
            actions: { comment: 'Aggiorna', resolve: 'Risolvi', hold: 'On Hold', cancel: 'Annulla' },
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
            style: 'Stile',
            styles: { flat: 'Flat', glass: 'Vetro' },
            label: 'Etichetta',
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
    const BUILD = '2026-11-06 12:50';

    function createDefaultConfig() {
        const now = new Date().toISOString();
        return {
            version: STORAGE.version, created: now, updated: now,
            profile: { assignedTo: { value: '', displayValue: '' }, label: 'azzurro' },
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
            profile: { assignedTo: normalizeReference(data.profile?.assignedTo), label: THEME.labels[data.profile?.label] ? data.profile.label : 'azzurro' },
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
        // Stile: vetro (predefinito) o flat.
        root.dataset.style = config.preferences?.style === 'flat' ? 'flat' : 'glass';
        blendTabsWithHeader(); // con il vetro cambia anche l'intestazione
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
    //   listQueue()                         → [{ conversationId, messageId, subject, mailbox, sender, receivedAt,
    //                                            preview, previous: { number, state } | null, candidates: [numero] }]
    //   listIgnored()                       → [{ conversationId, messageId, subject, mailbox, sender, receivedAt, discardedBy, discardedAt }]
    //                                          (mailbox = casella condivisa da cui arriva: AMPCoE, M365 Admin Center…)
    //   markThreadSeen(conversationId, utente) → i messaggi del thread non sono più "nuovi" per chi l'ha in carico
    //   listOpenTickets()                   → Ticket[]  (In Progress e On Hold del gruppo, copia su SharePoint)
    //   findTicketForMail(mail)             → { tickets: (Ticket & { via: 'inc'|'link'|'confirmed' })[], previous, candidates: Ticket[], discarded }
    //   getTicket(numero)                   → Ticket
    //   linkTicket(conversationId, numero, utente)  → come findTicketForMail (l'operatore conferma il collegamento)
    //   createTicket(mail, { companyId, companyName, activityId, activityName, fields }) → Ticket
    //                                          (mail = null: ticket aperto da Ticket senza mail, senza thread)
    //   assignTicket(numero, utente)        → Ticket    (Assigned to = utente)
    //   addTime(numero, minuti, nome)       → Ticket    (minuti anche negativi)
    //   getTicketActivity(numero, { offset, limit }) → { entries: [{ at, author, kind, text }], total }
    //   setPendingAction(numero, azione | null)      → Ticket   (eseguita all'invio della risposta)
    //   discardMail(conversationId, utente) / restoreMail(conversationId)
    //   subscribe(listener)                 → funzione per smettere (facoltativo)
    //   getMailBody(messageId)              → testo della mail (Graph, casella condivisa) per l'anteprima in In coda
    //                                          (facoltativo: senza, il pannello mostra le prime righe di listQueue)
    //   forwardMail(conversationId, gruppo, utente) → inoltro della richiesta a un gruppo (ES, PEL, PO); il thread lascia In coda
    //                                          (facoltativo: senza, Inoltra avvisa che i servizi non sono collegati)
    //   labelMail(conversationId, utente, colore) → { ok, label: { name, color, at } }: etichetta di chi gestisce la mail
    //                                          (una per thread: se c'è già quella di un altro, ok = false e torna quella)
    //   unlabelMail(conversationId, utente)  → toglie la propria etichetta
    //   updateTicket(numero, azione, testo, utente) → Ticket: ticket senza mail, aggiornato direttamente su ServiceNow
    //                                          (azione come setPendingAction, o null = solo commento; testo = commento / Resolution notes)
    //   labelMail(…, { force: true })        → sostituisce l'etichetta di un collega (tenendo premuta la sua, a scheda aperta)
    //   (listQueue e listIgnored: ogni mail ha label = { name, color } | null, che si toglie quando la mail viene
    //    ignorata o inoltrata; i ticket senza mail hanno label di chi li ha aperti)
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
    Services.getMailBody = async messageId => (typeof window.AmpBackend?.getMailBody === 'function' ? window.AmpBackend.getMailBody(messageId) : null);
    ['forwardMail', 'labelMail', 'unlabelMail', 'updateTicket'].forEach(name => {
        Services[name] = async (...args) => {
            if (typeof window.AmpBackend?.[name] !== 'function') throw new ServicesMissingError(name);
            return window.AmpBackend[name](...args);
        };
    });

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
            if (match) return { id: match[0], reason: TEXTS.mail.byDomain(domain), tag: `@${domain}` };
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

    // Gruppi a cui si inoltra da In coda (il tag va davanti all'oggetto della mail: [ES], [PEL], [PO]).
    const FORWARD_GROUPS = [
        { id: 'ES', name: 'Email Security' },
        { id: 'PEL', name: 'Amministratori PEL' },
        { id: 'PO', name: 'Post Office' }
    ];
    const SEARCH_ICON = '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><circle cx="8.5" cy="8.5" r="5"/><path d="M12.5 12.5l4 4"/></svg>';
    const FILTER_ICON = '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"><path d="M3 4h14l-5.5 6.5V16l-3 1.5v-7z"/></svg>';
    // Con i filtri o la ricerca aperti il loro pulsante cambia icona: ricliccandolo si richiudono.
    const FILTER_OFF_ICON = '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round" stroke-linecap="round"><path d="M3 4h14l-5.5 6.5V16l-3 1.5v-7z"/><path d="M2.5 2.5l15 15"/></svg>';
    const CLOSE_ICON = '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M5 5l10 10M15 5L5 15"/></svg>';
    const CHEVRON_UP_ICON = '<svg viewBox="0 0 20 20" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l5-5 5 5"/></svg>';
    // Icone delle azioni (selettori di In coda e dei ticket)
    const ACTION_ICONS = {
        take: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6.5V4h14v2.5a2 2 0 0 0 0 4V16H3v-5.5a2 2 0 0 0 0-4z"/><path d="M10 7.5v5M7.5 10h5"/></svg>',
        forward: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M11.5 4.5l5 5-5 5"/><path d="M16.5 9.5H8a4.5 4.5 0 0 0-4.5 4.5v1.5"/></svg>',
        ignore: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><rect x="2.5" y="3.5" width="15" height="4" rx="1"/><path d="M4 7.5v8a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1v-8M8 11h4"/></svg>',
        comment: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M3.5 4.5h13v8.5h-7l-3.5 3v-3H3.5z"/><path d="M7 8h6M7 10.5h4"/></svg>',
        resolve: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M4.5 10.5l3.5 3.5 7.5-8"/></svg>',
        hold: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M7 4.5v11M13 4.5v11"/></svg>',
        cancel: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M4 6h12M8 6V4h4v2M5.5 6l.8 10h7.4l.8-10"/></svg>'
    };
    // Colore del pulsante per ogni azione sul ticket
    const ACTION_FILLS = { comment: '', resolve: ' ao-fill-resolve', hold: ' ao-fill-hold', cancel: ' ao-fill-danger' };
    const ACTION_TONES = { take: 'accent', forward: 'forward', ignore: 'danger', comment: 'accent', resolve: 'resolve', hold: 'hold', cancel: 'danger' };
    const SWAP_ICON = '<svg viewBox="0 0 20 20" width="11" height="11" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 10h12M12 6l4 4-4 4"/></svg>';
    const TAG_ADD_ICON = '<svg viewBox="0 0 20 20" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3h6.5l7 7-6.5 6.5-7-7z"/><circle cx="6.8" cy="6.8" r="1.1"/><path d="M15.5 2.5v4M13.5 4.5h4"/></svg>';
    function svgIcon(markup) {
        const icon = h('span', { class: 'ao-svg-icon', 'aria-hidden': 'true' });
        icon.innerHTML = markup;
        return icon;
    }

    // Icone delle schede principali
    const TAB_ICONS = {
        settings: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 0 0 2.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 0 0 1.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 0 0-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 0 0-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 0 0-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 0 0-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 0 0 1.066-2.573c-.94-1.543.826-3.31 2.37-2.37 1 .608 2.296.07 2.572-1.065z"/><circle cx="12" cy="12" r="3"/></svg>',
        queue: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"><path d="M3 11l2.5-6h9L17 11v5H3z"/><path d="M3 11h4l1 2h4l1-2h4"/></svg>',
        tickets: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"><path d="M3 6.5V4h14v2.5a2 2 0 0 0 0 4V16H3v-5.5a2 2 0 0 0 0-4z"/><path d="M8 8h4M8 11h4"/></svg>',
        debug: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M12.5 3.5a3.5 3.5 0 0 0-3.3 4.6L3.5 13.8a1.5 1.5 0 0 0 2.1 2.1l5.7-5.7a3.5 3.5 0 0 0 4.6-3.3l-2 2-2-.5-.5-2z"/></svg>'
    };
    // Icone del selettore delle Impostazioni
    const SETTINGS_ICONS = {
        profile: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><circle cx="10" cy="7" r="3.2"/><path d="M4 16.5c.8-3 3.2-4.5 6-4.5s5.2 1.5 6 4.5"/></svg>',
        companies: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="2.5" y="6" width="15" height="10" rx="1.5"/><path d="M7 6V4h6v2M2.5 10.5h15"/></svg>',
        activities: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M8 5.5h8.5M8 10h8.5M8 14.5h8.5"/><path d="M3.5 5.5l1 1 1.8-2M3.5 10l1 1 1.8-2M3.5 14.5l1 1 1.8-2"/></svg>',
        responses: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7h8a4 4 0 0 1 0 8H9"/><path d="M7 4L4 7l3 3"/></svg>'
    };
    // Icone dei selettori a segmenti
    const SEGMENT_ICONS = {
        active: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><circle cx="10" cy="10" r="7"/><path d="M10 6v4l2.5 2"/></svg>',
        ignored: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="2.5" y="3.5" width="15" height="4" rx="1"/><path d="M4 7.5v8a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1v-8M8 11h4"/></svg>',
        mail: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="2.5" y="4.5" width="15" height="11" rx="1.5"/><path d="M3 5.5l7 5 7-5"/></svg>',
        none: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6.5V4h14v2.5a2 2 0 0 0 0 4V16H3v-5.5a2 2 0 0 0 0-4z"/><path d="M10 7.5v5M7.5 10h5"/></svg>'
    };
    // Icone dei selettori di tema e stile (Impostazioni → Profilo)
    const APPEARANCE_ICONS = {
        auto: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><circle cx="10" cy="10" r="7"/><path d="M10 3a7 7 0 0 1 0 14z" fill="currentColor" stroke="none"/></svg>',
        light: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><circle cx="10" cy="10" r="3.2"/><path d="M10 2.5v1.8M10 15.7v1.8M2.5 10h1.8M15.7 10h1.8M4.7 4.7L6 6M14 14l1.3 1.3M4.7 15.3L6 14M14 6l1.3-1.3"/></svg>',
        dark: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M16 12.5A6.5 6.5 0 0 1 7.5 4a6.5 6.5 0 1 0 8.5 8.5z"/></svg>',
        flat: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="12" height="12" rx="2"/></svg>',
        glass: '<svg viewBox="0 0 20 20" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="10" height="10" rx="2"/><rect x="7" y="7" width="10" height="10" rx="2" fill="currentColor" fill-opacity=".18"/></svg>'
    };
    function tabIcon(view) {
        if (!TAB_ICONS[view]) return null;
        const icon = h('span', { class: 'ao-tab-icon' });
        icon.innerHTML = TAB_ICONS[view];
        icon.append(notifyBadge(0)); // aggiornamenti: sempre visibili sull'icona
        return icon;
    }

    // Solo stato dell'interfaccia: i dati dei ticket arrivano ogni volta dai servizi.
    const state = {
        view: 'queue', // queue | tickets | mail | settings
        queue: [],
        tickets: [],
        loaded: false,
        loading: false,
        serviceError: '',
        ticketSource: 'mail', // Ticket: mail (aperti da una mail) | none (senza mail)
        ticketOperator: 'mine', // Ticket: mine (vista iniziale) | nome del collega | all
        ticketStatus: 'all', // Ticket: all (vista iniziale) | open (In Progress) | hold (On Hold)
        ticketFilterOpen: false, // imbuto: la barra mostra i filtri al posto di Da mail | Senza mail
        filterMenu: null, // elenco aperto dei filtri: status | operator
        ticketSearchOpen: false, // ricerca (lente) aperta
        ticketQuery: '', // testo cercato: numero INC o oggetto
        ticketSearchShown: '', // testo dei risultati mostrati
        ticketSearchPending: false, // si sta scrivendo: scheletro di caricamento al posto dei risultati
        mail: null, // mail aperta in Outlook
        target: null, // cosa mostra la vista Mail: { kind: 'mail', mail } | { kind: 'ticket', number }
        result: null, // triangolazione della mail mostrata
        searching: false,
        ticketCache: {}, // ultimo Ticket letto per numero (anche chiusi o non in elenco)
        newTicketOpen: false, // modulo "Nuovo ticket" (senza mail) aperto in Ticket
        ticketDrafts: {}, // per conversazione: Agenzia, Attività, Short description
        actionDrafts: {}, // per ticket: azione all'invio
        history: {}, // per ticket: storico caricato
        editing: null, // impostazioni: { kind, id }
        settingsTab: 'profile', // profile | companies | activities | responses
        backTo: null, // vista a cui torna il pulsante ← (Mail e Impostazioni aperte da un'altra vista)
        openedTicket: null, // ticket aperto dall'elenco Ticket: con ← la sua scheda torna al suo posto
        ignored: [], // archivio dei thread ignorati (In coda, in fondo)
        queueTab: 'active', // In coda: active (da prendere) | ignored (archivio)
        extraTicketFor: null, // thread su cui si sta aprendo un ticket in più
        expandedTickets: [], // ticket con i dettagli aperti (thread con più ticket)
        barOffset: 0, // di quanti px la barra secondaria è scivolata via (segue lo scorrimento)
        queueExpanded: null,
        queueChoice: {}, // scheda di In coda aperta: azione scelta nel selettore e gruppo dell'inoltro
        liftFrom: null, // scheda che si sta aprendo: disegnata non sollevata, poi si solleva
        previewScroll: null, // scheda aperta del tutto: il testo della mail può scorrere
        mailBodies: {}, // testo delle mail per l'anteprima in In coda (per messageId)
        holdRender: false, // true durante un'animazione: il pannello non si ridisegna (le schede sparirebbero di colpo)
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
    // Contatore delle schede (solo sulla scheda aperta). Nascosto con una classe, non con hidden,
    // così compare e sparisce con un'animazione.
    const counter = () => h('span', { class: 'ao-counter ao-counter-off' });
    // Segnalino degli aggiornamenti sull'icona (schede e selettori): compare solo se c'è qualcosa di nuovo.
    // Il tondo prende il colore del testo intorno (currentColor), il numero quello dello sfondo.
    const notifyBadge = (updates, title) => h('span', { class: `ao-notify${updates ? '' : ' ao-notify-off'}`, title: updates ? title : undefined }, h('span', { class: 'ao-notify-number' }, updates || ''));
    // Solo la notifica: numero su sfondo blu.
    const newsBadge = (news, title = '') => h('span', { class: 'ao-news', title: title || undefined }, news);
    const chip = (text, tone = 'neutral', title) => h('span', { class: `ao-chip ao-tone-${tone}`, title }, text);
    // Stato del ticket: In Progress è la normalità e non ha etichetta; si vedono solo gli altri stati.
    const stateChip = ticket => {
        if (isHoldExpired(ticket)) return chip(TEXTS.tickets.holdExpired, 'danger');
        if (ticket.state === '2') return null;
        const info = TICKET_STATES[ticket.state] || { label: ticket.state, tone: 'neutral' };
        return chip(info.label, info.tone);
    };
    // Elenco a discesa del pannello (niente <select> di sistema): campo con il valore e la freccetta, elenco come
    // quello dei filtri. L'elenco si apre sopra a tutto (dentro #ao-root, non tagliato dalle schede) e si chiude
    // scegliendo, cliccando fuori, con Esc o scorrendo. Un valore vuoto ('') è il "Seleziona…", in grigio.
    function selectField(label, options, value, onchange, extra = {}) {
        let current = String(value ?? '');
        const text = h('span', { class: 'ao-select-value' });
        const show = () => {
            const option = options.find(item => String(item.value) === current) || options[0];
            text.textContent = option ? option.label : '';
            text.classList.toggle('ao-select-placeholder', !current && options[0]?.value === '');
        };
        show();
        const field = h('div', {
            class: 'ao-input ao-select', role: 'button', tabindex: '0', 'aria-haspopup': 'listbox', 'aria-label': label, ...extra,
            onclick: () => toggleSelect(field, options, () => current, next => { current = String(next); show(); onchange(next); }),
            onkeydown: event => {
                if (event.key === 'Enter' || event.key === ' ' || event.key === 'ArrowDown') { event.preventDefault(); field.click(); }
                if (event.key === 'Escape') closeSelects();
            }
        }, text, h('span', { class: 'ao-select-caret', 'aria-hidden': 'true' }, '▾'));
        return h('div', { class: 'ao-field' }, h('span', { class: 'ao-label' }, label), field);
    }
    function closeSelects() {
        root?.querySelectorAll('.ao-select-menu').forEach(menu => menu.remove());
        root?.querySelectorAll('.ao-select.ao-open').forEach(field => field.classList.remove('ao-open'));
    }
    function toggleSelect(field, options, getValue, pick) {
        const wasOpen = field.classList.contains('ao-open');
        closeSelects();
        if (wasOpen) return;
        const box = root.getBoundingClientRect();
        const rect = field.getBoundingClientRect();
        const menu = h('div', { class: 'ao-menu ao-select-menu', role: 'listbox' },
            options.map(option => h('button', {
                type: 'button', role: 'option', 'aria-selected': String(String(option.value) === getValue()),
                class: `ao-menu-item${String(option.value) === getValue() ? ' ao-active' : ''}${option.value === '' ? ' ao-select-placeholder' : ''}`,
                onclick: event => { event.stopPropagation(); closeSelects(); pick(option.value); field.focus(); }
            }, option.label)));
        // Sotto il campo se c'è spazio, altrimenti sopra; largo come il campo.
        menu.style.left = `${rect.left - box.left}px`;
        menu.style.width = `${rect.width}px`;
        root.append(menu);
        const below = box.bottom - rect.bottom - 8;
        if (menu.offsetHeight > below && rect.top - box.top > below) menu.style.top = `${rect.top - box.top - menu.offsetHeight - 4}px`;
        else menu.style.top = `${rect.bottom - box.top + 4}px`;
        field.classList.add('ao-open');
        menu.querySelector('.ao-active')?.scrollIntoView({ block: 'nearest' });
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
    function openTicketItem(ticket, { slide = true } = {}) {
        if (ticket.newMessages && ticket.conversationId) Services.markThreadSeen(ticket.conversationId, currentUser()).catch(() => {});
        state.backTo = 'tickets';
        state.openedTicket = ticket.number;
        state.target = { kind: 'ticket', number: ticket.number };
        const entering = state.view !== 'mail';
        state.view = 'mail';
        render();
        scrollToTop();
        if (entering && slide) slideInView(1);
        if (ticket.messageId) Outlook.openMessage(ticket.messageId);
    }
    // Dall'elenco alla pagina del ticket e ritorno, con la scheda in primo piano:
    // apertura: la scheda si solleva, sale in cima mentre il resto svanisce, e lì compare la pagina;
    // chiusura (←): la pagina svanisce, la scheda sollevata scende dalla cima al suo posto, si posa e ricompare il resto.
    const LIFT_EASING = 'cubic-bezier(.3, .8, .2, 1)';
    const LIFTED = 'translateY(-2px) scale(1.015)'; // come .ao-item.ao-opening
    // Tutto quello che sta intorno alla scheda (fratelli suoi e dei contenitori, fino all'area che scorre).
    function aroundCard(card) {
        const scroll = root.querySelector('.ao-scroll');
        const around = [];
        for (let node = card; node && node.parentElement && node !== scroll; node = node.parentElement) {
            around.push(...[...node.parentElement.children].filter(sibling => sibling !== node));
        }
        return around;
    }
    // Di quanto deve salire la scheda per arrivare in cima all'area (dove comincia la pagina).
    function riseOf(card) {
        const top = root.querySelector('.ao-scroll').getBoundingClientRect().top + 8;
        return top - card.getBoundingClientRect().top;
    }
    async function openFromCard(card, ticket) {
        if (state.holdRender || !card || reducedMotion()) return openTicketItem(ticket);
        await holdingRender(async () => {
            card.classList.add('ao-opening');
            await wait(MOTION.hold);
            const rise = riseOf(card);
            await Promise.all([
                card.animate([{ transform: LIFTED }, { transform: `translateY(${rise - 2}px) scale(1.015)` }], { duration: MOTION.unfold, easing: LIFT_EASING, fill: 'forwards' }).finished,
                fadeAll(aroundCard(card), 1, 0, { duration: MOTION.unfold * .6, easing: 'linear', fill: 'forwards' })
            ]).catch(() => {});
        });
        openTicketItem(ticket, { slide: false });
        const body = root.querySelector('.ao-body');
        await holdingRender(() => fadeAll([body], 0, 1, { duration: MOTION.swap, easing: 'linear' }));
    }
    async function closeToList(target) {
        const number = state.openedTicket;
        state.openedTicket = null;
        if (state.holdRender || !number || target !== 'tickets' || reducedMotion()) return setView(target);
        const body = root.querySelector('.ao-body');
        const out = body.animate([{ opacity: 1 }, { opacity: 0 }], { duration: MOTION.swap, easing: 'linear', fill: 'forwards' });
        await holdingRender(() => out.finished.catch(() => {}));
        setView(target, null, { slide: false });
        const card = [...root.querySelectorAll('.ao-scroll .ao-item')].find(item => item.querySelector('.ao-item-number')?.textContent === number);
        if (!card) { out.cancel(); return slideInView(-1); }
        const around = aroundCard(card);
        around.forEach(element => { element.style.opacity = '0'; });
        card.scrollIntoView({ block: 'nearest' });
        card.style.transition = 'none';
        card.classList.add('ao-opening');
        const rise = riseOf(card);
        out.cancel();
        await holdingRender(async () => {
            await card.animate([{ transform: `translateY(${rise - 2}px) scale(1.015)` }, { transform: LIFTED }], { duration: MOTION.unfold, easing: LIFT_EASING }).finished.catch(() => {});
            card.style.transition = '';
            card.classList.remove('ao-opening'); // si posa (transizione della scheda)
            await fadeAll(around, 0, 1, { duration: MOTION.swap, easing: 'linear' });
            around.forEach(element => { element.style.opacity = ''; });
        });
    }
    // backTo: vista a cui tornare con il pulsante ← (le schede lo azzerano).
    function setView(view, backTo = null, { slide = true } = {}) {
        const from = state.view;
        state.barOffset = 0;
        state.view = view;
        state.backTo = backTo;
        state.editing = null;
        state.filterMenu = null;
        render();
        scrollToTop();
        if (slide && from !== view) slideInView(VIEW_ORDER.indexOf(view) >= VIEW_ORDER.indexOf(from) ? 1 : -1);
    }
    // Cambio di pagina dalle schede o dall'ingranaggio: ricerca e filtri si chiudono e ogni pagina riparte dalla sua
    // voce predefinita (In coda → Attive, Ticket → Da mail, Impostazioni → Profilo). Il ritorno con ← invece ritrova l'elenco com'era.
    function openPage(view, backTo = null) {
        if (view !== state.view) {
            clearTimeout(searchTimer);
            Object.assign(state, {
                queueTab: 'active', ticketSource: 'mail', settingsTab: 'profile', ticketFilterOpen: false,
                ticketSearchOpen: false, ticketQuery: '', ticketSearchShown: '', ticketSearchPending: false
            });
            ['active|ignored', 'mail|none', Object.keys(TEXTS.settings.tabs).join('|')].forEach(key => { segmentMemory.delete(key); segmentPlace.delete(key); }); // la pillola compare ferma sulla voce predefinita
        }
        setView(view, backTo);
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
        // Una sola barra: freccia per tornare all'add-in e tre schede a posto fisso (In coda, Ticket, Impostazioni).
        // La scheda scelta mostra icona, nome e contatore; le altre solo l'icona (il nome passandoci sopra).
        root = h('div', { id: 'ao-root' },
            h('nav', { class: 'ao-tabs', role: 'tablist' },
                h('button', { type: 'button', class: 'ao-icon-button ao-back', title: TEXTS.back, 'aria-label': TEXTS.back, onclick: () => window.TestPluginAPI?.showMainView?.() }, '←'),
                // La vista di lavoro del ticket non è una scheda: si apre da sola (mail o ticket) e ha il suo "←".
                ['queue', 'tickets', 'settings'].map(view => h('button', {
                    type: 'button', role: 'tab', class: 'ao-tab', dataset: { view }, 'aria-label': TEXTS.tabs[view],
                    onclick: () => openPage(view)
                }, tabIcon(view), h('span', { class: 'ao-tab-label' }, TEXTS.tabs[view]), counter(0))),
                h('span', { class: 'ao-tab-indicator', 'aria-hidden': 'true' })),
            h('div', { class: 'ao-body' }),
            h('div', { class: 'ao-status', role: 'status', 'aria-live': 'polite', hidden: true }),
            // Barra di sistema fissa in fondo: stato dei servizi e ultimo allineamento, sempre visibile.
            h('footer', { class: 'ao-statusbar', role: 'status', 'aria-live': 'polite' }));
        // Elenchi a discesa: un clic fuori o lo scorrimento li chiude.
        root.addEventListener('click', event => {
            if (!event.target.closest('.ao-select, .ao-select-menu')) closeSelects();
        }, true);
        root.addEventListener('scroll', event => { if (!event.target.closest?.('.ao-select-menu')) closeSelects(); }, true);
        // Etichetta di un collega pronta al passaggio: un clic altrove annulla.
        root.addEventListener('click', event => {
            if (!state.labelArmed || event.target.closest('.ao-tag-swap')) return;
            state.labelArmed = null;
            render();
        }, true);
        // Menu degli operatori (Ticket): un clic fuori lo chiude.
        root.addEventListener('click', event => {
            if (!state.filterMenu || event.target.closest('.ao-menu, .ao-filter-dd')) return;
            state.filterMenu = null;
            render();
        }, true);
        // Anteprima aperta in In coda: un clic fuori dalla scheda prima la richiude (con l'animazione)
        // e non fa altro; solo il clic successivo apre un'altra richiesta. Schede e selettori restano liberi.
        root.addEventListener('click', event => {
            if (!state.queueExpanded || state.holdRender) return;
            const target = event.target;
            if (target.closest('.ao-queue-item.ao-expanded, .ao-select-menu, .ao-tabs, .ao-fixed, .ao-status, .ao-statusbar')) return;
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
    }

    // Sottolineatura blu delle schede principali: un elemento solo che scivola sulla scheda attiva
    // (come la pillola dei selettori). In Impostazioni nessuna scheda è attiva e sparisce.
    // Barra delle schede come continuazione dell'intestazione blu dell'add-in: riprende la stessa sfumatura
    // (letta dall'intestazione) e la prolunga, così le due fasce sembrano una sola.
    function blendTabsWithHeader() {
        const nav = root?.querySelector('.ao-tabs');
        const header = document.querySelector('.app-header');
        if (!nav || !header || !nav.offsetWidth) return;
        const W = header.offsetWidth, H = header.offsetHeight, w = nav.offsetWidth, n = nav.offsetHeight;
        // Vetro: intestazione e schede sono un'unica lastra. Sfumatura e riflesso si dipingono sull'altezza di tutte e due
        // e la barra delle schede ne mostra la parte bassa, così l'effetto continua senza giunture.
        if (root.dataset.style === 'glass') {
            const whole = { backgroundSize: `${W}px ${H + n}px`, backgroundRepeat: 'no-repeat' };
            Object.assign(header.style, whole);
            nav.style.background = '';
            Object.assign(nav.style, whole, { backgroundImage: getComputedStyle(header).backgroundImage, backgroundColor: 'transparent', backgroundPosition: `0 ${-H}px` });
            return;
        }
        Object.assign(header.style, { backgroundSize: '', backgroundRepeat: '' });
        // Flat: la barra riprende la sfumatura dell'intestazione dal punto in cui finisce e la prolunga.
        const colors = (getComputedStyle(header).backgroundImage.match(/rgba?\([^)]+\)/g) || [])
            .map(color => { const [r, g, b, a = 1] = color.match(/[\d.]+/g).map(Number); return [r, g, b, a]; });
        if (colors.length < 2) { nav.style.background = getComputedStyle(header).backgroundColor; return; }
        const [from, to] = [colors[0], colors[colors.length - 1]];
        const start = H / (W + H); // punto della sfumatura sul bordo basso dell'intestazione, a sinistra
        const mix = from.map((value, i) => value + (to[i] - value) * start);
        const css = ([r, g, b, a]) => `rgba(${Math.round(r)}, ${Math.round(g)}, ${Math.round(b)}, ${+a.toFixed(3)})`;
        const end = Math.min(100, (W / (w + n)) * 100);
        nav.style.background = `linear-gradient(135deg, ${css(mix)} 0%, ${css(to)} ${end.toFixed(1)}%)`;
    }
    function placeTabIndicator() {
        const indicator = root?.querySelector('.ao-tab-indicator');
        if (!indicator) return;
        const activeView = state.view === 'mail' ? (state.backTo || 'tickets') : state.view;
        const tab = root.querySelector(`.ao-tab[data-view="${activeView}"]`);
        if (!tab) { indicator.classList.remove('ao-ready'); return; }
        const changed = indicator.dataset.view && indicator.dataset.view !== activeView && indicator.classList.contains('ao-ready');
        indicator.classList.toggle('ao-moving', Boolean(changed));
        // Linea corta e arrotondata, centrata sotto icona e nome della scheda.
        const content = [...tab.children].reduce((sum, part) => sum + (part.classList.contains('ao-counter-off') ? 0 : part.scrollWidth), 0) + 10; // a scheda aperta
        const width = Math.min(tab.offsetWidth, Math.max(36, content + 12));
        indicator.style.width = `${width}px`;
        indicator.style.transform = `translateX(${tab.offsetLeft + (tab.offsetWidth - width) / 2}px)`;
        indicator.dataset.view = activeView;
        requestAnimationFrame(() => indicator.classList.add('ao-ready'));
    }
    // Segnalino (numero) accanto al nome della scheda: solo se le schede sono più di una.
    function render() {
        if (!root) return;
        if (state.holdRender) { renderPending = true; return; }
        closeSelects();
        const config = loadConfig();
        const user = currentUser(config);
        // Thread dei miei ticket con messaggi nuovi (uno per thread).
        const newReplies = new Set(state.tickets.filter(ticket => ticket.newMessages && isMine(ticket, user)).map(ticket => ticket.conversationId)).size;
        const counts = { queue: state.queue.length, tickets: state.tickets.filter(ticket => isMine(ticket, user)).length };
        placeTabIndicator();
        root.querySelectorAll('.ao-tab').forEach(tab => {
            // Nella vista di lavoro resta evidenziata la scheda da cui si è arrivati.
            const active = tab.dataset.view === (state.view === 'mail' ? (state.backTo || 'tickets') : state.view);
            tab.classList.toggle('ao-active', active);
            tab.setAttribute('aria-selected', String(active));
            tab.title = active ? '' : TEXTS.tabs[tab.dataset.view];
            // Contatore solo sulla scheda aperta; gli aggiornamenti (Ticket: thread dei miei ticket con messaggi nuovi)
            // sono un segnalino sull'icona, sempre visibile.
            const count = counts[tab.dataset.view];
            const tabCounter = tab.querySelector('.ao-counter');
            tabCounter.textContent = count || '';
            tabCounter.classList.toggle('ao-counter-off', !count || !active);
            const updates = tab.dataset.view === 'tickets' ? newReplies : 0;
            const badge = tab.querySelector('.ao-notify');
            badge.firstElementChild.textContent = updates || '';
            badge.classList.toggle('ao-notify-off', !updates);
            if (updates) badge.title = TEXTS.tickets.newMessagesTitle(updates);
            else badge.removeAttribute('title');
        });
        renderStatusBar();
        const views = { queue: renderQueue, tickets: renderTickets, mail: renderMail, settings: renderSettings };
        const body = root.querySelector('.ao-body');
        // Chi resta fermo (barra secondaria, "←") sta in alto, separato da una linea; tutto il resto scorre
        // sotto, a tutta larghezza. La posizione di scorrimento resta uguale tra un aggiornamento e l'altro.
        const scrollKey = `${state.view}|${state.target?.kind || ''}|${state.target?.number || state.target?.mail?.conversationId || ''}|${state.queueTab}|${state.ticketSource}|${state.ticketOperator}|${state.ticketSearchOpen}|${state.settingsTab}`;
        const previousScroll = body.dataset.scrollKey === scrollKey ? body.querySelector('.ao-scroll')?.scrollTop || 0 : 0;
        const view = views[state.view](config, user);
        const children = [...view.children];
        const fixedCount = children.findIndex(child => !child.matches('.ao-segmented, .ao-back-row, .ao-source-bar'));
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
    // Si ricorda anche dove stava la pillola: ridisegnando con la stessa scelta (es. un clic su una mail in Outlook)
    // compare subito al suo posto, senza un istante di vuoto.
    const segmentMemory = new Map();
    const segmentPlace = new Map();
    const segmentKey = options => options.map(option => option.value).join('|');
    // Mentre la barra cambia forma (filtri, ricerca) la pillola si mette al suo posto solo alla fine del movimento.
    let barSettled = Promise.resolve();
    function settleBar() {
        let done;
        barSettled = new Promise(resolve => { done = resolve; });
        return done;
    }
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
    // key: memoria del selettore quando le stesse voci compaiono in più schede (azioni di In coda, di un ticket);
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
                label: [option.notify ? h('span', { class: 'ao-notify-anchor' }, svgIcon(option.icon), notifyBadge(option.notify, option.notifyTitle)) : svgIcon(option.icon), active ? h('span', { class: entering ? 'ao-segment-name ao-segment-name-in' : 'ao-segment-name' }, h('span', null, option.label)) : leaving(option, active)],
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
    // "←" della vista di lavoro: torna all'elenco da cui si è arrivati (di base Ticket).
    function backRow() {
        const target = state.backTo || 'tickets';
        return h('div', { class: 'ao-back-row' },
            h('button', { type: 'button', class: 'ao-link ao-back-link', onclick: () => closeToList(target) }, TEXTS.common.back(TEXTS.tabs[target])));
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
            // Stessa riga di Ticket: selettore della stessa larghezza, senza pulsanti ai lati.
            h('div', { class: 'ao-source-bar' }, h('div', { class: 'ao-bar-row' },
                h('span', { class: 'ao-bar-spacer', 'aria-hidden': 'true' }),
                iconBar([
                    { value: 'active', icon: SEGMENT_ICONS.active, label: TEXTS.queue.active, count: active.length },
                    { value: 'ignored', icon: SEGMENT_ICONS.ignored, label: TEXTS.queue.ignoredTab, count: state.ignored.length }
                ], showIgnored ? 'ignored' : 'active', value => { state.queueTab = value; state.queueExpanded = null; render(); }),
                h('span', { class: 'ao-bar-spacer', 'aria-hidden': 'true' }))),
            showIgnored
                ? [
                    state.ignored.length ? h('ul', { class: 'ao-list' }, state.ignored.map(item => ignoredItem(item, config)))
                        : (servicesBlocked() ? null : h('p', { class: 'ao-empty' }, TEXTS.queue.ignoredEmpty)),
                    // Spiegazione in fondo, come nota.
                    state.ignored.length ? h('p', { class: 'ao-footnote' }, TEXTS.queue.ignoredHint) : null
                ]
                : (active.length ? h('ul', { class: 'ao-list' }, active.map(item => queueItem(item, config))) : (servicesBlocked() ? null : h('p', { class: 'ao-empty' }, TEXTS.queue.empty))));
    }
    // Thread ignorato: si può prendere in carico direttamente (stesso modulo compatto della coda).
    // Scheda di In coda (Attive e Ignorate). Chiusa: oggetto, mittente, ora, etichette; il clic la apre.
    // Aperta: testo della mail (in coda non si vede altrove), agenzia e attività, Prendi in carico e Ignora.
    // Dopo la presa in carico si resta in In coda; l'avviso offre "Vai al ticket".
    function ignoredItem(item, config) {
        // Chi l'ha ignorata si legge solo aprendo la scheda, sopra il testo della mail.
        return queueCard(item, config, labelSlot(item, config), { allowIgnore: false,
            info: TEXTS.queue.ignoredBy(item.discardedBy || TEXTS.common.none, formatWhen(item.discardedAt)) });
    }
    function queueCard(item, config, chipNode, { allowIgnore = true, info = null } = {}) {
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
                    h('span', { class: 'ao-item-sub' }, item.sender?.name || item.sender?.email || ''),
                    chipNode)),
            expanded ? queueDetails(item, config, allowIgnore, info) : null);
    }
    // Scheda aperta: testo della mail, poi il selettore dell'azione (nessuna scelta all'inizio: si decide dopo aver letto).
    // Sotto compare il contenuto dell'azione scelta, con il suo pulsante; in fondo la freccia richiude la scheda.
    //  - Prendi in carico: agenzia e attività;  - Inoltra: gruppo (ES, PEL, PO);  - Ignora: "Confermi?".
    function queueDetails(item, config, allowIgnore, info = null) {
        const mail = mailFromQueueItem(item);
        const body = state.mailBodies[item.messageId];
        const choice = state.queueChoice[item.conversationId] || {};
        // Si gestisce solo con la propria etichetta: senza, o con quella di un collega, la mail si legge soltanto.
        const locked = !isMyLabel(item.label);
        const actions = [
            { value: 'take', label: TEXTS.mail.openTicket, tone: 'accent', disabled: locked },
            { value: 'forward', label: TEXTS.queue.forward, tone: 'forward', disabled: locked },
            allowIgnore ? { value: 'ignore', label: TEXTS.mail.discard, tone: 'danger', disabled: locked } : null
        ].filter(Boolean);
        return h('div', { class: 'ao-inline-form' },
            info ? h('p', { class: 'ao-hint ao-queue-info' }, info) : null,
            // Testo limitato a PREVIEW_MAX_CHARS; la barra di scorrimento compare solo a scheda aperta del tutto.
            h('div', { class: `ao-mail-preview${state.previewScroll === item.conversationId ? ' ao-scrollable' : ''}` },
                clipText(body ?? item.preview ?? '', PREVIEW_MAX_CHARS), body === undefined ? h('span', { class: 'ao-hint' }, ` ${TEXTS.common.loading}`) : null),
            iconBar(actions.map(action => ({ ...action, icon: ACTION_ICONS[action.value] })), choice.action, value => pickQueueAction(item, value), { key: `action:${item.conversationId}`, tones: true, slide: false }),
            choiceZone(`queue:${item.conversationId}`, choice.action ? queueActionContent(item, mail, config, choice) : null),
            h('button', {
                type: 'button', class: 'ao-collapse', 'aria-label': TEXTS.common.close, title: TEXTS.common.close,
                onclick: event => closeTakeCharge(event.currentTarget.closest('.ao-inline-form'))
            }, svgIcon(CHEVRON_UP_ICON)));
    }
    function queueActionContent(item, mail, config, choice) {
        const card = button => button.closest('.ao-queue-item');
        if (choice.action === 'take') return [ticketForm(mail, config, { compact: true, zone: true }), h('div', { class: 'ao-actions' }, openTicketButton(mail, config))];
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
    function mailFromQueueItem(item) {
        return {
            conversationId: item.conversationId, itemId: item.messageId, subject: item.subject, sender: item.sender,
            receivedAt: item.receivedAt, body: state.mailBodies[item.messageId] || item.preview, incNumbers: []
        };
    }
    // Etichetta della scheda: chi sta gestendo la mail. Vuota c'è il pulsantino per mettere la propria;
    // la propria si toglie con un clic; quella di un collega si vede soltanto (e la mail non si può gestire).
    function queueItem(item, config) {
        return queueCard(item, config, labelSlot(item, config));
    }
    const surname = name => String(name || '').trim().split(/\s+/).pop() || '';
    const isMyLabel = label => Boolean(label) && label.name.trim().toLowerCase() === currentUser().displayValue.trim().toLowerCase();
    function labelChip(label, { onclick, title, entering } = {}) {
        const color = THEME.labels[label.color] || THEME.labels.grafite;
        return h(onclick ? 'button' : 'span', {
            type: onclick ? 'button' : undefined, class: `ao-user-tag${entering ? ' ao-user-tag-in' : ''}`, style: `--ao-tag-color: ${color}`,
            title: title || label.name, onclick
        }, surname(label.name));
    }
    // Mail di un collega, a scheda aperta: un clic sulla sua etichetta la prepara al passaggio (il suo cognome barrato,
    // poi il proprio); un secondo clic la rende propria. Un clic altrove annulla.
    function takeoverTag(item, config) {
        const armed = state.labelArmed === item.conversationId;
        const myColor = THEME.labels[config.profile.label] || THEME.labels.azzurro;
        const color = THEME.labels[item.label.color] || THEME.labels.grafite;
        return h('button', {
            type: 'button', class: `ao-user-tag ao-tag-swap${armed ? ' ao-armed-tag' : ''}`, style: `--ao-tag-color: ${color}; --ao-tag-new: ${myColor}`,
            title: armed ? TEXTS.queue.labelTakeConfirm : TEXTS.queue.labelTake(item.label.name),
            onclick: event => {
                event.stopPropagation();
                if (!armed) { state.labelArmed = item.conversationId; render(); return; }
                state.labelArmed = null;
                setLabel(item, true, { force: true });
            }
        }, armed ? [h('s', {}, surname(item.label.name)), svgIcon(SWAP_ICON), h('span', { class: 'ao-tag-next' }, surname(currentUser(config).displayValue))] : surname(item.label.name));
    }
    function labelSlot(item, config) {
        const stop = handler => event => { event.stopPropagation(); handler(); };
        if (item.label && !isMyLabel(item.label) && state.queueExpanded === item.conversationId && currentUser(config).displayValue) return takeoverTag(item, config);
        if (item.label) {
            const mine = isMyLabel(item.label);
            return labelChip(mine ? { ...item.label, color: config.profile.label } : item.label, {
                entering: state.labelEntering === item.conversationId,
                title: mine ? TEXTS.queue.labelRemove : item.label.name,
                onclick: mine ? stop(() => setLabel(item, false)) : undefined
            });
        }
        if (!currentUser(config).displayValue) return null;
        return h('button', { type: 'button', class: 'ao-user-tag-add', title: TEXTS.queue.labelAdd, 'aria-label': TEXTS.queue.labelAdd, onclick: stop(() => setLabel(item, true)) }, svgIcon(TAG_ADD_ICON));
    }
    // Subito a schermo (l'etichetta compare con un piccolo rimbalzo), poi sul registro condiviso:
    // se nel frattempo l'ha presa un collega, torna la sua e lo si dice.
    async function setLabel(item, on, { force = false } = {}) {
        const config = loadConfig();
        const user = currentUser(config);
        const before = item.label;
        item.label = on ? { name: user.displayValue, color: config.profile.label } : null;
        state.labelEntering = on ? item.conversationId : null;
        render();
        state.labelEntering = null;
        try {
            if (on) {
                const result = await Services.labelMail(item.conversationId, user, config.profile.label, { force });
                if (!result?.ok) {
                    item.label = result?.label || before;
                    render();
                    showStatus(TEXTS.queue.labeledBy(result?.label?.name || ''), 'warning');
                }
            } else await Services.unlabelMail(item.conversationId, user);
        } catch (error) {
            console.error('[AMP]', error);
            item.label = before;
            render();
            showStatus(error instanceof ServicesMissingError ? TEXTS.common.servicesMissing : (error?.message || String(error)), 'danger');
        }
    }

    // ---------------------------- Ticket ----------------------------------
    // I ticket dello stesso thread stanno insieme; in cima i thread con messaggi nuovi.
    // Ticket: in alto il selettore Da mail | Senza mail (centrato). Ricliccando la voce scelta si apre il menu
    // degli operatori: Tutti (vista generale), Miei, poi un collega alla volta. Il filtro scelto si legge nella voce.
    // "+ Nuovo ticket" apre un ticket senza mail, quindi sta solo in Senza mail.
    function renderTickets(config, user) {
        const notice = serviceNotice();
        const open = state.tickets.slice();
        const operator = state.ticketOperator;
        const byOperator = ticket => operator === 'all' || (operator === 'mine' ? isMine(ticket, user)
            : !isMine(ticket, user) && (ticket.assignedTo?.name || TEXTS.common.unassigned) === operator);
        const fromMail = ticket => Boolean(ticket.messageId);
        // Ricerca (lente) "super partes": numero INC o oggetto su tutti i ticket aperti, di qualunque sezione e operatore.
        // Mentre si cerca, sezione e filtro non contano (e non si vedono); chiudendo la ricerca si torna dove si era.
        // Finché non si scrive niente non compare nessun risultato; mentre si scrive, un attimo di caricamento e poi i risultati.
        const searching = state.ticketSearchOpen;
        const query = searching ? state.ticketSearchShown : '';
        const matches = ticket => `${ticket.number} ${ticket.shortDescription || ''}`.toLowerCase().includes(query);
        const byStatus = ticket => state.ticketStatus === 'all' || (state.ticketStatus === 'hold' ? ticket.state === '3' : ticket.state !== '3');
        const visible = open.filter(ticket => byOperator(ticket) && byStatus(ticket));
        const items = searching ? (query && !state.ticketSearchPending ? open.filter(matches) : [])
            : visible.filter(ticket => (state.ticketSource === 'none' ? !fromMail(ticket) : fromMail(ticket)));
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
            ticketSourceBar(open, user, {
                mail: visible.filter(fromMail).length,
                mailUpdates: new Set(visible.filter(ticket => fromMail(ticket) && ticket.newMessages).map(ticket => ticket.conversationId)).size,
                none: visible.filter(ticket => !fromMail(ticket)).length
            }),
            state.ticketSource === 'none' && !searching ? newTicketSection(config) : null,
            searching && state.ticketSearchPending ? searchSkeleton() : null,
            groups.length ? h('ul', { class: `ao-list${searching ? ' ao-results-in' : ''}` }, groups.map(group => (group.tickets.length > 1
                ? h('li', { class: `ao-group${group.tickets.some(ticket => ticket.newMessages) ? ' ao-group-new' : ''}` },
                    h('div', { class: 'ao-group-head' }, TEXTS.tickets.threadGroup(group.tickets.length)),
                    newMessagesPreview(group.tickets[0]),
                    h('ul', { class: 'ao-list' }, group.tickets.map(ticket => ticketItem(ticket, user, { inGroup: true, query }))))
                : ticketItem(group.tickets[0], user, { query }))))
                : (servicesBlocked() || (searching && (!query || state.ticketSearchPending)) ? null
                    : h('p', { class: 'ao-empty' }, searching ? TEXTS.tickets.noResults : TEXTS.tickets.empty)));
    }
    // Barra di Ticket, simmetrica: imbuto (filtri) | Da mail | Senza mail | lente (ricerca).
    // L'imbuto trasforma il selettore nei filtri Stato (Tutti, Aperti, In attesa) e Operatore (Miei, i colleghi, Tutti);
    // valgono per entrambi gli elenchi. Ricliccandolo torna il selettore, con un puntino se un filtro è attivo.
    // La lente trasforma la barra nel campo di ricerca (stessa misura) e resta al suo posto, a destra:
    // apre e chiude la ricerca (come un interruttore). Si cerca su tutti i ticket, mentre si scrive.
    function ticketSourceBar(open, user, counts) {
        const operator = state.ticketOperator;
        const colleagues = [...new Set(open.filter(ticket => !isMine(ticket, user)).map(ticket => ticket.assignedTo?.name || TEXTS.common.unassigned))]
            .sort((a, b) => a.localeCompare(b, 'it'));
        const filtered = operator !== 'mine' || state.ticketStatus !== 'all';
        // Imbuto: apre e chiude i filtri al posto di Da mail | Senza mail (il puntino dice che un filtro è attivo).
        const filterButton = h('button', {
            type: 'button', class: `ao-bar-btn ao-bar-filter${filtered ? ' ao-active' : ''}${state.ticketFilterOpen ? ' ao-open' : ''}`,
            'aria-pressed': String(state.ticketFilterOpen), 'aria-label': TEXTS.tickets.filters, title: TEXTS.tickets.filters,
            onclick: () => morphBar(() => { state.ticketFilterOpen = !state.ticketFilterOpen; state.filterMenu = null; })
        }, svgIcon(state.ticketFilterOpen ? FILTER_OFF_ICON : FILTER_ICON));
        // Con i filtri aperti la lente è disattivata (e con la ricerca aperta l'imbuto): la barra ha una funzione alla volta.
        // Il riquadro resta sempre della misura di Da mail | Senza mail: il pulsante che non serve resta al suo posto, disattivato.
        const searchButton = h('button', {
            type: 'button', class: `ao-bar-btn ao-bar-search${state.ticketSearchOpen ? ' ao-active' : ''}`,
            'aria-label': TEXTS.tickets.search, title: TEXTS.tickets.search, disabled: state.ticketFilterOpen,
            onclick: state.ticketSearchOpen ? closeTicketSearch : openTicketSearch
        }, svgIcon(state.ticketSearchOpen ? CLOSE_ICON : SEARCH_ICON));
        if (state.ticketSearchOpen) {
            filterButton.disabled = true;
            return h('div', { class: 'ao-source-bar' }, h('div', { class: 'ao-bar-row' }, filterButton,
                h('div', { class: 'ao-segmented ao-search-bar', role: 'search' },
                    h('input', {
                        class: 'ao-search-input', type: 'search', value: state.ticketQuery, placeholder: TEXTS.tickets.searchPlaceholder, 'aria-label': TEXTS.tickets.search,
                        oninput: event => runTicketSearch(event.target.value)
                    })),
                searchButton));
        }
        const middle = state.ticketFilterOpen ? h('div', { class: 'ao-segmented ao-filter-bar' },
            filterField('status', TEXTS.tickets.status, state.ticketStatus, 'all',
                Object.entries(TEXTS.tickets.statuses).map(([value, label]) => ({ value, label })), value => { state.ticketStatus = value; }),
            filterField('operator', TEXTS.tickets.operator, operator, 'mine',
                [{ value: 'mine', label: TEXTS.tickets.mine }, ...colleagues.map(name => ({ value: name, label: name })), { value: 'all', label: TEXTS.tickets.allOperators, separated: true }],
                value => { state.ticketOperator = value; }))
            : iconBar([
                { value: 'mail', icon: SEGMENT_ICONS.mail, label: TEXTS.tickets.fromMail, count: counts.mail, notify: counts.mailUpdates, notifyTitle: TEXTS.tickets.newMessagesTitle(counts.mailUpdates) },
                { value: 'none', icon: SEGMENT_ICONS.none, label: TEXTS.tickets.withoutMail, count: counts.none }
            ], state.ticketSource, value => {
                if (value === state.ticketSource) return;
                state.ticketSource = value;
                render();
            });
        return h('div', { class: 'ao-source-bar' }, h('div', { class: 'ao-bar-row' }, filterButton, middle, searchButton));
    }
    // Un filtro nella barra: vuoto mostra il suo nome (vale la voce predefinita, segnata nell'elenco);
    // scelto mostra il valore con la × per toglierlo. La scelta si applica subito.
    function filterField(key, name, value, fallback, options, apply) {
        const set = value !== fallback;
        const pick = next => { apply(next); state.filterMenu = null; render(); };
        return h('div', {
            class: `ao-filter-dd${set ? ' ao-set' : ''}`, role: 'button', tabindex: '0', 'aria-haspopup': 'listbox', 'aria-expanded': String(state.filterMenu === key),
            onclick: () => { state.filterMenu = state.filterMenu === key ? null : key; render(); },
            onkeydown: event => { if (event.key === 'Enter') { state.filterMenu = state.filterMenu === key ? null : key; render(); } }
        },
            h('span', { class: 'ao-filter-value' }, set ? options.find(option => option.value === value)?.label || value : name),
            set ? h('button', { type: 'button', class: 'ao-chip-clear', 'aria-label': TEXTS.tickets.clearFilter, title: TEXTS.tickets.clearFilter, onclick: event => { event.stopPropagation(); pick(fallback); } }, '×')
                : h('span', { class: 'ao-filter-caret', 'aria-hidden': 'true' }, '▾'),
            state.filterMenu === key ? h('div', { class: `ao-menu ao-filter-menu ao-filter-menu-${key}`, role: 'listbox', onclick: event => event.stopPropagation() },
                options.map(option => [
                    option.separated ? h('hr', { class: 'ao-menu-sep' }) : null,
                    h('button', { type: 'button', role: 'option', 'aria-selected': String(option.value === value), class: `ao-menu-item${option.value === value ? ' ao-active' : ''}`, onclick: () => pick(option.value) },
                        h('span', {}, option.label), option.value === fallback ? h('span', { class: 'ao-menu-default' }, TEXTS.tickets.defaultFilter) : null)
                ])) : null);
    }
    // Cambio della barra secondaria (filtri, ricerca): dura BAR_TIME, in due metà uguali. Le scritte svaniscono, poi
    // compaiono quelle nuove; nella seconda metà le due parti dei filtri prendono la loro larghezza e il pulsante
    // che si spegne (o si riaccende) cambia, così tutto finisce insieme.
    const BAR_TIME = 500;
    const BAR_EASE = 'cubic-bezier(.3, .8, .2, 1)';
    const BAR_FADE_OUT = { duration: BAR_TIME / 2, easing: 'linear', fill: 'forwards' };
    const BAR_FADE_IN = { duration: BAR_TIME / 2, easing: 'linear' };
    const BAR_MOTION = { duration: BAR_TIME / 2, easing: BAR_EASE };
    const fadeAll = (elements, from, to, timing) => Promise.all(elements.filter(Boolean)
        .map(element => element.animate([{ opacity: from }, { opacity: to }], timing).finished)).catch(() => {});
    // Le scritte nuove compaiono mentre il riquadro prende la sua forma.
    async function fadeInWith(elements, move = async () => {}) {
        const contents = elements.filter(Boolean);
        contents.forEach(element => { element.style.opacity = '0'; });
        await Promise.all([move(), fadeAll(contents, 0, 1, BAR_FADE_IN)]);
        contents.forEach(element => { element.style.opacity = ''; });
    }
    const barRow = () => root.querySelector('.ao-source-bar .ao-bar-row');
    const barMiddle = () => root.querySelector('.ao-source-bar .ao-bar-row > .ao-segmented');
    // L'elenco sotto la barra svanisce e ricompare con lei (la barra resta).
    const listBelow = () => [...(root.querySelector('.ao-scroll')?.children || [])].filter(child => !child.classList.contains('ao-fixed'));
    // Filtri (imbuto): le due voci del selettore (Da mail | Senza mail) si trasformano nei due filtri (Stato | Operatore)
    // e viceversa. Il riquadro resta fermo: le scritte svaniscono, le due metà prendono la nuova larghezza, poi
    // compaiono le scritte nuove (e, tornando al selettore, la pillola della voce scelta).
    async function morphBar(change) {
        if (state.holdRender) return;
        const partsOf = bar => [...(bar?.children || [])].filter(child => child.matches('.ao-segment, .ao-filter-dd'));
        const old = barMiddle();
        if (!old || reducedMotion()) { change(); render(); return; }
        const widths = partsOf(old).map(part => part.offsetWidth);
        // Al ritorno il filtro impostato diventa la pillola della voce scelta (l'inverso dell'andata)
        const lit = old.querySelector('.ao-filter-dd.ao-set');
        const litFrom = lit && { width: `${lit.offsetWidth}px`, transform: `translateX(${lit.offsetLeft - 2}px)` }; // 2px = bordo interno
        // Anche l'icona dell'imbuto cambia (in dissolvenza); la lente si spegne (aprendo) o si riaccende (chiudendo).
        const funnelIcon = () => barRow()?.querySelector('.ao-bar-filter > *');
        const lensWasOff = barRow()?.querySelector('.ao-bar-search')?.disabled;
        await holdingRender(() => fadeAll([...partsOf(old).flatMap(part => [...part.children]), ...old.querySelectorAll('.ao-segment-thumb'), funnelIcon()], 1, 0, BAR_FADE_OUT));
        const settle = settleBar();
        change();
        render();
        easeDisabled('.ao-bar-search', lensWasOff);
        const fresh = barMiddle();
        const parts = partsOf(fresh);
        if (parts.length !== widths.length) return settle();
        const thumb = fresh.querySelector('.ao-segment-thumb');
        const active = fresh.querySelector('.ao-segment.ao-active');
        const pill = litFrom && thumb && active ? thumb : null;
        const contents = [...parts.flatMap(part => [...part.children]), thumb, funnelIcon()].filter(element => element && element !== pill);
        await holdingRender(() => fadeInWith(contents, async () => {
            const targets = parts.map(part => part.offsetWidth); // tutte prima di muovere: ogni metà cambia le altre
            const pillTo = pill && { width: `${active.offsetWidth}px`, transform: `translateX(${active.offsetLeft - 2}px)` };
            // Si muovono le prime metà; l'ultima riempie lo spazio che resta, così il bordo destro non si sposta.
            const moves = parts.map((part, index) => part.animate(index < parts.length - 1
                ? [{ flex: `0 0 ${widths[index]}px` }, { flex: `0 0 ${targets[index]}px` }]
                : [{ flex: '1 1 0px', minWidth: '0px' }, { flex: '1 1 0px', minWidth: '0px' }], BAR_MOTION).finished);
            if (pill) {
                Object.assign(pill.style, pillTo);
                pill.classList.remove('ao-moving');
                pill.classList.add('ao-ready');
                moves.push(pill.animate([litFrom, pillTo], BAR_MOTION).finished);
            }
            await Promise.all(moves).catch(() => {});
            settle();
        }));
    }
    // Ricerca (lente): come i filtri. Le scritte e l'elenco svaniscono e, nello stesso riquadro (che non cambia misura),
    // compare il campo di ricerca; l'imbuto si spegne. Chiudendo ricompare tutto.
    // La lente diventa una × mentre la ricerca è aperta: l'icona cambia in dissolvenza con il resto.
    const lensIcon = () => barRow()?.querySelector('.ao-bar-search > *');
    // Il pulsante che si spegne (o si riaccende) non sparisce: passa da acceso a spento con il movimento della barra.
    function easeDisabled(selector, wasOff) {
        const button = barRow()?.querySelector(selector);
        if (!button || button.disabled === Boolean(wasOff)) return;
        const on = { opacity: 1 }, off = { opacity: THEME.state.disabledOpacity };
        button.animate(wasOff ? [off, on] : [on, off], BAR_MOTION);
    }
    async function openTicketSearch() {
        if (state.holdRender) return;
        state.filterMenu = null;
        const open = () => Object.assign(state, { ticketSearchOpen: true, ticketQuery: '', ticketSearchShown: '', ticketSearchPending: false });
        const box = barMiddle();
        if (!box || reducedMotion()) { open(); render(); focusTicketSearch(); return; }
        await holdingRender(() => fadeAll([...box.children, lensIcon(), ...listBelow()], 1, 0, BAR_FADE_OUT));
        open();
        render();
        easeDisabled('.ao-bar-filter', false);
        focusTicketSearch();
        await holdingRender(() => fadeInWith([...barMiddle().children, lensIcon()]));
    }
    async function closeTicketSearch() {
        if (state.holdRender) return;
        clearTimeout(searchTimer);
        const close = () => Object.assign(state, { ticketSearchOpen: false, ticketQuery: '', ticketSearchShown: '', ticketSearchPending: false });
        const bar = barMiddle();
        if (!bar || reducedMotion()) { close(); render(); return; }
        await holdingRender(() => fadeAll([...bar.children, lensIcon(), ...listBelow()], 1, 0, BAR_FADE_OUT));
        const settle = settleBar();
        close();
        render();
        easeDisabled('.ao-bar-filter', true);
        await holdingRender(() => fadeInWith([...barMiddle().children, lensIcon(), ...listBelow()], async () => settle()));
    }
    // Ricerca mentre si scrive: subito lo scheletro di caricamento, i risultati quando ci si ferma un attimo.
    let searchTimer = null;
    const SEARCH_DELAY = 350;
    function runTicketSearch(value) {
        state.ticketQuery = value;
        clearTimeout(searchTimer);
        const query = value.trim().toLowerCase();
        if (!query) {
            Object.assign(state, { ticketSearchShown: '', ticketSearchPending: false });
        } else {
            state.ticketSearchPending = true;
            searchTimer = setTimeout(() => {
                Object.assign(state, { ticketSearchShown: query, ticketSearchPending: false });
                render();
                focusTicketSearch();
            }, SEARCH_DELAY);
        }
        render();
        focusTicketSearch();
    }
    function searchSkeleton() {
        return h('div', { class: 'ao-skeleton-list', 'aria-busy': 'true', 'aria-label': TEXTS.common.loading },
            [0, 1, 2].map(() => h('div', { class: 'ao-skeleton' },
                h('span', { class: 'ao-skeleton-line ao-skeleton-short' }),
                h('span', { class: 'ao-skeleton-line' }),
                h('span', { class: 'ao-skeleton-line ao-skeleton-mid' }))));
    }
    // Dopo ogni ridisegno il campo di ricerca torna attivo, con il cursore in fondo.
    function focusTicketSearch() {
        const input = root?.querySelector('.ao-search-input');
        if (!input) return;
        input.focus();
        input.setSelectionRange(input.value.length, input.value.length);
    }
    // Testo con le parti trovate dalla ricerca evidenziate.
    function markText(text, query) {
        const value = String(text || '');
        if (!query) return value;
        const parts = [];
        let from = 0;
        const lower = value.toLowerCase();
        for (let at = lower.indexOf(query); at >= 0; at = lower.indexOf(query, from)) {
            parts.push(value.slice(from, at), h('mark', { class: 'ao-mark' }, value.slice(at, at + query.length)));
            from = at + query.length;
        }
        parts.push(value.slice(from));
        return parts;
    }
    // Nuovo ticket senza mail: si apre da Ticket. Stesso modulo della presa in carico, più Short description,
    // data e ora di ricezione (hh:mm, poi nel formato di ServiceNow) e Description. Caller = chi apre il ticket.
    const NEW_TICKET_KEY = '__new__';
    function manualMail() {
        return { conversationId: NEW_TICKET_KEY, manual: true, subject: '', body: '', sender: null, receivedAt: null, incNumbers: [] };
    }
    function newTicketSection(config) {
        if (!state.newTicketOpen) {
            return h('button', { type: 'button', class: 'ao-new-ticket', onclick: () => openNewTicket() }, svgIcon(SEGMENT_ICONS.none), TEXTS.tickets.newTicket);
        }
        const mail = manualMail();
        const draft = ticketDraft(mail, config);
        const today = new Date();
        draft.date ||= `${today.getFullYear()}-${pad(today.getMonth() + 1)}-${pad(today.getDate())}`;
        draft.time ??= `${pad(today.getHours())}:${pad(today.getMinutes())}`;
        draft.description ??= '';
        const ready = () => Boolean(config.companies[draft.companyId] && config.activities[draft.activityId] && String(draft.subject || '').trim());
        const refreshButton = () => { const button = root.querySelector('.ao-new-ticket-card .ao-primary'); if (button) button.disabled = !ready(); };
        // Scrivendo la Short description si aggiornano solo i suggerimenti (il campo non perde il cursore).
        const refreshSuggestions = () => {
            const slot = root.querySelector(`[data-suggest-for="${NEW_TICKET_KEY}"]`);
            if (slot) slot.replaceChildren(...[suggestionPills(draft, config, draft.subject)].filter(Boolean));
        };
        const zone = ticketForm(mail, config, { zone: true });
        zone.append(
            h('div', { class: 'ao-field-row' },
                textField(TEXTS.tickets.receivedDate, draft.date, value => { draft.date = value; }, { type: 'date' }),
                textField(TEXTS.tickets.receivedTime, draft.time, value => { draft.time = value; }, { placeholder: 'hh:mm', maxlength: '5', inputmode: 'numeric' })),
            textField(TEXTS.tickets.shortDescription, draft.subject, value => { draft.subject = value; refreshSuggestions(); refreshButton(); }, { maxlength: '160' }),
            textField(TEXTS.tickets.description, draft.description, value => { draft.description = value; }, { multiline: true, rows: '4' }));
        return h('section', { class: 'ao-card ao-new-ticket-card' },
            // Stesso riquadro della presa in carico da In coda: Agenzia, Attività, data e ora di ricezione,
            // poi Short description e Description.
            h('div', { class: 'ao-inline-form' },
                zone,
                h('div', { class: 'ao-actions' },
                    h('button', { type: 'button', class: 'ao-button ao-primary', disabled: !ready(), onclick: () => openTicket(manualMail()) }, TEXTS.tickets.openNew),
                    h('span', { class: 'ao-spacer' }),
                    h('button', { type: 'button', class: 'ao-button ao-subtle', onclick: event => closeNewTicket(event.currentTarget.closest('.ao-inline-form')) }, TEXTS.common.cancel))));
    }
    async function openNewTicket() {
        await whenIdle();
        state.newTicketOpen = true;
        render();
        await holdingRender(() => unfold(root.querySelector('.ao-new-ticket-card .ao-inline-form')));
    }
    async function closeNewTicket(form) {
        holdRender();
        await fold(form);
        state.newTicketOpen = false;
        delete state.ticketDrafts[NEW_TICKET_KEY];
        releaseRender();
        render();
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
    function ticketItem(ticket, user, { inGroup = false, query = '' } = {}) {
        const mine = isMine(ticket, user);
        const holdDate = ticket.state === '3' ? parseDate(ticket.holdExpiration) : null;
        const highlight = ticket.newMessages && !inGroup ? ' ao-item-new' : '';
        return clickable(`ao-item${isHoldExpired(ticket) ? ' ao-item-alert' : ''}${highlight}`, event => openFromCard(event.currentTarget, ticket),
            // Prima il numero INC, sotto la Short description (oggetto della mail), poi agenzia e attività.
            h('div', { class: 'ao-item-top' },
                h('span', { class: 'ao-item-number' }, markText(ticket.number, query)),
                stateChip(ticket)),
            h('div', { class: 'ao-item-title' }, markText(ticket.shortDescription, query)),
            // Durante una ricerca i risultati arrivano da entrambi gli elenchi: si dice da quale.
            h('div', { class: 'ao-item-sub' }, [ticket.companyName, ticket.activityName, query ? (ticket.messageId ? TEXTS.tickets.fromMail : TEXTS.tickets.withoutMail).toLowerCase() : ''].filter(Boolean).join(' · ')),
            inGroup ? null : newMessagesPreview(ticket),
            h('div', { class: 'ao-item-meta' },
                ticket.messageId ? (mine ? null : chip(ticket.assignedTo?.name || TEXTS.common.unassigned, 'neutral'))
                    : labelChip(mine ? { name: currentUser().displayValue, color: loadConfig().profile.label } : (ticket.label || { name: ticket.assignedTo?.name || '', color: 'grafite' })),
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
        } else if (!isCurrentMail) {
            // Prendere in carico (o ignorare) si fa solo guardando la mail: finché Outlook non la mostra, niente azioni.
            if (result.discarded) content.push(h('p', { class: 'ao-muted' }, TEXTS.mail.discardedBy(result.discarded.by || TEXTS.common.none, formatWhen(result.discarded.at))));
            content.push(h('div', { class: 'ao-open-first' },
                h('p', {}, TEXTS.mail.openFirst),
                h('button', { type: 'button', class: 'ao-button ao-primary', onclick: () => Outlook.openMessage(mail.itemId) }, TEXTS.mail.openInOutlook)));
        } else if (result.discarded) {
            content.push(discardedCard(mail, result.discarded));
            content.push(newTicketCard(mail, config, user, { allowIgnore: false }));
        } else {
            if (result.previous) content.push(previousCard(result.previous));
            if (result.candidates?.length) content.push(candidatesCard(mail, result.candidates, user));
            content.push(newTicketCard(mail, config, user));
        }
        return h('div', { class: 'ao-view' },
            backRow(),
            notice,
            mailHeader(mail, !isCurrentMail && Boolean(result?.tickets?.length)), // senza ticket c'è già il riquadro "Apri la mail"
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
    // ---------------------------- Suggerimenti --------------------------------
    // Ogni operatore ha i suoi: il pannello conta, senza mostrarlo, le combinazioni agenzia · attività con cui
    // apre i ticket e le parole degli oggetti. Con l'agenzia individuata (dal dominio o scelta a mano) propone
    // le combinazioni più usate per quell'agenzia; quelle che corrispondono alle parole della richiesta vanno in cima.
    // Il contatore parte vuoto e sta in una chiave a parte (non è nel file delle configurazioni).
    const USAGE_KEY = 'amp_panel_usage_v1';
    const USAGE_STOPWORDS = new Set(['buongiorno', 'buonasera', 'salve', 'grazie', 'saluti', 'cordiali', 'della', 'delle', 'dello', 'degli', 'nella', 'nelle', 'sono', 'questo', 'questa', 'anche', 'come', 'però', 'dopo', 'prima', 'molto', 'mille', 'potete', 'possiamo', 'vorrei', 'richiesta']);
    const usageWords = text => [...new Set((String(text || '').toLowerCase().replace(/^\s*((re|r|fw|fwd|i)\s*:\s*)+/i, '').match(/[a-zàèéìòù0-9]{4,}/g) || [])
        .filter(word => !USAGE_STOPWORDS.has(word)))];
    const Usage = {
        load() {
            try { return JSON.parse(localStorage.getItem(USAGE_KEY) || '{}') || {}; } catch { return {}; }
        },
        record(companyId, activityId, text) {
            if (!companyId || !activityId) return;
            const data = Usage.load();
            const key = `${companyId}|${activityId}`;
            const entry = data[key] || (data[key] = { n: 0, w: {} });
            entry.n += 1;
            usageWords(text).forEach(word => { entry.w[word] = (entry.w[word] || 0) + 1; });
            try { localStorage.setItem(USAGE_KEY, JSON.stringify(data)); } catch { /* memoria piena o bloccata */ }
        },
        // Attività più probabili per l'agenzia, al massimo 3.
        suggest(companyId, text, config) {
            if (!companyId) return [];
            const words = usageWords(text);
            return Object.entries(Usage.load())
                .filter(([key]) => key.startsWith(`${companyId}|`))
                .map(([key, entry]) => ({ activityId: key.slice(companyId.length + 1), score: entry.n + 5 * words.reduce((sum, word) => sum + (entry.w[word] || 0), 0) }))
                .filter(item => config.activities[item.activityId])
                .sort((a, b) => b.score - a.score)
                .slice(0, 3)
                .map(item => item.activityId);
        }
    };
    function suggestionPills(draft, config, text) {
        const company = config.companies[draft.companyId];
        const ids = company ? Usage.suggest(draft.companyId, text, config) : [];
        if (!ids.length) return null;
        return h('div', { class: 'ao-chips ao-suggestions' }, ids.map(activityId => h('button', {
            type: 'button',
            class: `ao-fav${activityId === draft.activityId ? ' ao-active' : ''}`,
            onclick: () => { draft.activityId = activityId; render(); }
        }, h('span', { class: 'ao-fav-company' }, `${company.name} ·`), ` ${config.activities[activityId].name}`)));
    }

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
    function ticketForm(mail, config, { compact = false, zone = false, user = currentUser(config) } = {}) {
        const draft = ticketDraft(mail, config);
        const companies = Object.entries(config.companies).sort(byName);
        const activities = Object.entries(config.activities).sort(byName);
        const company = config.companies[draft.companyId];
        const activity = config.activities[draft.activityId];
        const rerender = () => render();
        // Parole per i suggerimenti: oggetto e inizio della mail; senza mail, la Short description scritta.
        const suggestText = mail.manual ? draft.subject : `${mail.subject || ''} ${String(mail.body || '').slice(0, 400)}`;
        // In coda: campi raccolti in un riquadro; l'agenzia suggerita è un'etichetta accanto ad "Agenzia"
        // (la frase intera passando sopra) e le pillole stanno sotto "Attività", che è quello che scelgono.
        if (zone) {
            const suggested = draft.suggestion && draft.suggestion.id === draft.companyId;
            const companyField = selectField(TEXTS.mail.company, [{ value: '', label: TEXTS.mail.choose }, ...companies.map(([id, item]) => ({ value: id, label: item.name }))],
                draft.companyId, value => { draft.companyId = value; rerender(); });
            if (suggested) companyField.querySelector('.ao-label').append(
                h('span', { class: 'ao-suggest-tag', title: TEXTS.mail.suggestedFrom(draft.suggestion.reason) }, draft.suggestion.tag || TEXTS.mail.suggestedFrom(draft.suggestion.reason)));
            const activityField = selectField(TEXTS.mail.activity, [{ value: '', label: TEXTS.mail.choose }, ...activities.map(([id, item]) => ({ value: id, label: item.name }))],
                draft.activityId, value => { draft.activityId = value; rerender(); });
            activityField.insertBefore(h('div', { class: 'ao-suggestions-slot', dataset: { suggestFor: mail.conversationId } }, suggestionPills(draft, config, suggestText)),
                activityField.querySelector('.ao-select'));
            return h('div', { class: 'ao-form-fields' }, companyField, activityField);
        }
        return [
            selectField(TEXTS.mail.company, [{ value: '', label: TEXTS.mail.choose }, ...companies.map(([id, item]) => ({ value: id, label: item.name }))],
                draft.companyId, value => { draft.companyId = value; rerender(); }),
            // Riga sempre presente (anche vuota): se l'avviso sparisce, il modulo non cambia altezza.
            h('div', { class: 'ao-hint ao-hint-accent ao-hint-slot' },
                draft.suggestion && draft.suggestion.id === draft.companyId ? TEXTS.mail.suggestedFrom(draft.suggestion.reason) : ''),
            h('div', { class: 'ao-suggestions-slot', dataset: { suggestFor: mail.conversationId } }, suggestionPills(draft, config, suggestText)),
            selectField(TEXTS.mail.activity, [{ value: '', label: TEXTS.mail.choose }, ...activities.map(([id, item]) => ({ value: id, label: item.name }))],
                draft.activityId, value => { draft.activityId = value; rerender(); }),
            // Short description = oggetto della mail (+ prefisso dell'attività), Description = corpo: non si modificano.
            // I prefissi tra parentesi quadre (es. [FWD]) li gestirà in automatico il flusso dell'inoltro.
            !compact && company && activity ? h('dl', { class: 'ao-fields' },
                h('dt', {}, TEXTS.mail.fieldsPreview), h('dd', {}),
                fieldRow('Company', company.company?.displayValue || company.name),
                fieldRow('Assignment group', company.assignmentGroup?.displayValue),
                fieldRow('Assigned to', user.displayValue || TEXTS.common.none),
                fieldRow('Category', activity.category?.displayValue),
                fieldRow('Impact / Urgency', `${activity.impact || '?'} / ${activity.urgency || '?'}`),
                fieldRow('Time worked', TEXTS.mail.minutes(Number(activity.timeWorked) || 0)),
                fieldRow('Short description', buildShortDescription(draft.subject, activity)),
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
            // Vista del ticket (in In coda c'è il tasto rosso dal basso: ignoreBar).
            const ok = await run(() => Services.discardMail(mail.conversationId, currentUser()));
            if (ok) await refreshLists();
            render(); // se non è riuscito, la card ricompare
        });
    }
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
    function fieldRow(label, value) {
        return [h('dt', {}, label), h('dd', {}, value || TEXTS.common.none)];
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
        // Corpo completo se è la mail aperta in Outlook; altrimenti l'anteprima della Coda. Senza mail: quanto scritto.
        const body = mail.manual ? draft.description : (state.mail?.conversationId === mail.conversationId ? state.mail.body : mail.body);
        let reportedOn = mail.receivedAt ? formatServiceNowDate(parseDate(mail.receivedAt)) : '';
        if (mail.manual) {
            if (!String(draft.subject || '').trim()) return showStatus(TEXTS.tickets.missingShort, 'danger');
            if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(draft.time || '')) return showStatus(TEXTS.tickets.badTime, 'danger');
            const when = parseDate(`${draft.date} ${draft.time}`);
            if (!when) return showStatus(TEXTS.tickets.badTime, 'danger');
            reportedOn = formatServiceNowDate(when);
        }
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
            reportedOn
        };
        outcomeStart(card, 'accent', `${SPINNER}<span>${TEXTS.mail.outcomeWorking}</span>`);
        const ticket = await run(() => Services.createTicket(mail.manual ? null : mail, {
            companyId: draft.companyId, companyName: company.name,
            activityId: draft.activityId, activityName: activity.name, fields
        }), result => TEXTS.mail.ticketOpened(result.number));
        if (!ticket) {
            outcomeCancel(card);
            render();
            return;
        }
        outcomeText(card, `${CHECK_ICON}<span>${TEXTS.mail.outcomeTaken(ticket.number)}</span>`);
        await outcomeEnd(card);
        delete state.ticketDrafts[mail.conversationId];
        state.ticketCache[ticket.number] = ticket;
        if (state.queueExpanded === mail.conversationId) state.queueExpanded = null;
        if (state.extraTicketFor === mail.conversationId) state.extraTicketFor = null;
        Usage.record(draft.companyId, draft.activityId, mail.manual ? draft.subject : mail.subject);
        if (mail.manual) {
            await holdingRender(() => fold(root.querySelector('.ao-new-ticket-card .ao-inline-form')));
            state.newTicketOpen = false;
        }
        config.recent = [{ company: draft.companyId, activity: draft.activityId, timestamp: new Date().toISOString() },
            ...config.recent.filter(item => !(item.company === draft.companyId && item.activity === draft.activityId))].slice(0, 10);
        saveConfig(config);
        render();
        await refreshLists();
        outcomeRelease(card);
        // Da In coda si resta in In coda (si smista la prossima); l'avviso porta al ticket per lavorarlo.
        if (card || mail.manual) showStatus(TEXTS.mail.ticketOpened(ticket.number), 'success', { label: TEXTS.queue.goToTicket, onclick: () => openTicketItem(ticket) });
    }

    // Ticket collegato: assegnatario, tempo, azione all'invio, storico.
    function ticketSection(ticket, config, user, { isCurrentMail, showOpen = false, inDetails = false }) {
        const open = OPEN_STATES.includes(ticket.state);
        return [
            ticketCard(ticket, showOpen, inDetails, user),
            open && ticket.messageId && !ticket.mailUnconfirmed ? onSendCard(ticket, config, isCurrentMail) : null,
            open && !ticket.messageId ? updateCard(ticket, config) : null,
            open && ticket.messageId && ticket.mailUnconfirmed ? h('p', { class: 'ao-hint' }, TEXTS.mail.noLinkedMail) : null,
            historyCard(ticket)
        ];
    }
    // Assegnatario come etichetta (me / nome del collega) e, se è di un collega, "Prendi in carico".
    // Assegnatario: solo se non sono io (i miei ticket non hanno l'etichetta col nome).
    function assigneeTag(ticket, user) {
        if (isMine(ticket, user)) return null;
        return chip(ticket.assignedTo?.name || TEXTS.common.unassigned, 'person', TEXTS.mail.assignedTo(ticket.assignedTo?.name || TEXTS.common.unassigned));
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
    // Azioni sul ticket (Aggiorna, Risolvi, On Hold, Annulla) con le icone e il colore dell'azione scelta.
    function ticketActionBar(ticket, draft) {
        return iconBar(Object.entries(TEXTS.mail.actions).map(([value, label]) => ({ value, label, icon: ACTION_ICONS[value] })), draft.type,
            value => { if (value !== draft.type) swapZone(`ticket:${ticket.number}`, () => { draft.type = value; }); },
            { key: `ticket:${ticket.number}`, tones: true, slide: false });
    }
    // Contenuto che dipende dal selettore dentro una scheda (azioni di In coda e dei ticket): quello vecchio svanisce,
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
    // Ticket senza mail: lo stesso strumento, ma il testo si scrive qui e va direttamente su ServiceNow.
    function updateCard(ticket, config) {
        const draft = actionDraft(ticket, config);
        draft.text ??= '';
        const refresh = () => { const button = root.querySelector('.ao-update-card .ao-primary'); if (button) button.disabled = !draft.text.trim(); };
        return h('section', { class: 'ao-card ao-update-card' },
            // Prima la risposta (il testo), poi cosa fare del ticket.
            h('h3', { class: 'ao-card-title' }, TEXTS.mail.reply),
            textField(TEXTS.mail.updateText, draft.text, value => { draft.text = value; refresh(); }, { multiline: true, rows: '4' }),
            ticketActionBar(ticket, draft),
            choiceZone(`ticket:${ticket.number}`, actionFields(draft)),
            h('div', { class: 'ao-actions' },
                h('button', { type: 'button', class: `ao-button ao-primary${ACTION_FILLS[draft.type] || ''}`, disabled: !draft.text.trim(), onclick: () => updateTicket(ticket) }, TEXTS.mail.updateButtons[draft.type])));
    }
    async function updateTicket(ticket) {
        const config = loadConfig();
        const draft = actionDraft(ticket, config);
        if (draft.type === 'hold' && !parseDate(draft.holdExpiration)) return showStatus(TEXTS.mail.holdExpiration, 'danger');
        const user = currentUser(config);
        const updated = await run(() => Services.updateTicket(ticket.number, buildPendingAction(draft, user.displayValue), draft.text.trim(), user));
        if (!updated) return;
        rememberTicket(updated);
        delete state.actionDrafts[ticket.number];
        if (OPEN_STATES.includes(updated.state)) {
            render();
            showStatus(TEXTS.mail.updated(updated.number));
        } else {
            // Chiuso (risolto o annullato): si torna all'elenco, dove non c'è più.
            setView('tickets');
            showStatus(TEXTS.mail.closedTicket(updated.number));
        }
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
        // Prima la risposta (testo di partenza e riferimento), poi cosa fare del ticket all'invio.
        return h('section', { class: 'ao-card' },
            h('h3', { class: 'ao-card-title' }, TEXTS.mail.reply),
            ticket.pendingAction ? h('div', { class: 'ao-pending' },
                h('span', {}, TEXTS.mail.pending(pendingLabel(ticket.pendingAction))),
                h('button', { type: 'button', class: 'ao-link', onclick: () => savePendingAction(ticket, null) }, TEXTS.mail.clearPending)) : null,
            selectField(TEXTS.mail.replyText, templates, draft.template, value => { draft.template = value; }),
            checkboxField(TEXTS.mail.includeReference(ticket.number), draft.includeReference, value => { draft.includeReference = value; }),
            ticketActionBar(ticket, draft),
            choiceZone(`ticket:${ticket.number}`, actionFields(draft)),
            h('div', { class: 'ao-actions' },
                h('button', { type: 'button', class: `ao-button ao-primary${ACTION_FILLS[draft.type] || ''}`, onclick: () => prepareReply(ticket, isCurrentMail) }, TEXTS.mail.prepareReply)));
    }
    // Campi che dipendono dall'azione scelta (nessuno per Aggiorna).
    function actionFields(draft) {
        if (draft.type === 'comment') return null;
        return h('div', { class: 'ao-form-fields' },
            draft.type === 'resolve' ? selectField(TEXTS.mail.resolutionCode, RESOLUTION_CODES.map(code => ({ value: code, label: code })), draft.resolutionCode, value => { draft.resolutionCode = value; }) : null,
            draft.type === 'hold' ? [
                selectField(TEXTS.mail.holdReason, HOLD_REASONS, draft.holdReason, value => { draft.holdReason = value; }),
                textField(TEXTS.mail.holdExpiration, draft.holdExpiration, value => { draft.holdExpiration = value; }, { type: 'datetime-local' })
            ] : null,
            textField(TEXTS.mail.workNotes, draft.workNotes, value => { draft.workNotes = value; }, { multiline: true, rows: '2' }));
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
        const bar = iconBar(tabs.map(value => ({
            value, icon: SETTINGS_ICONS[value], label: TEXTS.settings.tabs[value],
            count: value !== 'profile' ? Object.keys(config[collectionOf(value)]).length : undefined
        })), tab, value => { state.settingsTab = value; state.editing = null; render(); });
        return h('div', { class: 'ao-view' },
            bar,
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
                // Colore dell'etichetta: si salva subito; la voce scelta mostra il cognome
                h('div', { class: 'ao-field' },
                    h('span', { class: 'ao-label' }, TEXTS.settings.label),
                    h('div', { class: 'ao-tag-palette', role: 'group', 'aria-label': TEXTS.settings.label },
                        Object.entries(THEME.labels).map(([key, color]) => h('button', {
                            type: 'button', class: `ao-user-tag ao-tag-swatch${key === config.profile.label ? ' ao-active' : ''}`, style: `--ao-tag-color: ${color}`,
                            title: key, 'aria-label': key, 'aria-pressed': String(key === config.profile.label),
                            onclick: () => {
                                const current = loadConfig();
                                current.profile = { ...current.profile, label: key };
                                saveConfig(current);
                                render();
                            }
                        }, key === config.profile.label ? surname(currentUser(config).displayValue) : '')))),
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
            // Tema e stile nella stessa sezione, con i selettori a icone.
            h('section', { class: 'ao-card' },
                h('h3', { class: 'ao-card-title' }, TEXTS.settings.theme),
                iconBar(Object.entries(TEXTS.settings.themes).map(([value, label]) => ({ value, label, icon: APPEARANCE_ICONS[value] })), config.preferences?.theme || 'auto', value => {
                    const current = loadConfig();
                    current.preferences = { ...(current.preferences || {}), theme: value };
                    saveConfig(current);
                    applyTheme(current);
                    render();
                }),
                iconBar(Object.entries(TEXTS.settings.styles).map(([value, label]) => ({ value, label, icon: APPEARANCE_ICONS[value] })), config.preferences?.style === 'flat' ? 'flat' : 'glass', value => {
                    const current = loadConfig();
                    current.preferences = { ...(current.preferences || {}), style: value };
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
        const vars = colors => Object.entries(colors).map(([key, value]) => `--ao-${kebab(key)}: ${value};`).join(' ');
        const glass = vars(THEME.glassColors), glassDark = vars(THEME.glassDarkColors); // il vetro scuro ridefinisce solo i valori che cambiano
        return `#ao-root { ${lines.join(' ')} }
#ao-root[data-theme="dark"] { ${dark} }
@media (prefers-color-scheme: dark) { #ao-root:not([data-theme="light"]) { ${dark} } }
#ao-root[data-style="glass"] { ${glass} }
#ao-root[data-style="glass"][data-theme="dark"] { ${glassDark} }
@media (prefers-color-scheme: dark) { #ao-root[data-style="glass"]:not([data-theme="light"]) { ${glassDark} } }`;
    }
    // Regole del CSS (per tenerlo ordinato):
    // - colori, ombre e effetti dipendono da flat/vetro e chiaro/scuro solo tramite le variabili di THEME:
    //   niente regole "solo per il vetro" o "solo per lo scuro" sui singoli elementi;
    // - una sola regola per elemento e proprietà; i gruppi comuni (sfocatura delle superfici, passaggio del mouse)
    //   sono nel blocco Superfici, in fondo;
    // - classi e variabili che non servono più si tolgono subito.
    const STYLES = `
/* L'area AMP ha l'altezza esatta del riquadro visibile: schede e barra di sistema ferme, scorre solo il contenuto */
#ao-root { position: relative; display: flex; flex-direction: column; overflow: hidden; font-family: var(--ao-font); font-size: var(--ao-fs-base); color: var(--ao-text); background: var(--ao-backdrop); line-height: 1.4; }
#ao-root *, #ao-root *::before, #ao-root *::after { box-sizing: border-box; }
#ao-root [hidden] { display: none !important; }
#ao-root.ao-busy { cursor: progress; }
#ao-root button { font: inherit; color: inherit; }

/* Intestazione e schede */
#ao-root .ao-tabs { position: relative; flex: none; display: flex; align-items: center; gap: 2px; padding: 0 4px; background: #0078d4; color: #fff; -webkit-backdrop-filter: var(--ao-nav-blur); backdrop-filter: var(--ao-nav-blur); } /* sfondo: continua l'intestazione (blendTabsWithHeader) */
#ao-root .ao-tabs .ao-icon-button { flex: none; font-size: var(--ao-fs-lg); color: rgba(255, 255, 255, .8); }
/* Freccia per tornare all'add-in: invisibile finché non la si raggiunge con il mouse, così le schede hanno tutta la larghezza */
#ao-root .ao-tabs .ao-back { position: absolute; left: 4px; top: 50%; z-index: 1; transform: translateY(-50%); opacity: 0; transition: opacity .2s; }
#ao-root .ao-tabs .ao-back:is(:hover, :focus-visible) { opacity: 1; }
#ao-root .ao-tabs .ao-icon-button:hover { background: rgba(255, 255, 255, .14); border-color: transparent; color: #fff; }
#ao-root .ao-tab { flex: 1; min-width: 0; display: flex; align-items: center; justify-content: center; gap: 0; padding: 9px 4px 7px; border: 0; border-bottom: 2px solid transparent; border-radius: var(--ao-radius-base) var(--ao-radius-base) 0 0; background: none; color: rgba(255, 255, 255, .72); font-size: var(--ao-fs-lg); font-weight: var(--ao-fw-normal); cursor: pointer; }
#ao-root .ao-tab:hover { color: #fff; }
#ao-root .ao-tab.ao-active { color: #fff; }
#ao-root .ao-tab-icon { position: relative; display: inline-flex; opacity: .8; transition: color .25s, opacity .25s; }
#ao-root .ao-tab.ao-active .ao-tab-icon { color: #fff; opacity: 1; }
/* Ogni scheda ha il suo posto: chiusa solo l'icona, aperta anche nome e contatore (si allargano nel suo posto) */
#ao-root .ao-tab-label { max-width: 0; margin-left: 0; overflow: hidden; white-space: nowrap; opacity: 0; transition: max-width var(--ao-selection-motion), margin-left var(--ao-selection-motion), opacity .25s; }
#ao-root .ao-tab.ao-active .ao-tab-label { max-width: 120px; margin-left: 9px; opacity: 1; } /* spazio anche per il segnalino */
@media (prefers-reduced-motion: reduce) { #ao-root .ao-tab-label { transition: none; } }
#ao-root .ao-tab-indicator { position: absolute; left: 0; bottom: 0; height: 3px; width: 0; border-radius: 3px 3px 0 0; background: #fff; pointer-events: none; opacity: 0; }
#ao-root .ao-tab-indicator.ao-ready { opacity: 1; }
#ao-root .ao-tab-indicator.ao-moving { transition: transform .32s cubic-bezier(.3, .8, .2, 1), width .32s cubic-bezier(.3, .8, .2, 1); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-tab-indicator.ao-moving { transition: none; } }
/* Scheda aperta: leggero zoom della scritta (e del numero), che resta finché è selezionata */
#ao-root .ao-tab > * { transition: transform .35s ease-in-out; }
#ao-root .ao-tab.ao-active > * { transform: scale(1.04); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-tab > * { transition: none; } }
/* Schede: il conteggio è un numero pulito, solo sulla scheda aperta. */
#ao-root .ao-counter { display: inline-block; min-width: 18px; height: 18px; padding: 0 4px; border-radius: var(--ao-radius-badge); color: var(--ao-text-muted); font-size: var(--ao-fs-md); font-weight: var(--ao-fw-medium); line-height: 18px; text-align: center; }
#ao-root .ao-tab.ao-active .ao-counter { color: #fff; }
/* Segnalino della scheda: sparisce rimpicciolendosi e svanendo, e la scritta si ricentra senza scatti (e il contrario quando torna) */
#ao-root .ao-tab .ao-counter { margin-left: 5px; max-width: 40px; color: rgba(255, 255, 255, .72); overflow: hidden; transition: opacity .2s ease, max-width .28s ease, min-width .28s ease, margin-left .28s ease, padding .28s ease, transform .28s ease; }
#ao-root .ao-tab .ao-counter.ao-counter-off { opacity: 0; max-width: 0; min-width: 0; margin-left: 0; padding: 0; transform: scale(.6); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-tab .ao-counter { transition: none; } }
/* Aggiornamenti: segnalino tondo con il numero sull'icona, sempre visibile, del colore del testo intorno. */
#ao-root .ao-notify { position: absolute; top: -6px; right: -7px; min-width: 15px; height: 15px; padding: 0 3px; border-radius: 999px; background: currentColor; font-size: var(--ao-fs-xs); font-weight: var(--ao-fw-bold); line-height: 15px; text-align: center; pointer-events: auto; transition: opacity .2s, transform .2s; }
#ao-root .ao-notify-number { color: var(--ao-notify-ink); }
#ao-root .ao-tab .ao-notify-number { color: #005a9e; } /* blu dell'intestazione */
#ao-root .ao-notify-off { opacity: 0; transform: scale(.6); pointer-events: none; }
#ao-root .ao-notify-anchor { position: relative; display: inline-flex; margin-right: 5px; }
/* Card: numero dei messaggi nuovi su sfondo blu. */
#ao-root .ao-news { display: inline-block; min-width: 18px; height: 18px; padding: 0 4px; border-radius: var(--ao-radius-badge); background: var(--ao-accent); color: var(--ao-white); font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-bold); line-height: 18px; text-align: center; }
#ao-root .ao-new-head .ao-news { align-self: center; }
#ao-root .ao-body { flex: 1; min-height: 0; display: flex; flex-direction: column; }
#ao-root .ao-view { flex: 1; min-height: 0; display: flex; flex-direction: column; }
/* In alto, fermi: barra secondaria o "←", con una linea sottile come confine */
/* Barra secondaria scontornata: nessuna fascia né separatore, solo il menu.
   A casa sta in cima all'elenco; fuori casa compare sovrapposto al contenuto, con un'ombra sotto il menu stesso. */
#ao-root .ao-fixed { --ao-lift: 0; position: sticky; top: 0; z-index: 3; display: grid; gap: 10px; padding-top: 12px; pointer-events: none; will-change: transform; }
/* Completamento del movimento quando lo scorrimento si ferma a metà */
#ao-root .ao-fixed.ao-snap { transition: transform .18s ease-out; }
/* Ricomparsa: discesa morbida con un lieve rimbalzo; l'ombra si accende mentre arriva */
#ao-root .ao-fixed.ao-reveal { transition: transform .42s cubic-bezier(.2, 1.3, .4, 1); }
#ao-root .ao-fixed.ao-reveal > * { transition: box-shadow .35s ease-out .08s, background-color .35s ease-out .08s; }
#ao-root .ao-fixed > * { pointer-events: auto; }
/* Ombra (e fondo della riga "indietro") graduati da --ao-lift: 0 a casa, 1 sovrapposta al contenuto */
#ao-root .ao-fixed .ao-segmented,
#ao-root .ao-fixed .ao-bar-btn,
#ao-root .ao-fixed .ao-back-row { box-shadow: 0 3px 10px rgba(0,0,0, calc(.18 * var(--ao-lift))); transition: box-shadow .25s ease-out, background-color .25s ease-out; }
/* Ricerca aperta: il suo bagliore (vetro) insieme all'ombra di sollevamento */
#ao-root .ao-fixed .ao-search-bar, #ao-root .ao-fixed .ao-bar-search.ao-active, #ao-root .ao-fixed .ao-bar-filter.ao-open { box-shadow: var(--ao-hover-glow), 0 3px 10px rgba(0,0,0, calc(.18 * var(--ao-lift))); }
#ao-root .ao-fixed .ao-back-row { justify-self: start; margin-left: -10px; padding: 4px 10px; border-radius: var(--ao-radius-base); background: color-mix(in srgb, var(--ao-surface) calc(100% * var(--ao-lift)), transparent); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-fixed.ao-snap, #ao-root .ao-fixed.ao-reveal, #ao-root .ao-fixed.ao-reveal > * { transition: none; } }
/* Sotto la linea scorre il contenuto, a tutta larghezza (barra di scorrimento sottile, spazio riservato) */
/* Barra di scorrimento sottile (4px): lo spazio a destra è quasi uguale a quello a sinistra */
#ao-root .ao-scroll { flex: 1; min-height: 0; display: grid; align-content: start; gap: 10px; padding: 0 8px 12px 12px; scroll-padding-top: 60px; overflow-y: auto; overflow-x: hidden; scrollbar-gutter: stable; }
@supports not selector(::-webkit-scrollbar) { #ao-root .ao-scroll { scrollbar-width: thin; scrollbar-color: var(--ao-line-strong) transparent; } }
#ao-root .ao-scroll::-webkit-scrollbar { width: 4px; }
#ao-root .ao-scroll::-webkit-scrollbar-thumb { border-radius: 2px; background: var(--ao-line-strong); }
#ao-root .ao-scroll::-webkit-scrollbar-track { background: transparent; }

/* Pulsanti */
#ao-root .ao-button { display: inline-flex; align-items: center; justify-content: center; gap: 6px; min-height: var(--ao-control-height); padding: 0 12px; border: 1px solid var(--ao-line-strong); border-radius: var(--ao-radius-base); background: var(--ao-surface); font-weight: var(--ao-fw-medium); cursor: pointer; transition: background-color .3s, box-shadow .35s, filter .3s; }
#ao-root .ao-button:hover:where(:not(:disabled)) { background: var(--ao-hover); border-color: var(--ao-text-subtle); }
/* Pulsante principale, ovunque: pieno del colore della sua azione (--ao-btn, classi ao-fill-*), scritta bianca;
   nel vetro è vetro colorato (pietre preziose) con un filo di luce; al passaggio si accende appena */
#ao-root .ao-button.ao-primary { --ao-btn: var(--ao-tone-take); background: color-mix(in srgb, var(--ao-btn) var(--ao-tone-fill), transparent); border-color: var(--ao-tone-line); color: var(--ao-white); box-shadow: var(--ao-btn-highlight); }
#ao-root .ao-button.ao-primary:hover:where(:not(:disabled)) { background: color-mix(in srgb, var(--ao-btn) var(--ao-tone-fill), transparent); border-color: var(--ao-tone-line); filter: var(--ao-tone-hover-filter); box-shadow: var(--ao-btn-highlight); }
#ao-root .ao-button.ao-fill-forward { --ao-btn: var(--ao-tone-forward); }
#ao-root .ao-button.ao-fill-danger { --ao-btn: var(--ao-tone-ignore); }
#ao-root .ao-button.ao-fill-resolve { --ao-btn: var(--ao-tone-resolve); }
#ao-root .ao-button.ao-fill-hold { --ao-btn: var(--ao-tone-hold); }
#ao-root .ao-subtle { border-color: transparent; background: transparent; color: var(--ao-text-muted); }
#ao-root .ao-small { min-height: 26px; padding: 0 9px; font-size: var(--ao-fs-md); }
/* Prendi in carico (righe di In coda): azzurro leggero al passaggio del mouse */
#ao-root .ao-take:hover { color: var(--ao-text-accent); background: var(--ao-accent-bg); border-color: var(--ao-accent-border); }
/* Ignora: solo la scritta diventa rossa (niente sfondo né bordo), anche quando chiede conferma */
#ao-root .ao-ignore:hover { color: var(--ao-danger); background: transparent; border-color: transparent; }
#ao-root .ao-ignore.ao-armed, #ao-root .ao-ignore.ao-armed:hover { background: transparent !important; border-color: transparent !important; }
#ao-root .ao-armed, #ao-root .ao-armed:hover { background: var(--ao-danger-bg) !important; border-color: var(--ao-danger) !important; color: var(--ao-danger) !important; }
#ao-root .ao-icon-button { display: inline-flex; align-items: center; justify-content: center; min-width: 28px; height: 28px; padding: 0 6px; border: 1px solid transparent; border-radius: var(--ao-radius-base); background: none; color: var(--ao-text-muted); cursor: pointer; }
#ao-root .ao-icon-button:hover { background: var(--ao-hover); border-color: var(--ao-line); }
#ao-root .ao-danger:hover { color: var(--ao-danger); }
#ao-root .ao-link { padding: 0; border: 0; background: none; color: var(--ao-text-accent); font-size: var(--ao-fs-md); cursor: pointer; text-align: left; }
#ao-root .ao-link:hover { text-decoration: underline; }
#ao-root :disabled { opacity: var(--ao-disabled-opacity); cursor: not-allowed; }
#ao-root :focus-visible { outline: none; box-shadow: var(--ao-shadow-focus); }

/* Filtri e scelte a segmenti */
#ao-root .ao-segmented { position: relative; display: flex; padding: 2px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-segment-track); }
/* Al passaggio del mouse la voce cambia colore (e nel vetro si schiarisce appena); la pillola della voce scelta scivola */
#ao-root .ao-segment { position: relative; z-index: 1; flex: 1; display: inline-flex; align-items: center; justify-content: center; gap: 5px; min-height: 28px; padding: 0 6px; border: 0; border-radius: calc(var(--ao-radius-base) - 2px); background: none; color: var(--ao-text-muted); font-size: var(--ao-fs-md); cursor: pointer; white-space: nowrap; transition: color .25s; }
#ao-root .ao-segment:hover { color: var(--ao-text); }
#ao-root .ao-segment:not(.ao-active):hover { background: var(--ao-segment-hover); }
#ao-root .ao-segment.ao-active { color: var(--ao-text-accent); font-weight: var(--ao-fw-medium); }
#ao-root .ao-segment-thumb { position: absolute; top: 2px; bottom: 2px; left: 2px; width: 0; border-radius: calc(var(--ao-radius-base) - 2px); background: var(--ao-selection); box-shadow: inset 0 0 0 1px var(--ao-selection-line), var(--ao-selection-shadow); opacity: 0; pointer-events: none; }
#ao-root .ao-segment-thumb.ao-ready { opacity: 1; }
#ao-root .ao-segment-thumb.ao-moving { transition: transform var(--ao-selection-motion), width var(--ao-selection-motion); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-segment-thumb.ao-moving, #ao-root .ao-segment { transition: none; } }
/* Ticket: selettore Da mail | Senza mail largo come gli altri, con il filtro operatore nella voce scelta e il menu sotto */
#ao-root .ao-source-bar { position: relative; display: grid; gap: 6px; }
/* Riga della barra secondaria: [pulsante] selettore [pulsante]. I pulsanti sono staccati; dove non servono
   (In coda) resta lo spazio vuoto, così il selettore ha la stessa larghezza in tutte le sezioni. */
#ao-root .ao-bar-row { display: flex; align-items: stretch; gap: 6px; min-width: 0; }
#ao-root .ao-bar-row > .ao-segmented { flex: 1; min-width: 0; }
#ao-root .ao-bar-btn, #ao-root .ao-bar-spacer { position: relative; flex: none; width: 34px; }
#ao-root .ao-bar-btn { display: inline-flex; align-items: center; justify-content: center; padding: 0; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); color: var(--ao-text-muted); cursor: pointer; transition: background-color .3s, color .2s, box-shadow .35s, filter .3s; }
#ao-root .ao-bar-btn:hover, #ao-root .ao-bar-btn.ao-open { background: var(--ao-hover); }
#ao-root .ao-bar-btn.ao-active { color: var(--ao-text-accent); }
#ao-root .ao-bar-filter.ao-active:not(.ao-open)::after { content: ""; position: absolute; top: 5px; right: 5px; width: 6px; height: 6px; border-radius: 50%; background: var(--ao-accent); }
#ao-root .ao-bar-search.ao-active, #ao-root .ao-bar-filter.ao-open { border-color: var(--ao-search-lens-line); background: var(--ao-search-lens-bg); filter: var(--ao-hover-filter); box-shadow: var(--ao-hover-glow); }
/* Ricerca: prende il posto della barra, stessa misura (lente a sinistra, × a destra) */
#ao-root .ao-search-bar { align-items: stretch; min-height: 34px; border-color: var(--ao-search-line); background: var(--ao-search-bg); filter: var(--ao-hover-filter); box-shadow: var(--ao-hover-glow); }
/* Barra di ricerca in movimento (apertura e chiusura): sopra il selettore, alla stessa altezza */
#ao-root .ao-search-bar:focus-within { border-color: var(--ao-search-focus-line); }
#ao-root .ao-search-input { flex: 1; min-width: 0; padding: 0 10px; border: 0; outline: none; background: none; color: var(--ao-text); font: inherit; }
#ao-root .ao-search-input::-webkit-search-cancel-button { display: none; }
/* Filtri (imbuto): Da mail | Senza mail si trasformano nei due filtri (morphBar). Due metà sul fondo del
   selettore; il filtro attivo è pieno come la voce scelta dei selettori. */
#ao-root .ao-filter-bar { flex: 1; min-width: 0; gap: 2px; }
#ao-root .ao-filter-dd { position: relative; flex: 1; min-width: 0; min-height: 28px; display: flex; align-items: center; gap: 4px; padding: 0 6px 0 10px; border-radius: calc(var(--ao-radius-base) - 2px); background: none; color: var(--ao-text-subtle); font-size: var(--ao-fs-md); cursor: pointer; transition: background-color .2s, color .2s; }
#ao-root .ao-filter-dd:hover { color: var(--ao-text); }
#ao-root .ao-filter-dd.ao-set { background: var(--ao-selection); box-shadow: inset 0 0 0 1px var(--ao-selection-line), var(--ao-selection-shadow); color: var(--ao-text-accent); font-weight: var(--ao-fw-medium); }
#ao-root .ao-filter-value { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
#ao-root .ao-filter-caret { flex: none; font-size: var(--ao-fs-xs); }
#ao-root .ao-filter-menu { left: 0; top: calc(100% + 6px); font-weight: var(--ao-fw-normal); }
#ao-root .ao-filter-menu-operator { left: auto; right: 0; }
#ao-root .ao-menu-default { margin-left: auto; color: var(--ao-text-subtle); font-size: var(--ao-fs-xs); font-weight: var(--ao-fw-normal); }
#ao-root .ao-search-input:focus, #ao-root .ao-search-input:focus-visible { outline: none; box-shadow: none; }
#ao-root .ao-chip-clear { display: inline-flex; flex: none; align-items: center; justify-content: center; width: 18px; height: 18px; padding: 0; border: 0; border-radius: 50%; background: var(--ao-line-light); color: var(--ao-text-muted); font-size: var(--ao-fs-md); line-height: 1; cursor: pointer; }
#ao-root .ao-chip-clear:hover { background: var(--ao-hover); color: var(--ao-text); }
/* Ricerca: scheletro di caricamento che pulsa, poi i risultati entrano dal basso */
#ao-root .ao-skeleton-list { display: grid; gap: 6px; }
#ao-root .ao-skeleton { display: grid; gap: 7px; padding: 11px 10px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); }
#ao-root .ao-skeleton-line { display: block; height: 9px; border-radius: 5px; background: linear-gradient(90deg, var(--ao-line-light) 25%, var(--ao-bg) 50%, var(--ao-line-light) 75%); background-size: 200% 100%; animation: ao-shimmer 1.1s linear infinite; }
#ao-root .ao-skeleton-short { width: 32%; }
#ao-root .ao-skeleton-mid { width: 58%; }
@keyframes ao-shimmer { from { background-position: 100% 0; } to { background-position: -100% 0; } }
#ao-root .ao-results-in > * { animation: ao-result-in .28s cubic-bezier(.2, .8, .2, 1) backwards; }
#ao-root .ao-results-in > :nth-child(2) { animation-delay: .04s; }
#ao-root .ao-results-in > :nth-child(3) { animation-delay: .08s; }
#ao-root .ao-results-in > :nth-child(n+4) { animation-delay: .12s; }
@keyframes ao-result-in { from { opacity: 0; transform: translateY(6px); } }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-skeleton-line, #ao-root .ao-results-in > * { animation: none; } }
#ao-root .ao-mark { padding: 0 1px; border-radius: 2px; background: var(--ao-warning-bg); color: inherit; }
#ao-root .ao-menu { position: absolute; top: calc(100% + 4px); z-index: 4; display: grid; min-width: 170px; padding: 4px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-menu-bg); box-shadow: 0 6px 16px rgba(0,0,0,.14); animation: ao-menu-in .16s ease-out; }
@keyframes ao-menu-in { from { opacity: 0; transform: translateY(-4px); } }
#ao-root .ao-menu-item { display: flex; align-items: center; gap: 10px; padding: 6px 10px; border: 0; border-radius: calc(var(--ao-radius-base) - 2px); background: none; color: var(--ao-text); text-align: left; font-size: var(--ao-fs-md); cursor: pointer; }
#ao-root .ao-menu-item:hover { background: var(--ao-accent-bg); }
#ao-root .ao-menu-item.ao-active { color: var(--ao-text-accent); font-weight: var(--ao-fw-medium); } /* scelta: solo la scritta; lo sfondo resta al passaggio del mouse */
#ao-root .ao-menu-sep { margin: 3px 0; border: 0; border-top: 1px solid var(--ao-line-light); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-menu { animation: none; } }
/* Selettore con icone (Impostazioni): la voce scelta è più larga e ha il nome, che compare in dissolvenza */
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
@media (prefers-reduced-motion: reduce) { #ao-root .ao-segment-name-in { animation: none; } #ao-root .ao-segment-name-out { display: none; } #ao-root .ao-segmented-icons.ao-growing .ao-segment { transition: none; } }
#ao-root .ao-count { color: var(--ao-text-subtle); font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-normal); }
#ao-root .ao-card-title .ao-count { margin-left: 6px; }

/* Elenchi di Coda e Ticket */
#ao-root .ao-list { display: grid; gap: 6px; margin: 0; padding: 0; list-style: none; }
#ao-root .ao-item { display: grid; gap: 3px; padding: 9px 10px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); box-shadow: var(--ao-shadow-card); cursor: pointer; transition: background-color .3s, box-shadow .35s, filter .3s, transform .32s cubic-bezier(.3, .8, .2, 1); }
/* Ticket che si apre o si richiude: sollevato come la scheda aperta di In coda (openFromCard, closeToList) */
#ao-root .ao-item.ao-opening { position: relative; z-index: 1; transform: translateY(-2px) scale(1.015); box-shadow: 0 6px 16px rgba(0,0,0,.12), 0 1px 3px rgba(0,0,0,.08); }
/* Il blu (titolo e bordo) è riservato alla scheda selezionata (modulo o dettagli aperti), che non reagisce all'hover
   (il passaggio del mouse, uguale per tutte le superfici, è nel blocco Superfici). */
#ao-root .ao-expanded .ao-item-title { color: var(--ao-text-accent); }
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
#ao-root .ao-chip { display: inline-flex; align-items: center; padding: 1px 8px; border-radius: var(--ao-radius-pill); box-shadow: inset 0 0 0 1px var(--ao-chip-line); font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-medium); white-space: nowrap; }
#ao-root .ao-chip.ao-tone-neutral { background: var(--ao-line-light); color: var(--ao-text-muted); }
#ao-root .ao-chip.ao-tone-person { background: var(--ao-surface); color: var(--ao-text); box-shadow: inset 0 0 0 1px var(--ao-text-subtle); }
#ao-root .ao-tags { display: flex; flex-direction: column; align-items: flex-end; gap: 4px; flex: none; }
#ao-root .ao-chip.ao-tone-accent { background: var(--ao-accent-bg); color: var(--ao-text-accent); }
#ao-root .ao-chip.ao-tone-success { background: var(--ao-success-bg); color: var(--ao-success); }
#ao-root .ao-chip.ao-tone-warning { background: var(--ao-warning-bg); color: var(--ao-warning); }
#ao-root .ao-chip.ao-tone-danger { background: var(--ao-danger-bg); color: var(--ao-danger); }
#ao-root .ao-chip.ao-tone-hold { background: var(--ao-hold-bg); color: var(--ao-hold); }
/* Etichetta dell'operatore: un colore a testa (--ao-tag-color), fondo e scritta ricavati secondo lo stile */
#ao-root .ao-user-tag { display: inline-flex; align-items: center; flex: none; height: 20px; padding: 0 9px; border: 0; border-radius: var(--ao-radius-pill); background: color-mix(in srgb, var(--ao-tag-color) var(--ao-tag-fill), var(--ao-tag-base)); color: color-mix(in srgb, var(--ao-tag-color) var(--ao-tag-ink), var(--ao-tag-ink-base)); box-shadow: inset 0 0 0 1px var(--ao-tag-line); font: inherit; font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-medium); line-height: 1; white-space: nowrap; }
#ao-root button.ao-user-tag { cursor: pointer; }
#ao-root .ao-user-tag-in { animation: ao-tag-in .28s cubic-bezier(.32, 1.28, .5, 1); }
@keyframes ao-tag-in { from { transform: scale(.6); opacity: 0; } }
/* Posto vuoto per la propria etichetta: più stretto di qualunque etichetta, con il fondo dei selettori */
#ao-root .ao-user-tag-add { display: inline-flex; align-items: center; justify-content: center; flex: none; width: 28px; height: 20px; padding: 0; border: 0; border-radius: var(--ao-radius-pill); background: var(--ao-segment-track); box-shadow: inset 0 0 0 1px var(--ao-line), inset 0 0 0 1px var(--ao-chip-line); color: var(--ao-text-subtle); cursor: pointer; transition: color .2s, background-color .2s; }
#ao-root .ao-user-tag-add:hover { color: var(--ao-text); background: linear-gradient(var(--ao-segment-hover), var(--ao-segment-hover)), var(--ao-segment-track); }
/* Profilo: scelta del colore */
#ao-root .ao-tag-palette { display: grid; grid-template-columns: repeat(5, 1fr); gap: 6px; }
#ao-root .ao-tag-swatch { justify-content: center; height: 24px; padding: 0 4px; overflow: hidden; }
#ao-root .ao-tag-swatch.ao-active { outline: 2px solid color-mix(in srgb, var(--ao-tag-color) var(--ao-tag-ink), var(--ao-tag-ink-base)); outline-offset: 2px; }
/* Etichetta di un collega a scheda aperta: cliccabile; pronta al passaggio mostra il suo cognome barrato e il proprio */
#ao-root .ao-tag-swap { gap: 4px; }
#ao-root .ao-tag-swap s { opacity: .7; }
#ao-root .ao-tag-swap .ao-tag-next { color: color-mix(in srgb, var(--ao-tag-new) var(--ao-tag-ink), var(--ao-tag-ink-base)); }
#ao-root .ao-armed-tag { box-shadow: inset 0 0 0 1px color-mix(in srgb, var(--ao-tag-new) 55%, transparent); animation: ao-fade-in .2s linear; }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-user-tag-in, #ao-root .ao-armed-tag { animation: none; } }

/* Riquadri */
#ao-root .ao-card { display: grid; gap: 8px; padding: 12px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); box-shadow: var(--ao-shadow-card); }
#ao-root .ao-card-head { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
#ao-root .ao-card-title { margin: 0; font-size: var(--ao-fs-lg); font-weight: var(--ao-fw-medium); }
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
/* Elenco a discesa: campo come gli altri, con la freccetta; l'elenco è quello dei filtri */
#ao-root .ao-select { display: flex; align-items: center; gap: 6px; cursor: pointer; user-select: none; }
#ao-root .ao-select-value { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
#ao-root .ao-select-placeholder { color: var(--ao-text-subtle); }
#ao-root .ao-select-caret { flex: none; color: var(--ao-text-subtle); font-size: var(--ao-fs-xs); transition: transform .2s; }
#ao-root .ao-select.ao-open { border-color: var(--ao-field-focus-line); box-shadow: var(--ao-field-focus-glow); }
#ao-root .ao-select.ao-open .ao-select-caret { transform: rotate(180deg); }
#ao-root .ao-select-menu { z-index: 20; max-height: 240px; overflow-y: auto; min-width: 0; scrollbar-width: thin; }
#ao-root .ao-input:focus { border-color: var(--ao-field-focus-line); box-shadow: var(--ao-field-focus-glow); outline: none; }
#ao-root .ao-suggestions-slot:empty { display: none; }
/* Campi dei moduli (Prendi in carico, Nuovo ticket, Risposta): campi semplici sulla scheda, come in tutto il pannello;
   l'agenzia suggerita sta accanto all'etichetta, le pillole delle attività sotto "Attività" */
#ao-root .ao-form-fields { display: grid; gap: 8px; }
#ao-root .ao-form-fields .ao-label { display: flex; align-items: center; justify-content: space-between; gap: 6px; }
#ao-root .ao-suggest-tag { padding: 0 6px; border-radius: var(--ao-radius-badge); background: var(--ao-accent-bg); color: var(--ao-text-accent); font-weight: var(--ao-fw-medium); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
#ao-root .ao-form-fields .ao-suggestions-slot { margin-bottom: 3px; }
#ao-root .ao-form-fields .ao-fav { padding: 4px 10px; border-color: transparent; box-shadow: 0 1px 2px rgba(0,0,0,.06); color: var(--ao-text); transition: background-color .2s, color .2s; }
#ao-root .ao-form-fields .ao-fav:hover { border-color: transparent; background: var(--ao-hover); }
#ao-root .ao-form-fields .ao-fav.ao-active { border-color: transparent; box-shadow: none; background: color-mix(in srgb, var(--ao-accent) 14%, var(--ao-surface)); color: var(--ao-text-accent); font-weight: var(--ao-fw-normal); }
#ao-root .ao-form-fields .ao-fav-company { color: var(--ao-text-subtle); }
#ao-root .ao-form-fields .ao-fav.ao-active .ao-fav-company { color: inherit; }
#ao-root .ao-new-ticket { display: flex; align-items: center; justify-content: center; gap: 6px; margin-bottom: 6px; padding: 8px; border: var(--ao-call-border); border-radius: var(--ao-radius-base); background: var(--ao-surface); box-shadow: var(--ao-call-glow); color: var(--ao-text-accent); font: inherit; font-weight: var(--ao-fw-medium); cursor: pointer; transition: background-color .3s, box-shadow .35s, filter .3s; }
/* Aperto: come una scheda di In coda aperta (sollevato, gli altri ticket sfocati), col suo bagliore */
#ao-root .ao-new-ticket-card { position: relative; z-index: 1; border-color: var(--ao-call-open-border); transform: translateY(-2px) scale(1.015); box-shadow: var(--ao-call-glow), 0 6px 16px rgba(0,0,0,.12), 0 1px 3px rgba(0,0,0,.08); animation: ao-lift .32s cubic-bezier(.3, .8, .2, 1); }
@keyframes ao-lift { from { transform: none; box-shadow: var(--ao-call-glow); } }
#ao-root .ao-view:has(.ao-new-ticket-card) .ao-list .ao-item, #ao-root .ao-view:has(.ao-new-ticket-card) .ao-list .ao-ticket-compact { filter: blur(5px); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-new-ticket-card { animation: none; } }
#ao-root .ao-new-ticket-card .ao-inline-form { border-top: 0; padding-top: 0; }
#ao-root .ao-field-row { display: grid; grid-template-columns: minmax(0, 1fr) 110px; gap: 8px; }
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
#ao-root .ao-config-name { font-weight: var(--ao-fw-medium); }
#ao-root .ao-editor { display: grid; gap: 8px; padding: 10px; border: 1px solid var(--ao-accent-border); border-radius: var(--ao-radius-base); background: var(--ao-accent-bg); }

/* In coda: riga con azioni e modulo compatto */
#ao-root .ao-open-first { display: grid; justify-items: start; gap: 8px; padding: 12px; border: 1px dashed var(--ao-line-strong); border-radius: var(--ao-radius-base); color: var(--ao-text-muted); }
#ao-root .ao-open-first p { margin: 0; }
#ao-root .ao-item-main { display: grid; grid-template-columns: minmax(0, 1fr) auto; gap: 3px 10px; cursor: pointer; border-radius: calc(var(--ao-radius-base) - 2px); }
#ao-root .ao-item-left { display: grid; gap: 3px; align-content: start; min-width: 0; }
#ao-root .ao-inline-form { display: grid; gap: 8px; padding-top: 8px; border-top: 1px solid var(--ao-line-light); }
/* Scheda di In coda aperta: testo della mail */
#ao-root .ao-mail-preview { max-height: 180px; overflow: hidden; padding: 0 0 10px; border-bottom: 1px solid var(--ao-line-light); color: var(--ao-text); font-size: var(--ao-fs-md); line-height: 1.5; white-space: pre-line; overflow-wrap: anywhere; scrollbar-width: thin; }
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
#ao-root .ao-choice-accent { --ao-choice-bg: color-mix(in srgb, var(--ao-accent) 16%, var(--ao-surface)); --ao-choice-text: var(--ao-text-accent); } /* più deciso dell'azzurro dei filtri: si stacca dal fondo grigio */
#ao-root .ao-choice-forward { --ao-choice-bg: var(--ao-choice-forward-bg); --ao-choice-text: var(--ao-choice-forward-text); }
#ao-root .ao-choice-danger { --ao-choice-bg: var(--ao-choice-danger-bg); --ao-choice-text: var(--ao-choice-danger-text); }
#ao-root .ao-choice-resolve { --ao-choice-bg: var(--ao-choice-resolve-bg); --ao-choice-text: var(--ao-choice-resolve-text); }
#ao-root .ao-choice-hold { --ao-choice-bg: var(--ao-choice-hold-bg); --ao-choice-text: var(--ao-choice-hold-text); }
/* Selettori delle azioni (In coda, ticket): la voce scelta prende il colore della sua azione */
#ao-root .ao-segmented.ao-toned .ao-segment-thumb { background: color-mix(in srgb, var(--ao-choice-bg) var(--ao-selection-tone), var(--ao-selection)); }
#ao-root .ao-segmented.ao-toned .ao-segment.ao-active { color: var(--ao-choice-text); }
#ao-root .ao-choice-option.ao-active { color: var(--ao-choice-text); }
#ao-root .ao-choice-zone:empty { display: none; }
#ao-root .ao-choice-content { display: grid; gap: 8px; }
#ao-root .ao-collapse { position: relative; top: 5px; /* più distante dai pulsanti, dentro il margine della scheda: l'altezza non cambia */ display: flex; justify-content: center; align-items: center; justify-self: center; width: 44px; height: 22px; margin-top: -2px; padding: 0; border: 0; border-radius: var(--ao-radius-base); background: none; color: var(--ao-text-muted); cursor: pointer; }
#ao-root .ao-collapse:hover { background: var(--ao-hover); color: var(--ao-text); }
#ao-root .ao-svg-icon { display: inline-flex; }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-choice-thumb, #ao-root .ao-choice-option { transition: none; animation: none; } }
/* Esito su una scheda di In coda: banner colorato (le animazioni sono nel blocco Animazioni dello script) */
#ao-root .ao-queue-item { position: relative; overflow: hidden; gap: 6px; transition: background-color .3s, border-color .3s, transform .32s cubic-bezier(.3, .8, .2, 1), box-shadow .32s ease, filter .32s ease; }
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
#ao-root .ao-queue-item.ao-expanded[data-choice="forward"] .ao-item-title { color: var(--ao-choice-forward-text); }
#ao-root .ao-queue-item.ao-expanded[data-choice="ignore"] .ao-item-title { color: var(--ao-choice-danger-text); }
#ao-root .ao-resolving { pointer-events: none; }
/* Colore dell'esito come quello dei pulsanti (vetro colorato nel vetro), scritta bianca */
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
@media (prefers-reduced-motion: reduce) { #ao-root .ao-queue-item, #ao-root .ao-outcome { transition: none; animation: none; } }

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
#ao-root .ao-item.ao-item-new:not(.ao-opening) { box-shadow: inset 3px 0 0 var(--ao-accent), var(--ao-shadow-card); }
#ao-root .ao-item.ao-item-new:not(.ao-expanded):hover { box-shadow: inset 3px 0 0 var(--ao-accent), var(--ao-hover-glow); } /* ticket con messaggi nuovi: bordino azzurro */
#ao-root .ao-multi { display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 8px 10px; border-radius: var(--ao-radius-base); background: var(--ao-accent-bg); font-size: var(--ao-fs-md); }
#ao-root .ao-multi-count { font-size: var(--ao-fs-lg); color: var(--ao-text-accent); }
#ao-root .ao-ticket-compact { display: grid; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); box-shadow: var(--ao-shadow-card); transition: background-color .3s, box-shadow .35s, filter .3s; }
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
#ao-root .ao-candidate { display: flex; align-items: center; gap: 8px; padding: 8px 10px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); transition: box-shadow .35s, filter .3s; }
#ao-root .ao-candidate-text { flex: 1; display: grid; gap: 2px; min-width: 0; }
#ao-root .ao-check { display: flex; align-items: center; gap: 6px; color: var(--ao-text-muted); font-size: var(--ao-fs-md); cursor: pointer; }
#ao-root .ao-check input { margin: 0; accent-color: var(--ao-accent); }

/* Barra di sistema in fondo all'area (stato dei servizi, allineamento), sempre visibile */
#ao-root .ao-statusbar { flex: none; display: flex; align-items: center; gap: 6px; height: 28px; padding: 0 12px; border-top: 1px solid var(--ao-line); background: var(--ao-surface); color: var(--ao-text-muted); font-size: var(--ao-fs-sm); }
/* Superfici: quali elementi sfocano quello che hanno dietro (solo nel vetro) e il passaggio del mouse,
   uguale per tutte (nel flat solo lo sfondo, nel vetro si accendono con un bagliore leggero). */
#ao-root .ao-item, #ao-root .ao-ticket-compact, #ao-root .ao-card, #ao-root .ao-config-card, #ao-root .ao-segmented, #ao-root .ao-bar-btn, #ao-root .ao-new-ticket, #ao-root .ao-candidate, #ao-root .ao-statusbar { -webkit-backdrop-filter: var(--ao-blur); backdrop-filter: var(--ao-blur); }
#ao-root .ao-item:not(.ao-expanded):hover, #ao-root .ao-ticket-compact:not(.ao-expanded):hover { background: var(--ao-hover); filter: var(--ao-hover-filter); box-shadow: var(--ao-hover-glow); }
#ao-root .ao-item:not(.ao-expanded):hover :is(.ao-user-tag, .ao-news, .ao-new-preview, .ao-notify) { filter: var(--ao-hover-inner-filter); }
#ao-root .ao-new-ticket:hover { background: var(--ao-hover); filter: var(--ao-hover-filter); box-shadow: var(--ao-call-glow-hover); }
/* pulsantini della barra e candidati: lo sfondo al passaggio è nelle loro regole; i pulsanti normali solo si schiariscono */
#ao-root .ao-bar-btn:hover, #ao-root .ao-candidate:hover { filter: var(--ao-hover-filter); box-shadow: var(--ao-hover-glow); }
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
    // Fuori dall'area AMP, solo mentre è aperta (pannello di sviluppo): intestazione dell'add-in più bassa, con le schede
    // più vicine al logo. Nel vetro lo sfondo sfumato del pannello (fisso, quindi senza giunture) passa anche dietro
    // l'intestazione, che diventa di vetro leggero (stesso blu, appena trasparente, piatto).
    const glassHeaderRule = look => `background: ${look.background}; box-shadow: ${look.light};`;
    const OUTSIDE_STYLES = `
:has(> .app-header):has(#amp-view:not([hidden])) > .app-header { padding-bottom: 6px; }
:has(> .app-header):has(#amp-view:not([hidden]) #ao-root[data-style="glass"]) { background: ${THEME.glassBackdrop}; }
:has(> .app-header):has(#amp-view:not([hidden]) #ao-root[data-style="glass"][data-theme="dark"]) { background: ${THEME.glassDarkBackdrop}; }
@media (prefers-color-scheme: dark) { :has(> .app-header):has(#amp-view:not([hidden]) #ao-root[data-style="glass"]:not([data-theme="light"])) { background: ${THEME.glassDarkBackdrop}; } }
:has(> .app-header):has(#amp-view:not([hidden]) #ao-root[data-style="glass"]) > .app-header { -webkit-backdrop-filter: ${THEME.glassColors.navBlur}; backdrop-filter: ${THEME.glassColors.navBlur}; ${glassHeaderRule(THEME.glassHeader.light)} }
:has(> .app-header):has(#amp-view:not([hidden]) #ao-root[data-style="glass"][data-theme="dark"]) > .app-header { ${glassHeaderRule(THEME.glassHeader.dark)} }
@media (prefers-color-scheme: dark) { :has(> .app-header):has(#amp-view:not([hidden]) #ao-root[data-style="glass"]:not([data-theme="light"])) > .app-header { ${glassHeaderRule(THEME.glassHeader.dark)} } }
`;
    function injectStyles() {
        if (document.getElementById('ao-styles')) return;
        document.head.append(h('style', { id: 'ao-styles' }, themeCssVariables() + STYLES + OUTSIDE_STYLES));
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
