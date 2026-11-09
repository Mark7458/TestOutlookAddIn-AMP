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
            navInk: '#fff', // scritte e icone dell'intestazione e delle schede
            navInkSoft: 'rgba(255, 255, 255, .72)', // schede non scelte
            chipLine: 'transparent', // bordino di etichette e stati dentro le schede (nel vetro chiaro serve a staccarli dal bianco)
            hoverGlow: '0 0 transparent', // bagliore intorno alla superficie al passaggio del mouse (nessuno)
            raisedShadow: '0 6px 16px rgba(0, 0, 0, .12), 0 1px 3px rgba(0, 0, 0, .08)', // scheda sollevata (in viaggio, al clic)
            openShadow: 'var(--ao-raised-shadow)', // scheda aperta (In coda, ticket nel thread, Impostazioni): nel vetro bagliore + ombra leggera, come + Nuovo ticket aperto
            itemHover: 'var(--ao-surface)', // scheda al passaggio del mouse: non si schiarisce, si accende il bordo (e nel vetro il bagliore)
            itemHoverEdge: 'color-mix(in srgb, var(--ao-accent) 35%, transparent)', // flat: bordo blu sottile e tenue
            itemHoverGlow: '0 0 transparent',
            quoteBg: 'var(--ao-bg)', // description come citazione incassata dentro la scheda
            segmentTrack: 'var(--ao-surface)', // fondo dei selettori a segmenti
            segmentHover: 'transparent', // voce non scelta di un selettore al passaggio del mouse
            choiceTrack: 'var(--ao-bg)', // fondo del selettore delle azioni (Prendi in carico / Inoltra / Ignora)
            selection: 'var(--ao-accent-bg)', // voce scelta nei selettori a segmenti
            selectionLine: 'transparent', // bordo della voce scelta
            selectionShadow: '0 0 transparent', // ombra della voce scelta
            selectionTone: '100%', // quanto colore dell'azione nella voce scelta del selettore delle azioni
            selectionMotion: '.32s cubic-bezier(.3, .8, .2, 1)', // come scivola la voce scelta
            menuBg: 'var(--ao-surface)', // elenchi aperti: sempre pieni, si leggono sopra a tutto
            // + Nuovo ticket e ＋ Ticket (thread): nel flat bordo blu sottile e tenue, pieno al passaggio del mouse; nel vetro
            // card come le altre con un bagliore fisso che "chiama": al passaggio del mouse non si schiarisce, il bagliore
            // dà un impulso breve e misurato e resta un po' più forte; aperto tiene il bagliore
            callBorder: '1px solid color-mix(in srgb, var(--ao-accent) 30%, transparent)',
            callOpenBorder: 'var(--ao-accent)',
            callGlow: '0 0 transparent',
            lockLine: 'var(--ao-accent)', // azione confermata: nel Flat un bordo (nel vetro basta il bagliore)
            lockGlow: '0%', // intensità del bagliore dell'azione confermata (nel flat niente)
            favBg: 'var(--ao-quote-bg)', // impostazioni suggerite nel modulo: pillole più scure del fondo (nel flat si confondevano)
            callGlowHover: '0 0 transparent',
            callGlowPulse: '0 0 transparent',
            callEdgeHover: 'var(--ao-accent)',
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
            chipLine: 'rgba(0, 50, 110, .16)',
            notifyInk: '#fff', // vetro: superficie trasparente, il numero va su un colore pieno
            hoverGlow: '0 0 12px rgba(120, 170, 255, .2)',
            itemHover: 'var(--ao-surface)', // vetro: la scheda non si schiarisce; bordo vicino al colore della scheda e alone azzurro
            itemHoverEdge: 'rgba(255, 255, 255, .95)',
            itemHoverGlow: '0 0 16px rgba(110, 165, 255, .38)',
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
            lockLine: 'transparent',
            lockGlow: '55%',
            openShadow: '0 0 18px rgba(80, 150, 255, .45), 0 6px 16px rgba(0, 0, 0, .1)',
            favBg: 'var(--ao-surface)',
            callGlowHover: '0 0 24px rgba(80, 150, 255, .66), 0 0 6px rgba(80, 150, 255, .38)',
            callGlowPulse: '0 0 28px rgba(80, 150, 255, .78), 0 0 7px rgba(80, 150, 255, .48)',
            callEdgeHover: 'var(--ao-line)', // il bordo resta com'è
            searchBg: 'var(--ao-hover)', // ricerca aperta: accesa come una card al passaggio del mouse, senza contorno
            searchLensBg: 'var(--ao-hover)',
            searchLine: 'transparent',
            searchFocusLine: 'transparent',
            fieldFocusLine: 'rgba(120, 170, 255, .45)', // vetro: un bagliore soffuso al posto dell'anello
            fieldFocusGlow: '0 0 6px rgba(120, 170, 255, .14)', // campo attivo: appena accennato (l'alone forte è per attirare l'attenzione)
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
            navInk: 'rgba(236, 240, 247, .93)', // vetro scuro: bianco meno acceso
            navInkSoft: 'rgba(236, 240, 247, .64)',
            chipLine: 'transparent',
            notifyInk: '#23262e',
            hoverGlow: '0 0 12px rgba(90, 150, 255, .18)',
            itemHoverEdge: 'rgba(255, 255, 255, .2)',
            itemHoverGlow: '0 0 16px rgba(90, 150, 255, .3)',
            segmentTrack: 'rgba(28, 31, 38, .62)',
            segmentHover: 'rgba(255, 255, 255, .06)',
            selectionShadow: '0 1px 2px rgba(0, 0, 0, .3)',
            menuBg: '#2b2e36',
            fieldFocusLine: 'rgba(90, 150, 255, .4)',
            fieldFocusGlow: '0 0 6px rgba(90, 150, 255, .12)',
            callGlow: '0 0 18px rgba(70, 140, 255, .5), 0 0 4px rgba(70, 140, 255, .25)',
            callGlowHover: '0 0 24px rgba(70, 140, 255, .66), 0 0 6px rgba(70, 140, 255, .38)',
            callGlowPulse: '0 0 28px rgba(70, 140, 255, .78), 0 0 7px rgba(70, 140, 255, .48)',
            openShadow: '0 0 18px rgba(70, 140, 255, .45), 0 6px 16px rgba(0, 0, 0, .3)',
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
            ignore: 'Ignora',
            messageOf: (index, total, when) => `${index} di ${total} · ${when}`,
            previousMessage: 'Messaggio precedente',
            nextMessage: 'Messaggio successivo'
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
            newMessagesTitle: n => (n === 1 ? '1 messaggio nuovo nel thread' : `${n} messaggi nuovi nel thread`),
            holdExpired: 'On Hold scaduto',
            holdUntil: date => `fino al ${date}`,
            holdExpiredOn: date => `scaduto il ${date}`
        },
        mail: {
            outcomeWorking: 'Apro il ticket…',
            outcomeTaken: number => `${number} preso in carico`,
            noMail: 'Apri una mail in Outlook, oppure scegli una richiesta dalla Coda o un ticket.',
            searching: 'Ricerca del ticket su ServiceNow…',
            openInOutlook: 'Apri la mail',
            openFirst: 'Apri la mail in Outlook per prenderla in carico o ignorarla.',
            // Collegamento con il ticket
            closedTickets: n => `Ticket chiusi (${n})`,
            addTicket: '＋ Ticket',
            threadCount: count => (count === 1 ? '1 ticket' : `${count} ticket · 1 principale`),
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
            assignedToMe: name => `Assegnato a ${name} (me)`,
            notAssigned: 'Non assegnato',
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
            opened: 'Aperto',
            reported: 'Reported',
            openInServiceNow: 'Apri in ServiceNow',
            timeWorked: 'Tempo lavorato',
            minutes: n => `${n} min`,
            timeChanged: (delta, total) => `Tempo ${delta > 0 ? '+' : '−'}${formatMinutes(Math.abs(delta))} · totale ${formatMinutes(total)}.`,
            closed: state => `Ticket ${state}: nessuna azione disponibile.`,
            noLinkedMail: 'Nessuna mail collegata: la risposta non è disponibile da qui.',
            // Azione all'invio
            reply: 'Attività',
            updateText: 'Testo',
            updateButtons: { comment: 'Aggiorna', resolve: 'Risolvi', hold: 'Metti in On Hold', cancel: 'Annulla il ticket' },
            updated: number => `${number} aggiornato`,
            closedTicket: number => `${number} chiuso`,
            actions: { comment: 'Aggiorna', resolve: 'Risolvi', hold: 'On Hold', cancel: 'Annulla' },
            resolutionCode: 'Resolution code',
            holdReason: 'On hold reason',
            holdExpiration: 'Scadenza',
            replyText: 'Rispondi con',
            replyEmpty: 'Vuoto',
            replyCombination: 'Risposta di Agenzia + Attività',
            reference: number => `Rif. ${number}`,
            includeReference: number => `Aggiungi “Rif. ${number}” in fondo alla risposta`,
            prepareReply: 'Salva e compila',
            pendingCleared: 'Azione annullata.',
            changeChoice: 'Cambia selezione',
            timeInvalid: 'Scrivi per esempio 2h, 1h30, +2h o 120',
            pendingSaved: 'Azione salvata: verrà eseguita all’invio della risposta.',
            replyOpened: 'Risposta aperta in Outlook.'
        },
        settings: {
            title: 'Impostazioni',
            tabs: { profile: 'Profilo', companies: 'Agenzie', activities: 'Attività', responses: 'Risposte' },
            newItem: { companies: '+ Nuova agenzia', activities: '+ Nuova attività', responses: '+ Nuova risposta' },
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
            category: 'Category',
            impact: 'Impact',
            urgency: 'Urgency',
            timeWorkedMinutes: 'Minuti',
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
    // Tempo lavorato: passo di − e +; tenendo premuto il passo cresce (dopo alcune ripetizioni).
    const TIME_STEP = { base: 5, faster: 15, fastest: 30, after: [5, 12], repeat: 140, hold: 400 };
    const formatMinutes = minutes => (minutes >= 60 ? `${Math.floor(minutes / 60)} h${minutes % 60 ? ` ${minutes % 60} min` : ''}` : `${minutes} min`);
    const DEFAULT_HOLD_DAYS = 2; // scadenza proposta per l'On Hold

    // ======================================================================
    // 2. CONFIGURAZIONI (salvate nel browser, stesso formato dello script Tampermonkey)
    // Un file esportato da Tampermonkey si importa qui senza modifiche; in più le Agenzie
    // hanno i domini email e il profilo ha il tag di Outlook.
    // ======================================================================

    const STORAGE = { key: 'snow_ticket_assistant_v1', version: 3 };
    // Versione del pannello (si vede in Impostazioni → Profilo): serve a capire se il browser usa l'ultima.
    const BUILD = '2026-11-09 14:30';

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
    //   listThreadMessages(conversationId)  → [{ messageId, sender, receivedAt, preview }]: messaggi ricevuti del thread, dal più vecchio
    //                                          (In coda: con più messaggi l'operatore sceglie da quale aprire il ticket)
    //   forwardMail(conversationId, gruppo, utente) → inoltro della richiesta a un gruppo (ES, PEL, PO); il thread lascia In coda
    //                                          (facoltativo: senza, Inoltra avvisa che i servizi non sono collegati)
    //   labelMail(conversationId, utente, colore) → { ok, label: { name, color, at } }: etichetta di chi gestisce la mail
    //                                          (una per thread: se c'è già quella di un altro, ok = false e torna quella)
    //   unlabelMail(conversationId, utente)  → toglie la propria etichetta
    //   updateTicket(numero, azione, testo, utente) → Ticket: ticket senza mail o collegato (non principale), aggiornato direttamente su ServiceNow
    //                                          (azione come setPendingAction, o null = solo commento; testo = commento / Resolution notes)
    //   labelMail(…, { force: true })        → sostituisce l'etichetta di un collega (tenendo premuta la sua, a scheda aperta)
    //   (listQueue e listIgnored: ogni mail ha label = { name, color } | null, che si toglie quando la mail viene
    //    ignorata o inoltrata; i ticket senza mail hanno label di chi li ha aperti)
    //
    // mail   = { conversationId, itemId, subject, sender: { name, email }, receivedAt, body, incNumbers }
    // utente = { displayValue, value: sys_id }   (il profilo del pannello)
    // Ticket = { number, sysId, url, state, assignedTo: { name, sysId }, companyName, activityName,
    //            shortDescription, description, reportedOn, openedAt, closedAt, timeWorkedMinutes, holdReason,
    //            holdExpiration, conversationId, messageId, pendingAction, owner, newMessages,
    //            lastNewMessage: { from, at, preview } | null, lastMessageAt, syncedAt }
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
    // Pagina del thread precaricata passando sopra una scheda: il risultato resta valido qualche secondo e la pagina,
    // al clic, compare subito (poi si aggiorna comunque in silenzio).
    const threadCache = new Map();
    const THREAD_CACHE_MS = 15000;
    Services.prefetchThread = conversationId => {
        const cached = threadCache.get(conversationId);
        if (!conversationId || (cached && Date.now() - cached.at < THREAD_CACHE_MS)) return;
        const entry = { at: Date.now(), result: null, request: null };
        threadCache.set(conversationId, entry);
        entry.request = Services.findTicketForMail({ conversationId, incNumbers: [] });
        entry.request.then(result => { entry.result = result; }).catch(() => threadCache.delete(conversationId));
    };
    // Richiesta di precaricamento ancora in corso: la pagina aspetta quella invece di farne un'altra.
    Services.pendingThread = conversationId => {
        const cached = threadCache.get(conversationId);
        return cached && !cached.result && Date.now() - cached.at < THREAD_CACHE_MS ? cached.request : null;
    };
    Services.cachedThread = conversationId => {
        const cached = threadCache.get(conversationId);
        return cached?.result && Date.now() - cached.at < THREAD_CACHE_MS ? cached.result : null;
    };
    Services.available = () => SERVICE_METHODS.every(name => typeof window.AmpBackend?.[name] === 'function');
    Services.subscribe = listener => window.AmpBackend?.subscribe?.(listener) || (() => {});
    Services.getMailBody = async messageId => (typeof window.AmpBackend?.getMailBody === 'function' ? window.AmpBackend.getMailBody(messageId) : null);
    Services.listThreadMessages = async conversationId => (typeof window.AmpBackend?.listThreadMessages === 'function' ? window.AmpBackend.listThreadMessages(conversationId) : null);
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
    const CHEVRON_LEFT_ICON = '<svg viewBox="0 0 20 20" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12.5 5l-5 5 5 5"/></svg>';
    const CHEVRON_RIGHT_ICON = '<svg viewBox="0 0 20 20" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M7.5 5l5 5-5 5"/></svg>';
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
    // Spazi tra icona, nome e contatore della scheda aperta (usati dal CSS e dalla sottolineatura).
    const TAB_GAP = { label: 9, counter: 5 };
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
        newTicketShape: null, // forma del pulsante "+ Nuovo ticket" (il modulo ci torna con Annulla)
        newItemShape: null, // lo stesso per "+ Nuova agenzia/attività/risposta" in Impostazioni
        liftConfig: null, // voce di Impostazioni appena aperta: si disegna ancora a terra, poi si solleva
        ticketDrafts: {}, // per conversazione: Agenzia, Attività, Short description
        actionDrafts: {}, // per ticket: azione all'invio
        editing: null, // impostazioni: { kind, id }
        settingsTab: 'profile', // profile | companies | activities | responses
        backTo: null, // vista a cui torna il pulsante ← (Mail e Impostazioni aperte da un'altra vista)
        openedTicket: null, // ticket aperto dall'elenco Ticket: con ← la sua scheda torna al suo posto
        journey: null, // ticket aperto dalla lista mentre le schede viaggiano: nella pagina i loro posti restano vuoti finché non arrivano
        liftTicket: null, // ticket del thread appena aperto: si disegna ancora a terra, poi si solleva
        ignored: [], // archivio dei thread ignorati (In coda, in fondo)
        queueTab: 'active', // In coda: active (da prendere) | ignored (archivio)
        extraTicketFor: null, // thread su cui si sta aprendo un ticket in più
        ticketExpanded: {}, // ticket aperti o chiusi a mano nella pagina del thread (altrimenti: estesa se è l'unico)
        closedOpen: null, // thread di cui si vedono i ticket chiusi
        barOffset: 0, // di quanti px la barra secondaria è scivolata via (segue lo scorrimento)
        queueExpanded: null,
        queueChoice: {}, // scheda di In coda aperta: azione scelta nel selettore e gruppo dell'inoltro
        liftFrom: null, // scheda che si sta aprendo: disegnata non sollevata, poi si solleva
        previewScroll: null, // scheda aperta del tutto: il testo della mail può scorrere
        mailBodies: {}, // testo delle mail per l'anteprima in In coda (per messageId)
        threadMessages: {}, // messaggi ricevuti di un thread in In coda (per conversationId): con più di uno si sceglie da quale aprire il ticket
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
    }
    // Triangolazione per la mail mostrata nella vista Mail.
    async function refreshTarget(quiet = false) {
        if (state.target?.kind !== 'mail' || !Services.available()) return;
        const sequence = ++searchSequence;
        const cached = Services.cachedThread(state.target.mail.conversationId);
        if (cached && !state.result) { state.result = cached; quiet = true; } // precaricata: subito, poi si aggiorna in silenzio
        if (!quiet) {
            state.searching = true;
            render();
        }
        try {
            const result = await (Services.pendingThread(state.target.mail.conversationId) || Services.findTicketForMail(state.target.mail));
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
    // Dall'elenco alla pagina del ticket e ritorno: un viaggio della scheda (o del mazzo).
    // Andata: la scheda si stacca (si solleva, il resto si sfoca), viaggia fino al posto della scheda del ticket
    // nella pagina mentre il resto sparisce, si posa; poi la pagina si costruisce intorno (sopra scende, sotto sale).
    // Ritorno: l'esatto inverso. Durante il viaggio si muove una copia della scheda, sopra a tutto.
    const JOURNEY = {
        lift: 220, travel: 420, land: 200, build: 280, stagger: 50,
        easing: 'cubic-bezier(.3, .8, .2, 1)', travelEasing: 'cubic-bezier(.45, 0, .2, 1)',
        lifted: 'translateY(-3px) scale(1.035)', liftShadow: '0 12px 28px rgba(0, 0, 0, .18)',
        floated: 'translateY(-6px) scale(1.045)', float: 700 // se la pagina tarda: la scheda continua piano a salire, mai ferma
    };
    // Scheda del ticket nella pagina (arriva poco dopo, con la ricerca della mail): si aspetta senza bloccare il ridisegno.
    const headOf = number => [...root.querySelectorAll('.ao-scroll .ao-ticket-card .ao-item-number')]
        .find(element => element.textContent === number)?.closest('.ao-ticket-card');
    async function pageCardOf(number) {
        const find = () => headOf(number);
        for (let waited = 0; !find() && waited < 900; waited += 30) await wait(30);
        return find();
    }
    const frontOf = card => (card.classList.contains('ao-deck') ? card.querySelector(':scope > .ao-item') : card);
    // Copia della scheda che viaggia sopra a tutto, partendo dalla posizione dell'elemento dato.
    function travelCopy(source, at) {
        const box = at.getBoundingClientRect();
        const copy = source.cloneNode(true);
        copy.classList.add('ao-journey');
        Object.assign(copy.style, { left: `${box.left}px`, top: `${box.top}px`, width: `${box.width}px`, visibility: 'visible' });
        root.append(copy);
        placeAt(copy, box.top, box.left);
        return copy;
    }
    // Mette la copia esattamente a quelle coordinate dello schermo (il riferimento può essere il pannello, non la finestra).
    function placeAt(copy, top, left = copy.getBoundingClientRect().left) {
        const now = copy.getBoundingClientRect();
        copy.style.top = `${parseFloat(copy.style.top) + top - now.top}px`;
        copy.style.left = `${parseFloat(copy.style.left) + left - now.left}px`;
    }
    // La scheda si solleva prima di partire.
    const lift = element => element.animate([{ transform: 'none' }, { transform: JOURNEY.lifted, boxShadow: JOURNEY.liftShadow }],
        { duration: JOURNEY.lift, easing: JOURNEY.easing, fill: 'forwards' }).finished;
    // Copia della sola testata di un ticket della pagina (senza description né Attività), ferma nel punto dato.
    function faceCopy(card, at) {
        const copy = card.cloneNode(true);
        copy.querySelector('.ao-ticket-more')?.remove();
        copy.classList.remove('ao-away', 'ao-expanded', 'ao-unlifted');
        copy.classList.add('ao-journey');
        Object.assign(copy.style, { left: `${at.left}px`, top: `${at.top}px`, width: `${at.width}px`, visibility: 'visible', clipPath: '' });
        root.append(copy);
        placeAt(copy, at.top, at.left);
        return copy;
    }
    // Mazzo: prima di partire le fasce dietro rientrano sotto la scheda davanti (fold); all'arrivo nella lista riescono.
    function deckFold(copy, fold) {
        const front = copy.querySelector(':scope > .ao-item');
        return Promise.all([...copy.querySelectorAll('.ao-deck-strip')].map(strip => {
            strip.style.opacity = '';
            const down = `translateY(${front.offsetTop - strip.offsetTop}px)`;
            return strip.animate(fold ? [{ transform: 'none', opacity: 1 }, { transform: down, opacity: 0 }] : [{ transform: down, opacity: 0 }, { transform: 'none', opacity: 1 }],
                { duration: fold ? JOURNEY.lift : JOURNEY.land, easing: JOURNEY.easing, fill: 'forwards' }).finished;
        }));
    }
    // Le parti della scheda della lista che nel thread non ci sono (anteprima del messaggio, termine): si chiudono
    // durante il viaggio di andata (fold) e si riaprono in quello di ritorno.
    function listExtras(copy, fold) {
        return Promise.all([...copy.querySelectorAll('.ao-new-preview, .ao-item-meta')].map(part => {
            const style = getComputedStyle(part); // misure a scheda aperta (prima di chiuderla)
            const open = { height: `${part.offsetHeight}px`, opacity: 1, marginTop: style.marginTop, paddingTop: style.paddingTop, paddingBottom: style.paddingBottom };
            const shut = { height: '0px', opacity: 0, marginTop: '0px', paddingTop: '0px', paddingBottom: '0px' };
            part.style.overflow = 'hidden';
            return part.animate(fold ? [open, shut] : [shut, open], { duration: JOURNEY.travel * .7, easing: JOURNEY.easing, fill: 'forwards' }).finished;
        }));
    }
    // Dopo che le schede si sono posate la pagina si apre: sopra scendono "←" e l'intestazione della mail, tutte le
    // schede si allungano insieme da sotto la testata mostrando il resto, poi compare quel che segue (ticket chiusi…).
    // inward = false: l'esatto inverso (tutto si richiude nelle testate).
    function unfoldPage(view, inward) {
        const parts = [...(view?.children || [])].filter(child => !child.matches('.ao-fixed:not(.ao-pinned)')); // la testata del thread si apre con la pagina
        const cards = parts.filter(child => child.classList.contains('ao-ticket-card'));
        const first = cards.length ? parts.indexOf(cards[0]) : parts.length;
        const above = parts.slice(0, first);
        const after = parts.slice(first).filter(child => !cards.includes(child));
        const timing = (delay = 0) => ({ duration: JOURNEY.build, delay, easing: JOURNEY.easing, fill: inward ? 'backwards' : 'forwards' });
        const play = (element, from, to, delay) => element.animate(inward ? [from, to] : [to, from], timing(delay)).finished;
        const moves = [];
        above.forEach((element, index) => moves.push(play(element, { opacity: 0, transform: 'translateY(-8px)' }, { opacity: 1, transform: 'none' }, (inward ? index : above.length - 1 - index) * JOURNEY.stagger)));
        cards.forEach(card => {
            const more = card.querySelector('.ao-ticket-more');
            if (!more) return;
            const cardRect = card.getBoundingClientRect(), faceRect = card.querySelector('.ao-ticket-face').getBoundingClientRect();
            const closed = `inset(0px 0px ${cardRect.bottom - faceRect.bottom - 8}px 0px round var(--ao-radius-base))`;
            moves.push(play(card, { clipPath: closed }, { clipPath: 'inset(0px 0px 0px 0px round var(--ao-radius-base))' }, JOURNEY.stagger));
            const inner = [...more.children];
            inner.forEach((element, index) => moves.push(play(element, { opacity: 0, transform: 'translateY(-10px)' }, { opacity: 1, transform: 'none' },
                inward ? JOURNEY.stagger * 2 + index * JOURNEY.stagger : (inner.length - 1 - index) * JOURNEY.stagger)));
        });
        after.forEach(element => moves.push(play(element, { opacity: 0 }, { opacity: 1 }, inward ? JOURNEY.build : 0)));
        return Promise.all(moves).catch(() => {});
    }
    const DECK_STRIP = 22; // quanto spunta ogni fascia del mazzo sopra la scheda davanti (come in .ao-deck-strip)
    const DEAL_STEP = 60; // distacco tra una carta e l'altra quando il mazzo si distribuisce nella pagina
    const blurTo = (elements, from, to, duration) => Promise.all(elements.map(element => element.animate(
        [{ filter: `blur(${from[0]}px)`, opacity: from[1] }, { filter: `blur(${to[0]}px)`, opacity: to[1] }], { duration, fill: 'forwards' }).finished));
    // Viaggio e atterraggio in un solo movimento continuo, eseguito dalla scheda grafica (fluido anche a PC occupato).
    // to: { dx, dy } rispetto al punto in cui la copia è stata messa; from: com'è al via (di base già sollevata;
    // 'none': si solleva partendo, in un solo movimento).
    function travelAndLand(copy, to, delay = 0, from = JOURNEY.lifted) {
        const total = JOURNEY.travel + JOURNEY.land;
        copy.getAnimations().forEach(animation => animation.cancel()); // prende il posto del sollevamento (stessi valori)
        return copy.animate([
            { transform: from, boxShadow: from === 'none' ? '0 0 0 transparent' : JOURNEY.liftShadow, easing: JOURNEY.travelEasing },
            { transform: `translate(${to.dx}px, ${to.dy}px) ${JOURNEY.lifted}`, boxShadow: JOURNEY.liftShadow, offset: JOURNEY.travel / total, easing: JOURNEY.easing },
            { transform: `translate(${to.dx}px, ${to.dy}px)`, boxShadow: '0 0 0 transparent' }
        ], { duration: total, delay, fill: 'both' }).finished;
    }
    // Copia ferma della lista, sopra la pagina che intanto si prepara: si sfoca e svanisce intorno alla scheda.
    function listGhost(scroll, card) {
        const box = scroll.getBoundingClientRect();
        const ghost = scroll.cloneNode(true);
        ghost.classList.add('ao-ghost');
        Object.assign(ghost.style, { left: `${box.left}px`, top: `${box.top}px`, width: `${box.width}px`, height: `${box.height}px` });
        // la scheda che parte non c'è nella copia (stessa posizione nell'albero)
        const path = [];
        for (let node = card; node !== scroll; node = node.parentElement) path.unshift([...node.parentElement.children].indexOf(node));
        root.append(ghost);
        ghost.scrollTop = scroll.scrollTop;
        placeAt(ghost, box.top, box.left);
        const twin = path.reduce((node, index) => node?.children[index], ghost);
        if (twin) twin.style.visibility = 'hidden';
        return ghost;
    }
    // card: la scheda o il mazzo nella lista; ticket: quello da aprire (nel mazzo, il principale, davanti).
    // La pagina si prepara nascosta mentre la scheda si stacca. Poi ogni ticket del thread scivola dal mazzo al suo
    // posto nella pagina, tutti insieme (un attimo di distacco tra uno e l'altro): chi finisce fuori dalla vista ci va
    // e basta. Quando sono tutti posati, si aprono insieme. Con un ticket solo è lo stesso movimento, con una scheda.
    async function openFromCard(card, ticket) {
        resetTicketFilters();
        if (state.holdRender || !card || reducedMotion()) return openTicketItem(ticket);
        const body = root.querySelector('.ao-body');
        const start = frontOf(card).getBoundingClientRect(); // nel mazzo conta la scheda davanti
        const copy = travelCopy(card, card);
        const ghost = listGhost(root.querySelector('.ao-scroll'), card);
        body.style.opacity = '0';
        state.journey = ticket.number; // nella pagina le schede restano vuote finché non arrivano
        openTicketItem(ticket, { slide: false });
        blurTo([ghost], [0, 1], [3, .55], JOURNEY.lift).then(() => blurTo([ghost], [3, .55], [5, 0], JOURNEY.travel * .5)).catch(() => {});
        const ready = pageCardOf(ticket.number);
        let page = null;
        ready.then(found => { page = found; });
        await Promise.all([lift(copy), deckFold(copy, true)]).catch(() => {});
        await wait(0);
        // Pagina non ancora pronta (dati in arrivo): la scheda continua piano a salire invece di fermarsi.
        const float = page ? null : copy.animate([{ transform: JOURNEY.lifted, boxShadow: JOURNEY.liftShadow }, { transform: JOURNEY.floated, boxShadow: JOURNEY.liftShadow }],
            { duration: JOURNEY.float, easing: 'cubic-bezier(.2, .6, .3, 1)', fill: 'forwards' });
        page = await ready;
        const floated = float ? getComputedStyle(copy).transform : JOURNEY.lifted;
        const cards = page ? [...page.closest('.ao-scroll').querySelectorAll(':scope > .ao-ticket-card')] : [];
        const from = cards.indexOf(page);
        const copies = [];
        await holdingRender(() => Promise.all([listExtras(copy, true), ...cards.map((target, index) => {
            const place = target.getBoundingClientRect();
            const traveller = target === page ? copy : faceCopy(target, start); // gli altri partono da sotto quella davanti
            if (traveller !== copy) { traveller.style.zIndex = String(29 - Math.abs(index - from)); copies.push(traveller); }
            return travelAndLand(traveller, { dx: place.left - start.left, dy: place.top - start.top }, Math.abs(index - from) * DEAL_STEP, traveller === copy ? floated : JOURNEY.lifted);
        })]).catch(() => {}));
        state.journey = null;
        state.liftTicket = ticket.number; // si posa a terra, si solleva dopo essersi aperta
        render();
        state.liftTicket = null;
        const target = headOf(ticket.number);
        body.style.opacity = '';
        [ghost, copy, ...copies].forEach(element => element.remove());
        if (!target) return;
        await holdingRender(() => unfoldPage(target.closest('.ao-scroll'), true));
        target.classList.remove('ao-unlifted'); // aperta (ticket singolo): ora si solleva, come in In coda
    }
    // Ritorno. Con più ticket: il principale (in fondo) si solleva e sale fino al primo, e ogni collegato che incontra
    // gli scivola sotto e viaggia con lui (il mazzo si ricompone); intanto tutte le schede si richiudono nelle testate.
    // Poi il mazzo torna al suo posto nella lista in un solo movimento, mentre la lista ricompare.
    // Con un ticket solo: la scheda si richiude e torna nella lista.
    async function closeToList(target) {
        const number = state.openedTicket;
        state.openedTicket = null;
        if (state.holdRender || !number || target !== 'tickets' || reducedMotion()) return setView(target);
        const pageCard = await pageCardOf(number);
        if (!pageCard) return setView(target);
        const view = pageCard.closest('.ao-scroll');
        pageCard.classList.add('ao-unlifted'); // torna a terra mentre la pagina si richiude
        const cards = [...view.querySelectorAll(':scope > .ao-ticket-card')];
        const head = pageCard.getBoundingClientRect();
        const others = cards.filter(item => item !== pageCard)
            .map(item => ({ item, top: item.getBoundingClientRect().top }))
            .sort((a, b) => b.top - a.top) // dal più vicino al principale
            .map((entry, index) => ({ ...entry, step: index + 1 }));
        // come nel mazzo della lista: ogni collegato raccolto spunta sopra il principale di una fascia, più stretto
        const climb = others.length ? Math.min(...others.map(entry => entry.top + DECK_STRIP * entry.step)) - head.top : 0;
        // mosse fotogramma per fotogramma: senza le transizioni delle schede, che le farebbero rincorrere
        Object.assign(pageCard.style, { position: 'relative', zIndex: String(others.length + 2), transition: 'none' });
        others.forEach(entry => Object.assign(entry.item.style, { position: 'relative', zIndex: String(others.length + 1 - entry.step), transition: 'none' }));
        const gather = () => new Promise(resolve => {
            if (!others.length) return resolve();
            const duration = JOURNEY.build * 2;
            const ease = t => (t < .5 ? 2 * t * t : 1 - Math.pow(-2 * t + 2, 2) / 2);
            let first = null;
            const frame = now => {
                if (first === null) first = now;
                const t = Math.min(1, (now - first) / duration);
                const up = Math.min(1, t * 4); // si solleva nei primi istanti della salita
                const offset = climb * ease(t);
                pageCard.style.transform = `translateY(${offset - 3 * up}px) scale(${1 + .035 * up})`;
                pageCard.style.boxShadow = `0 12px 28px rgba(0, 0, 0, ${.18 * up})`;
                others.forEach(entry => {
                    if (head.top + offset > entry.top + DECK_STRIP * entry.step + 1) return; // non ancora raggiunto
                    entry.item.style.transform = `translateY(${head.top + offset - DECK_STRIP * entry.step - entry.top - 3 * up}px) scaleX(${1 - .06 * entry.step})`;
                });
                if (t < 1) requestAnimationFrame(frame);
                else resolve();
            };
            requestAnimationFrame(frame);
        });
        await holdingRender(() => Promise.all([unfoldPage(view, false), gather()]).catch(() => {}));
        const from = head.top + climb; // dove sta ora il principale (in cima)
        setView(target, null, { slide: false });
        const found = [...root.querySelectorAll('.ao-scroll .ao-item')].find(item => item.querySelector('.ao-item-number')?.textContent === number);
        const card = found?.closest('.ao-deck') || found;
        if (!card) return slideInView(-1);
        card.scrollIntoView({ block: 'nearest' });
        const scroll = root.querySelector('.ao-scroll');
        const offset = frontOf(card).getBoundingClientRect().top - card.getBoundingClientRect().top;
        const listed = card.getBoundingClientRect();
        const copy = travelCopy(card, card);
        placeAt(copy, from - offset, head.left);
        card.style.visibility = 'hidden';
        scroll.style.willChange = 'filter, opacity';
        // il mazzo è già ricomposto (le fasce visibili); una scheda sola parte da terra e si solleva partendo
        if (!others.length) copy.querySelectorAll('.ao-deck-strip').forEach(strip => { strip.style.opacity = '0'; });
        await holdingRender(async () => {
            await Promise.all([
                travelAndLand(copy, { dx: listed.left - head.left, dy: listed.top + offset - from }, 0, others.length ? JOURNEY.lifted : 'none'), listExtras(copy, false),
                blurTo([scroll], [5, 0], [3, .55], JOURNEY.travel).then(() => blurTo([scroll], [3, .55], [0, 1], JOURNEY.land))
            ]).catch(() => {});
            scroll.getAnimations().forEach(animation => animation.cancel());
            scroll.style.willChange = '';
            card.style.visibility = '';
            copy.remove();
        });
    }
    // Filtri della barra Ticket: valgono solo finché sono aperti (chiudendoli o aprendo un ticket si azzerano).
    function resetTicketFilters() {
        Object.assign(state, { ticketFilterOpen: false, ticketStatus: 'all', ticketOperator: 'mine', filterMenu: null });
    }
    // backTo: vista a cui tornare con il pulsante ← (le schede lo azzerano).
    function setView(view, backTo = null, { slide = true } = {}) {
        const from = state.view;
        state.barOffset = 0;
        state.view = view;
        state.backTo = backTo;
        state.editing = null;
        state.filterMenu = null;
        if (from !== view) state.newTicketOpen = false; // uscendo, il modulo "Nuovo ticket" si richiude (al ritorno c'è il pulsante)
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
        root.querySelector('.ao-tabs').addEventListener('transitionend', event => {
            if (event.target.classList.contains('ao-tab-label') && event.propertyName === 'max-width') settleTabIndicator();
        });
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
        closeSelects();
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
        placeTabIndicator(); // dopo i contatori: la sottolineatura comprende il numero della scheda aperta
        renderStatusBar();
        const views = { queue: renderQueue, tickets: renderTickets, mail: renderMail, settings: renderSettings };
        const body = root.querySelector('.ao-body');
        // Chi resta fermo (barra secondaria, "←") sta in alto, separato da una linea; tutto il resto scorre
        // sotto, a tutta larghezza. La posizione di scorrimento resta uguale tra un aggiornamento e l'altro.
        const scrollKey = `${state.view}|${state.target?.kind || ''}|${state.target?.number || state.target?.mail?.conversationId || ''}|${state.queueTab}|${state.ticketSource}|${state.ticketOperator}|${state.ticketSearchOpen}|${state.settingsTab}`;
        const previousScroll = body.dataset.scrollKey === scrollKey ? body.querySelector('.ao-scroll')?.scrollTop || 0 : 0;
        const view = views[state.view](config, user);
        const children = [...view.children];
        const fixedCount = children.findIndex(child => !child.matches('.ao-segmented, .ao-back-row, .ao-source-bar, .ao-thread-head'));
        const fixed = fixedCount < 0 ? children : children.slice(0, fixedCount);
        const scroll = h('div', { class: 'ao-scroll' }, ...(fixedCount < 0 ? [] : children.slice(fixedCount)));
        if (!previousScroll) state.barOffset = 0;
        // La barra è il primo elemento dell'area che scorre: la sua "casa" è in cima all'elenco.
        const bar = fixed.length ? h('div', { class: 'ao-fixed' }, ...fixed) : null;
        if (bar) scroll.prepend(bar);
        bar?.classList.toggle('ao-pinned', Boolean(bar.querySelector('.ao-thread-head'))); // il thread resta sempre visibile
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
        if (bar.classList.contains('ao-pinned')) {
            // Testata del thread: non scivola mai via; scorrendo si stringe (oggetto su una riga) e prende l'ombra.
            const pin = () => {
                const top = scroll.scrollTop;
                if (top > 24) bar.classList.add('ao-scrolled');
                else if (top <= 0) bar.classList.remove('ao-scrolled'); // margine tra le due soglie: niente sfarfallio
                bar.style.setProperty('--ao-lift', Math.min(1, top / BAR_SHADOW_FADE).toFixed(3));
            };
            pin();
            scroll.addEventListener('scroll', pin, { passive: true });
            return;
        }
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
        const messages = threadMessagesOf(item);
        const shown = messages.length > 1 ? chosenMessage(item) : null; // con più messaggi: quello che si sta guardando
        const body = state.mailBodies[shown?.messageId || item.messageId];
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
            shown ? messagePicker(item, messages, shown) : null,
            // Testo limitato a PREVIEW_MAX_CHARS; la barra di scorrimento compare solo a scheda aperta del tutto.
            h('div', { class: `ao-mail-preview${state.previewScroll === item.conversationId ? ' ao-scrollable' : ''}` },
                clipText(body ?? shown?.preview ?? item.preview ?? '', PREVIEW_MAX_CHARS), body === undefined ? h('span', { class: 'ao-hint' }, ` ${TEXTS.common.loading}`) : null),
            iconBar(actions.map(action => ({ ...action, icon: ACTION_ICONS[action.value] })), choice.action, value => pickQueueAction(item, value), { key: `action:${item.conversationId}`, tones: true, slide: false }),
            choiceZone(`queue:${item.conversationId}`, choice.action ? queueActionContent(item, mail, config, choice) : null),
            h('button', {
                type: 'button', class: 'ao-collapse', 'aria-label': TEXTS.common.close, title: TEXTS.common.close,
                onclick: event => closeTakeCharge(event.currentTarget.closest('.ao-inline-form'))
            }, svgIcon(CHEVRON_UP_ICON)));
    }
    // Messaggi del thread in In coda. Con un messaggio solo il ticket nasce da quello; con più messaggi l'operatore li
    // scorre con le frecce (si parte dall'ultimo): quello che sta guardando è quello da cui nasce il ticket.
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
    // Un altro messaggio del thread: il testo svanisce, la tendina si adatta al nuovo e questo compare (come nei ticket).
    async function showQueueMessage(item, index) {
        if (state.holdRender) return;
        const id = item.conversationId;
        const choice = state.queueChoice[id];
        const message = threadMessagesOf(item)[index];
        if (!message || !choice) return;
        const form = queueDetailsOf(id);
        const before = form?.offsetHeight || 0;
        const changing = () => [...(queueDetailsOf(id)?.querySelectorAll('.ao-message-where, .ao-mail-preview') || [])];
        const body = state.mailBodies[message.messageId] === undefined
            ? Services.getMailBody(message.messageId).then(text => text || message.preview || '').catch(() => message.preview || '') : null;
        await holdingRender(async () => {
            await Promise.all(changing().map(element => element.animate([{ opacity: 1 }, { opacity: 0 }], { duration: MOTION.hold, easing: 'ease-out', fill: 'forwards' }).finished)).catch(() => {});
            if (body) state.mailBodies[message.messageId] = await body;
            choice.at = index;
        });
        render();
        const next = queueDetailsOf(id);
        changing().forEach(element => element.animate([{ opacity: 0 }, { opacity: 1 }], { duration: MOTION.swap, easing: 'ease-in' }));
        if (next && before && next.scrollHeight !== before) await holdingRender(() => playOn(next, [{ maxHeight: `${before}px` }, { maxHeight: `${next.scrollHeight}px` }], MOTION.swap)
            .then(result => { if (result) { result.animation.cancel(); next.style.overflow = result.overflow; } }));
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
    function mailFromQueueItem(item) {
        const message = chosenMessage(item); // il ticket nasce dal messaggio che si sta guardando (o dall'unico)
        return {
            conversationId: item.conversationId, itemId: message.messageId || item.messageId, subject: item.subject, sender: message.sender || item.sender,
            receivedAt: message.receivedAt || item.receivedAt, body: state.mailBodies[message.messageId] || message.preview || item.preview, incNumbers: []
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
        // Ordine: in cima gli On Hold scaduti (vanno ripresi subito), poi per ultima attività (apertura o ultimo
        // messaggio arrivato), dalla più recente: un thread con un messaggio nuovo risale da solo.
        const activity = ticket => [String(ticket.openedAt || ''), String(ticket.lastMessageAt || '')].sort().slice(-1)[0];
        const rank = group => [group.tickets.some(isHoldExpired) ? 1 : 0, group.tickets.map(activity).sort().slice(-1)[0]];
        groups.forEach(group => { group.tickets = threadOrder(group.tickets).reverse(); }); // davanti il principale
        groups.sort((a, b) => {
            const [holdA, dateA] = rank(a);
            const [holdB, dateB] = rank(b);
            return holdB - holdA || dateB.localeCompare(dateA);
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
                ? ticketDeck(group, user, query)
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
            onclick: () => morphBar(() => { if (state.ticketFilterOpen) resetTicketFilters(); else state.ticketFilterOpen = true; state.filterMenu = null; })
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
                state.newTicketOpen = false;
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
    // "+ Nuovo ticket" si allarga al suo posto fino a diventare il modulo, poi compare il contenuto; Annulla fa
    // l'inverso (il contenuto svanisce e il modulo torna pulsante, con i suoi angoli).
    async function openNewTicket() {
        await whenIdle();
        const button = root.querySelector('.ao-new-ticket');
        state.newTicketShape = button && shape(button);
        state.newTicketOpen = true;
        render();
        const card = root.querySelector('.ao-new-ticket-card');
        if (card && state.newTicketShape && !reducedMotion()) await holdingRender(() => growFrom(card, state.newTicketShape));
    }
    async function closeNewTicket(form) {
        const card = form?.closest('.ao-new-ticket-card');
        if (card && state.newTicketShape && !reducedMotion()) await holdingRender(() => shrinkTo(card, state.newTicketShape));
        state.newTicketOpen = false;
        delete state.ticketDrafts[NEW_TICKET_KEY];
        render();
    }
    // La scheda cresce dalla forma del pulsante (altezza e angoli), poi compare il contenuto.
    async function growFrom(card, from) {
        const parts = [...card.children];
        const full = { height: `${card.offsetHeight}px`, borderRadius: getComputedStyle(card).borderTopLeftRadius };
        parts.forEach(part => { part.style.opacity = '0'; });
        card.style.overflow = 'hidden';
        await card.animate([{ height: `${from.rect.height}px`, borderRadius: from.radius }, full], { duration: JOURNEY.build + 60, easing: JOURNEY.easing }).finished.catch(() => {});
        card.style.overflow = '';
        parts.forEach(part => { part.style.opacity = ''; });
        await fadeAll(parts, 0, 1, { duration: MOTION.swap, easing: 'ease-in' });
    }
    // L'inverso: il contenuto svanisce e la scheda torna alla forma del pulsante.
    async function shrinkTo(card, to) {
        await fadeAll([...card.children], 1, 0, { duration: MOTION.hold, easing: 'ease-out', fill: 'forwards' });
        card.style.overflow = 'hidden';
        await card.animate([{ height: `${card.offsetHeight}px`, borderRadius: getComputedStyle(card).borderTopLeftRadius }, { height: `${to.rect.height}px`, borderRadius: to.radius }],
            { duration: JOURNEY.build + 60, easing: JOURNEY.easing, fill: 'forwards' }).finished.catch(() => {});
    }
    // Messaggi nuovi (numero su sfondo blu) e inizio dell'ultimo arrivato nel thread (chi scrive, quando, prime righe).
    function newMessagesPreview(ticket) {
        if (!ticket.newMessages) return null;
        const last = ticket.lastNewMessage;
        return h('div', { class: 'ao-new-preview', title: TEXTS.tickets.newMessagesTitle(ticket.newMessages) },
            h('div', { class: 'ao-new-head' },
                h('span', { class: 'ao-news' }, ticket.newMessages),
                last ? h('span', { class: 'ao-new-from' }, last.from) : null,
                last ? h('span', { class: 'ao-item-time' }, formatWhen(last.at)) : null),
            last?.preview ? h('div', { class: 'ao-new-text' }, last.preview) : null);
    }
    // Thread con più ticket: un mazzo a scaletta verso il basso. Dietro, in alto, le fasce con il numero dei ticket
    // meno recenti (più strette); davanti, completa, la scheda più recente con l'anteprima. Un clic porta al thread.
    // Ticket aperti dello stesso thread, nell'ordine in cui si lavorano: in cima i collegati, dal più recente;
    // per ultimo il principale, che si chiude per ultimo (con la risposta alla mail).
    function threadOrder(tickets) {
        const byOpening = (a, b) => String(a.openedAt || '').localeCompare(String(b.openedAt || ''));
        const primary = tickets.find(ticket => ticket.primary) || [...tickets].sort(byOpening)[0];
        return [...tickets.filter(ticket => ticket !== primary).sort((a, b) => byOpening(b, a)), primary].filter(Boolean);
    }
    // Mazzo: davanti il principale, dietro i collegati nell'ordine in cui escono nel thread (dal più vicino a lui).
    function ticketDeck(group, user, query) {
        const [top, ...rest] = group.tickets;
        const open = event => openFromCard(event.currentTarget.closest('.ao-deck'), top);
        return h('li', { class: 'ao-deck', dataset: { deck: group.conversationId } },
            rest.slice().reverse().map((ticket, index) => h('div', {
                class: 'ao-deck-strip', style: `--ao-deck-inset: ${(rest.length - index) * 3}%; --ao-deck-shade: ${(rest.length - index) * .035}`, onclick: open
            }, markText(ticket.number, query))),
            ticketItem(top, user, { query, onpick: open }));
    }
    // Testata del ticket, uguale nella lista e nel thread (è la stessa scheda che viaggia): numero, stato ed etichetta
    // dell'operatore affiancati; agenzia e attività (nel thread con il tempo lavorato a destra); short description.
    // owner: etichetta dell'operatore; takeover: un clic sull'etichetta di un collega la prepara, il secondo lo prende.
    // primary: il principale del thread (si lavora con la risposta via mail): icona della mail davanti al numero.
    function ticketHead(ticket, user, { query = '', owner = false, takeover = false, primary = false } = {}) {
        return [
            h('div', { class: 'ao-item-top' },
                h('span', { class: 'ao-item-number' }, primary ? h('span', { class: 'ao-primary-mark', 'aria-hidden': 'true' }, svgIcon(SEGMENT_ICONS.mail)) : null, markText(ticket.number, query)),
                h('div', { class: 'ao-item-tags' }, stateChip(ticket), owner ? ownerTag(ticket, user, takeover) : null)),
            h('div', { class: 'ao-item-where' },
                h('div', { class: 'ao-item-place' },
                    h('div', { class: 'ao-item-company' }, ticket.companyName || ''),
                    // Durante una ricerca i risultati arrivano da entrambi gli elenchi: si dice da quale.
                    h('div', { class: 'ao-item-activity' }, [ticket.activityName, query ? (ticket.messageId ? TEXTS.tickets.fromMail : TEXTS.tickets.withoutMail).toLowerCase() : ''].filter(Boolean).join(' · ')))),
            h('div', { class: 'ao-item-title' }, markText(ticket.shortDescription, query))
        ];
    }
    function ownerTag(ticket, user, takeover) {
        const config = loadConfig();
        const mine = isMine(ticket, user);
        const label = mine ? { name: currentUser(config).displayValue, color: config.profile.label }
            : (ticket.label || { name: ticket.assignedTo?.name || TEXTS.common.unassigned, color: 'grafite' });
        if (mine || !takeover || !OPEN_STATES.includes(ticket.state) || !currentUser(config).displayValue) return labelChip(label);
        const key = `ticket:${ticket.number}`;
        const armed = state.labelArmed === key;
        return h('button', {
            type: 'button', class: `ao-user-tag ao-tag-swap${armed ? ' ao-armed-tag' : ''}`,
            style: `--ao-tag-color: ${THEME.labels[label.color] || THEME.labels.grafite}; --ao-tag-new: ${THEME.labels[config.profile.label] || THEME.labels.azzurro}`,
            title: armed ? TEXTS.queue.labelTakeConfirm : TEXTS.queue.labelTake(label.name),
            onclick: async event => {
                event.stopPropagation();
                if (!armed) { state.labelArmed = key; render(); return; }
                state.labelArmed = null;
                const updated = await run(() => Services.assignTicket(ticket.number, config.profile.assignedTo), TEXTS.mail.taken(ticket.number));
                if (!updated) return render();
                rememberTicket(updated);
                render();
                refreshLists();
            }
        }, armed ? [h('s', {}, surname(label.name)), svgIcon(SWAP_ICON), h('span', { class: 'ao-tag-next' }, surname(currentUser(config).displayValue))] : surname(label.name));
    }
    function ticketItem(ticket, user, { inGroup = false, query = '', onpick = null } = {}) {
        const holdDate = ticket.state === '3' ? parseDate(ticket.holdExpiration) : null;
        const expired = isHoldExpired(ticket);
        const item = clickable(`ao-item${expired ? ' ao-item-alert' : ''}`, onpick || (event => openFromCard(event.currentTarget, ticket)),
            ticketHead(ticket, user, { query, owner: true }),
            inGroup ? null : newMessagesPreview(ticket),
            // In basso, solo se ci sono: termine (On Hold) e azione in attesa.
            holdDate || ticket.pendingAction ? h('div', { class: 'ao-item-meta' },
                holdDate ? h('span', { class: `ao-item-time${expired ? ' ao-item-late' : ''}` }, (expired ? TEXTS.tickets.holdExpiredOn : TEXTS.tickets.holdUntil)(formatWhen(holdDate))) : null,
                ticket.pendingAction ? chip(pendingLabel(ticket.pendingAction), 'warning') : null) : null);
        // Fermandosi sopra la scheda si precarica la sua pagina (un attimo di attesa: non ciò che il mouse attraversa)
        let hover = null;
        item.addEventListener('pointerenter', () => { hover = setTimeout(() => Services.prefetchThread(ticket.conversationId), 120); });
        item.addEventListener('pointerleave', () => clearTimeout(hover));
        item.addEventListener('pointerdown', () => Services.prefetchThread(ticket.conversationId)); // al tocco (telefono): subito
        return item;
    }

    // ---------------------------- Mail ------------------------------------
    function renderMail(config, user) {
        const notice = serviceNotice();
        const target = state.target;
        if (!target) return h('div', { class: 'ao-view' }, backRow(), notice, h('p', { class: 'ao-empty' }, TEXTS.mail.noMail));
        if (target.kind === 'ticket') {
            const ticket = ticketByNumber(target.number);
            return h('div', { class: 'ao-view' }, backRow(), notice,
                ticket ? ticketBox(ticket, config, user, { isCurrentMail: false, single: true }) : h('p', { class: 'ao-muted' }, TEXTS.common.loading));
        }
        const mail = target.mail;
        const isCurrentMail = state.mail?.conversationId === mail.conversationId;
        const result = state.result;
        const content = [];
        if (!result) content.push(h('p', { class: 'ao-muted' }, Services.available() ? TEXTS.mail.searching : ''));
        else if (result.tickets?.length) {
            // Ogni ticket aperto è un contenitore (aperto se è l'unico); i chiusi in fondo, in una sezione da aprire.
            const all = result.tickets.map(found => ({ ...found, ...(ticketByNumber(found.number) || {}) }));
            const opened = all.filter(ticket => OPEN_STATES.includes(ticket.state));
            const closed = all.filter(ticket => !OPEN_STATES.includes(ticket.state));
            const ordered = threadOrder(opened); // collegati in cima, il principale per ultimo
            // Un altro ticket sulla stessa richiesta (dal "＋ Ticket" della testata): il modulo si apre in cima alla pila,
            // dove poi comparirà il ticket nuovo (il collegato più recente).
            if (state.extraTicketFor === mail.conversationId) content.push(extraTicketCard(mail, config));
            ordered.forEach(ticket => content.push(ticketBox(ticket, config, user, { isCurrentMail, single: ordered.length === 1, primary: ticket === ordered[ordered.length - 1], others: ordered.length - 1 })));
            if (closed.length) content.push(closedTickets(mail.conversationId, closed));
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
        const count = (result?.tickets || []).filter(found => OPEN_STATES.includes((ticketByNumber(found.number) || found).state)).length;
        return h('div', { class: 'ao-view' },
            threadHead(mail, { showOpen: !isCurrentMail && Boolean(result?.tickets?.length), count }), // senza ticket c'è già il riquadro "Apri la mail"
            notice,
            content);
    }
    // Testata del thread, sempre visibile: "‹" per tornare, oggetto e mittente (scorrendo resta solo l'oggetto, su una
    // riga); sotto, con dei ticket, quanti sono e "＋ Ticket" per aprirne un altro sulla stessa richiesta.
    function threadHead(mail, { showOpen, count }) {
        const target = state.backTo || 'tickets';
        const adding = state.extraTicketFor === mail.conversationId;
        const back = TEXTS.common.back(TEXTS.tabs[target]);
        return h('div', { class: 'ao-thread-head' },
            h('div', { class: 'ao-thread-title' },
                h('button', { type: 'button', class: 'ao-thread-back', 'aria-label': back, title: back, onclick: () => closeToList(target) }, svgIcon(CHEVRON_LEFT_ICON)),
                h('div', { class: 'ao-mail-head' },
                    h('div', { class: 'ao-mail-subject' }, mail.subject || '(senza oggetto)'),
                    h('div', { class: 'ao-mail-from' },
                        h('span', {}, `${mail.sender?.name || ''}${mail.sender?.email ? ` <${mail.sender.email}>` : ''}`),
                        h('span', { class: 'ao-item-time' }, formatWhen(mail.receivedAt))),
                    showOpen ? h('button', { type: 'button', class: 'ao-link ao-mail-open', onclick: () => Outlook.openMessage(mail.itemId) }, TEXTS.mail.openInOutlook) : null)),
            count ? h('div', { class: 'ao-thread-bar' },
                h('span', { class: 'ao-thread-count' }, TEXTS.mail.threadCount(count)),
                h('button', {
                    type: 'button', class: `ao-pill${adding ? ' ao-away' : ''}`, 'aria-expanded': String(adding), // aperto: è diventato il modulo
                    onclick: () => openExtraTicket(mail)
                }, TEXTS.mail.addTicket)) : null);
    }
    // Modulo per un altro ticket del thread: lo stesso di "Nuovo ticket" in Senza mail (si solleva e si apre;
    // gli altri ticket si sfocano), con Agenzia e Attività.
    function extraTicketCard(mail, config) {
        return h('section', { class: 'ao-card ao-new-ticket-card' },
            h('div', { class: 'ao-inline-form' },
                ticketForm(mail, config, { zone: true }),
                h('div', { class: 'ao-actions' },
                    openTicketButton(mail, config),
                    h('span', { class: 'ao-spacer' }),
                    h('button', { type: 'button', class: 'ao-button ao-subtle', onclick: () => closeExtraTicket(mail) }, TEXTS.common.cancel))));
    }
    // "＋ Ticket" si trasforma nel modulo: un riquadro parte dalla pillola e si allarga fino al posto del modulo, in cima
    // alla pila; lì compare il contenuto. Annulla fa l'inverso (il modulo torna pillola). Intanto la pillola non c'è.
    // Aperto il ticket, il modulo compilato diventa la sua scheda (stesso posto) e la pillola ricompare.
    async function openExtraTicket(mail) {
        await whenIdle();
        const scroll = root.querySelector('.ao-scroll');
        if (scroll) scroll.scrollTop = 0;
        const pill = root.querySelector('.ao-thread-bar .ao-pill');
        const from = pill && shape(pill);
        state.extraTicketFor = mail.conversationId;
        render();
        const card = root.querySelector('.ao-new-ticket-card');
        if (!from || !card || reducedMotion()) return;
        await holdingRender(() => morphInto(card, from));
    }
    async function closeExtraTicket(mail) {
        if (state.holdRender) return;
        const card = root.querySelector('.ao-new-ticket-card');
        const pill = root.querySelector('.ao-thread-bar .ao-pill');
        if (card && pill && !reducedMotion()) await holdingRender(async () => {
            await fadeAll([...card.children], 1, 0, { duration: MOTION.hold, easing: 'ease-out', fill: 'forwards' });
            card.style.visibility = 'hidden';
            await morphShell(shape(card), shape(pill));
        });
        state.extraTicketFor = null;
        delete state.ticketDrafts[mail.conversationId];
        render();
        root.querySelector('.ao-thread-bar .ao-pill')?.animate([{ opacity: 0 }, { opacity: 1 }], { duration: MOTION.hold }); // il modulo è tornato pillola
    }
    // Ticket aperto dal modulo del thread: il contenuto svanisce, la pagina si aggiorna con il modulo ancora al suo
    // posto, poi il riquadro si adatta alla scheda del ticket nuovo (lì, in cima) e la pillola ricompare.
    async function extraTicketDone(mail, ticket, config, draft) {
        Usage.record(draft.companyId, draft.activityId, mail.subject);
        config.recent = [{ company: draft.companyId, activity: draft.activityId, timestamp: new Date().toISOString() },
            ...config.recent.filter(item => !(item.company === draft.companyId && item.activity === draft.activityId))].slice(0, 10);
        saveConfig(config);
        const form = root.querySelector('.ao-new-ticket-card');
        const from = form && !reducedMotion() ? shape(form) : null;
        holdRender();
        if (form && from) await fadeAll([...form.children], 1, 0, { duration: MOTION.hold, easing: 'ease-out', fill: 'forwards' });
        await refreshLists().catch(() => {});
        state.extraTicketFor = null;
        releaseRender();
        render();
        const pill = aside(root.querySelector('.ao-thread-bar .ao-pill'));
        const card = headOf(ticket.number);
        if (from && card) await holdingRender(() => morphInto(card, from));
        appear(pill);
    }
    // Forma di un elemento (posizione, misura, angoli) per il riquadro che si trasforma.
    const shape = element => {
        const rect = element.getBoundingClientRect();
        return { rect, radius: `${Math.min(parseFloat(getComputedStyle(element).borderTopLeftRadius) || 0, rect.height / 2)}px` }; // pillola: metà altezza, non 9999px
    };
    // Riquadro che va da una forma all'altra (con il bagliore del modulo); si toglie arrivato.
    async function morphShell(from, to) {
        const shell = h('div', { class: 'ao-morph' });
        root.append(shell);
        Object.assign(shell.style, { left: `${to.rect.left}px`, top: `${to.rect.top}px`, width: `${to.rect.width}px`, height: `${to.rect.height}px` });
        placeAt(shell, to.rect.top, to.rect.left);
        const frame = form => ({ transform: `translate(${form.rect.left - to.rect.left}px, ${form.rect.top - to.rect.top}px)`, width: `${form.rect.width}px`, height: `${form.rect.height}px`, borderRadius: form.radius });
        await shell.animate([frame(from), frame(to)], { duration: JOURNEY.build + 60, easing: JOURNEY.easing, fill: 'forwards' }).finished.catch(() => {});
        shell.remove();
    }
    // L'elemento prende forma da "from" (pillola o modulo): il riquadro lo raggiunge, poi compare il contenuto.
    async function morphInto(element, from) {
        element.style.visibility = 'hidden';
        await morphShell(from, shape(element));
        element.style.visibility = '';
        await fadeAll([...element.children], 0, 1, { duration: MOTION.swap, easing: 'ease-in' });
    }
    // Il pulsante "nuovo ticket" torna al suo posto dopo che il modulo è diventato il ticket.
    // Il pulsante resta nascosto mentre il modulo diventa la scheda (aside), poi ricompare (appear).
    const aside = element => { if (element) element.style.opacity = '0'; return element; };
    const appear = element => {
        if (!element) return;
        element.style.opacity = '';
        element.animate([{ opacity: 0, transform: 'scale(.9)' }, { opacity: 1, transform: 'none' }], { duration: MOTION.swap, easing: JOURNEY.easing });
    };
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
        // Corpo completo se è proprio la mail aperta in Outlook; altrimenti quello della Coda. Senza mail: quanto scritto.
        const body = mail.manual ? draft.description : (state.mail?.itemId && state.mail.itemId === mail.itemId ? state.mail.body : mail.body);
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
        if (state.extraTicketFor === mail.conversationId) return extraTicketDone(mail, ticket, config, draft);
        Usage.record(draft.companyId, draft.activityId, mail.manual ? draft.subject : mail.subject);
        config.recent = [{ company: draft.companyId, activity: draft.activityId, timestamp: new Date().toISOString() },
            ...config.recent.filter(item => !(item.company === draft.companyId && item.activity === draft.activityId))].slice(0, 10);
        saveConfig(config);
        if (mail.manual) {
            // Senza mail: il modulo compilato diventa la scheda del ticket nuovo, poi il pulsante torna al suo posto.
            const form = root.querySelector('.ao-new-ticket-card');
            const from = form && !reducedMotion() ? shape(form) : null;
            holdRender();
            if (from) await fadeAll([...form.children], 1, 0, { duration: MOTION.hold, easing: 'ease-out', fill: 'forwards' });
            await refreshLists().catch(() => {});
            state.newTicketOpen = false;
            delete state.ticketDrafts[NEW_TICKET_KEY];
            releaseRender();
            render();
            const button = aside(root.querySelector('.ao-scroll .ao-new-ticket'));
            const item = [...root.querySelectorAll('.ao-scroll .ao-item')].find(element => element.querySelector('.ao-item-number')?.textContent === ticket.number);
            if (from && item) await holdingRender(() => morphInto(item, from));
            appear(button);
            showStatus(TEXTS.mail.ticketOpened(ticket.number), 'success', { label: TEXTS.queue.goToTicket, onclick: () => openTicketItem(ticket) });
            return;
        }
        render();
        await refreshLists();
        outcomeRelease(card);
        // Da In coda si resta in In coda (si smista la prossima); l'avviso porta al ticket per lavorarlo.
        if (card || mail.manual) showStatus(TEXTS.mail.ticketOpened(ticket.number), 'success', { label: TEXTS.queue.goToTicket, onclick: () => openTicketItem(ticket) });
    }

    // Ticket nella pagina del thread: la stessa scheda della lista (che vi atterra), con in più la description come
    // citazione e le date. Compatta di base; con un clic si apre come in In coda (si solleva, le altre si sfocano)
    // e sotto scende Attività, con il tempo lavorato. Un ticket solo nel thread (o senza mail) si apre già aperto.
    // Con più ticket se ne apre uno alla volta: un clic su un altro prima richiude quello aperto.
    // Il principale (il primo collegato al thread) si lavora con la risposta via mail e si chiude per ultimo;
    // gli altri (collegati) sono indipendenti: testo per ServiceNow e azione subito, come i ticket senza mail.
    function ticketBox(ticket, config, user, { isCurrentMail, single = false, primary = false, others = 0 }) {
        const open = OPEN_STATES.includes(ticket.state);
        const known = ticket.number in state.ticketExpanded;
        const expanded = open && (known ? state.ticketExpanded[ticket.number] : single);
        const pick = event => {
            if (state.holdRender || event.target.closest('.ao-ticket-work, button, a, input')) return;
            const other = root.querySelector('.ao-ticket-card.ao-expanded');
            if (other && other !== event.currentTarget) return collapseTicket(other);
            if (!expanded) expandTicket(ticket.number);
        };
        const grounded = Boolean(state.journey) || state.liftTicket === ticket.number;
        return h('section', {
            class: `ao-item ao-ticket-card${open ? '' : ' ao-static'}${expanded ? ' ao-expanded' : ''}${expanded && grounded ? ' ao-unlifted' : ''}${state.journey ? ' ao-away' : ''}`,
            tabindex: open ? '0' : null, 'aria-expanded': open ? String(expanded) : null, onclick: open ? pick : null,
            onkeydown: event => { if (open && event.key === 'Enter' && event.target === event.currentTarget) pick(event); }
        },
            h('div', { class: 'ao-ticket-face' }, ticketHead(ticket, user, { owner: true, takeover: true, primary: primary && others > 0 && Boolean(ticket.messageId) })), // l'icona serve a distinguerlo dai collegati
            h('div', { class: 'ao-ticket-more' },
                ticket.description ? h('div', { class: 'ao-ticket-description' }, ticket.description) : null, // il corpo della mail che l'ha generato
                ticket.state === '3' && ticket.holdExpiration
                    ? h('div', { class: `ao-hint${isHoldExpired(ticket) ? ' ao-hint-danger' : ''}` },
                        `${HOLD_REASONS.find(item => item.value === ticket.holdReason)?.label || 'On Hold'} · ${TEXTS.tickets.holdUntil(formatWhen(ticket.holdExpiration))}`) : null,
                open ? null : h('p', { class: 'ao-muted' }, TEXTS.mail.closed(TICKET_STATES[ticket.state]?.label || ticket.state)),
                h('div', { class: 'ao-ticket-dates' },
                    h('span', null, h('b', null, TEXTS.mail.opened), ' ', formatWhen(ticket.openedAt)),
                    h('span', null, h('b', null, TEXTS.mail.reported), ' ', formatWhen(ticket.reportedOn))),
                expanded ? ticketWork(ticket, config, { isCurrentMail, primary, others }) : null));
    }
    // Attività del ticket aperto: con la mail (il principale) "Salva e compila"; altrimenti testo e azione subito.
    // Il tempo lavorato sta sulla riga del titolo. La freccia in fondo richiude, come in In coda.
    function ticketWork(ticket, config, { isCurrentMail, primary, others }) {
        const timer = timerControl(ticket, true);
        const withMail = Boolean(ticket.messageId);
        const work = !withMail || !primary ? updateCard(ticket, config, timer)
            : ticket.mailUnconfirmed ? h('section', { class: 'ao-card' }, workTitle(timer)) // mail da confermare: niente azioni
                : onSendCard(ticket, config, isCurrentMail, { timer, keepOpen: others > 0 });
        return h('div', { class: 'ao-ticket-work' }, work,
            h('button', {
                type: 'button', class: 'ao-collapse', 'aria-label': TEXTS.common.close, title: TEXTS.common.close,
                onclick: event => collapseTicket(event.currentTarget.closest('.ao-ticket-card'))
            }, svgIcon(CHEVRON_UP_ICON)));
    }
    const workTitle = timer => h('div', { class: 'ao-card-head' }, h('h3', { class: 'ao-card-title' }, TEXTS.mail.reply), timer);
    async function expandTicket(number) {
        await whenIdle();
        root.querySelectorAll('.ao-scroll .ao-ticket-card .ao-item-number').forEach(element => { state.ticketExpanded[element.textContent] = element.textContent === number; });
        state.liftTicket = number; // si disegna ancora a terra, poi si solleva con la transizione
        render();
        state.liftTicket = null;
        const card = headOf(number);
        card?.getBoundingClientRect();
        card?.classList.remove('ao-unlifted');
        await holdingRender(() => unfold(card?.querySelector('.ao-ticket-work')));
    }
    async function collapseTicket(card) {
        const number = card?.querySelector('.ao-item-number')?.textContent;
        if (!number || state.holdRender) return;
        card.classList.add('ao-unlifted'); // torna a terra mentre Attività risale
        holdRender();
        await fold(card.querySelector('.ao-ticket-work'));
        state.ticketExpanded[number] = false;
        releaseRender();
        render();
    }
    // Ticket chiusi del thread: una sezione chiusa in fondo, con un clic si vedono (solo da consultare).
    function closedTickets(conversationId, tickets) {
        const shown = state.closedOpen === conversationId;
        return h('div', { class: `ao-closed${shown ? ' ao-open' : ''}` },
            h('button', { type: 'button', class: 'ao-closed-toggle', 'aria-expanded': String(shown), onclick: () => { state.closedOpen = shown ? null : conversationId; render(); } },
                h('span', null, TEXTS.mail.closedTickets(tickets.length)), h('span', { class: 'ao-select-caret', 'aria-hidden': 'true' }, '▾')),
            shown ? tickets.map(ticket => h('div', { class: 'ao-closed-item' },
                h('div', { class: 'ao-ticket-head' }, h('span', { class: 'ao-item-number' }, ticket.number), stateChip(ticket)),
                h('div', { class: 'ao-item-sub' }, ticket.shortDescription))) : null);
    }
    // Tempo lavorato: − valore +, a larghezza fissa (i pulsanti non si spostano mai). Un clic = 5 minuti; tenendo
    // premuto il passo cresce. Clic sul valore: si scrive il tempo (2h, 1h30, 120) o la differenza (+2h, -30).
    // A ServiceNow va sempre e solo la differenza, una volta sola quando si rilascia o si conferma.
    function timerControl(ticket, enabled) {
        const total = ticket.timeWorkedMinutes || 0;
        let shown = total;
        const value = h('span', { class: 'ao-timer-value', role: 'button', tabindex: enabled ? '0' : null, title: enabled ? TEXTS.mail.timeWorked : null });
        const paint = () => { value.textContent = formatMinutes(shown); };
        const send = async () => {
            const delta = shown - total;
            if (!delta) return;
            const updated = await run(() => Services.addTime(ticket.number, delta, currentUser().displayValue));
            if (!updated) { shown = total; paint(); return; }
            rememberTicket(updated);
            render();
            showStatus(TEXTS.mail.timeChanged(updated.timeWorkedMinutes - total, updated.timeWorkedMinutes));
        };
        const stepper = direction => {
            let timer = null, repeats = 0;
            const go = () => {
                repeats += 1;
                const step = repeats > TIME_STEP.after[1] ? TIME_STEP.fastest : repeats > TIME_STEP.after[0] ? TIME_STEP.faster : TIME_STEP.base;
                shown = Math.max(0, shown + direction * step);
                paint();
            };
            const stop = () => {
                if (!repeats) return;
                clearTimeout(timer); clearInterval(timer);
                repeats = 0;
                send();
            };
            return h('button', {
                type: 'button', class: 'ao-timer-step', disabled: !enabled, 'aria-label': direction > 0 ? '+' : '−',
                onpointerdown: event => {
                    event.preventDefault();
                    event.currentTarget.setPointerCapture?.(event.pointerId);
                    go();
                    timer = setTimeout(() => { timer = setInterval(go, TIME_STEP.repeat); }, TIME_STEP.hold);
                },
                onpointerup: stop, onpointercancel: stop, onlostpointercapture: stop,
                onkeydown: event => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); go(); stop(); } },
                onclick: event => event.stopPropagation()
            }, direction > 0 ? '+' : '−');
        };
        const edit = event => {
            event.stopPropagation();
            if (!enabled || value.querySelector('input')) return;
            const input = h('input', { class: 'ao-timer-input', value: formatMinutes(shown), 'aria-label': TEXTS.mail.timeWorked });
            let finished = false;
            const done = keep => {
                if (finished) return;
                if (keep) {
                    const next = parseMinutes(input.value, shown);
                    if (next === null) { showStatus(TEXTS.mail.timeInvalid, 'danger'); input.focus(); return; }
                    shown = next;
                }
                finished = true;
                paint();
                if (keep) send();
            };
            input.onkeydown = key => { if (key.key === 'Enter') done(true); if (key.key === 'Escape') done(false); };
            input.onblur = () => { if (value.contains(input)) done(true); };
            input.onclick = inner => inner.stopPropagation();
            value.replaceChildren(input);
            input.focus();
            input.select();
        };
        value.onclick = edit;
        value.onkeydown = event => { if (event.key === 'Enter' && event.target === value) edit(event); };
        paint();
        return h('div', { class: 'ao-timer', onclick: event => event.stopPropagation() }, stepper(-1), value, stepper(1));
    }
    // "2h", "1h30", "1,5h", "120", "90m" = tempo totale; con + o − davanti = differenza da quello attuale.
    function parseMinutes(text, current) {
        const clean = String(text).trim().toLowerCase().replace(/\s+/g, '').replace(/min$/, 'm').replace('−', '-');
        const sign = clean.startsWith('-') ? -1 : 1;
        const relative = /^[+-]/.test(clean);
        const body = clean.replace(/^[+-]/, '');
        let minutes = null, match;
        if ((match = body.match(/^(\d+)h(?:(\d+)m?)?$/))) minutes = Number(match[1]) * 60 + (Number(match[2]) || 0);
        else if ((match = body.match(/^(\d+)[.,](\d+)h$/))) minutes = Math.round(parseFloat(`${match[1]}.${match[2]}`) * 60);
        else if ((match = body.match(/^(\d+)m?$/))) minutes = Number(match[1]);
        if (minutes === null) return null;
        return Math.max(0, relative ? current + sign * minutes : minutes);
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
                template: config.combinationResponses[key]?.response ? 'combination' : '',
                includeReference: true
            };
        }
        return state.actionDrafts[ticket.number];
    }
    // Azioni sul ticket (Aggiorna, Risolvi, On Hold, Annulla) con le icone e il colore dell'azione scelta.
    const CLOSING_ACTIONS = ['resolve', 'cancel'];
    function ticketActionBar(ticket, draft, hidden = []) { // hidden: azioni non proposte (il principale con collegati aperti non si chiude)
        return iconBar(Object.entries(TEXTS.mail.actions).filter(([value]) => !hidden.includes(value)).map(([value, label]) => ({ value, label, icon: ACTION_ICONS[value] })), draft.type,
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
    function updateCard(ticket, config, timer) {
        const draft = actionDraft(ticket, config);
        draft.text ??= '';
        const refresh = () => { const button = root.querySelector('.ao-update-card .ao-primary'); if (button) button.disabled = !draft.text.trim(); };
        return h('section', { class: 'ao-card ao-update-card' },
            // Prima la risposta (il testo), poi cosa fare del ticket.
            workTitle(timer),
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
        } else if (state.target?.kind === 'mail') {
            // Collegato chiuso: si resta sul thread, dove passa tra i chiusi (e il principale si può chiudere se era l'ultimo).
            state.ticketExpanded = {}; // se ne resta uno solo, si mostra aperto
            await refreshTarget(true);
            showStatus(TEXTS.mail.closedTicket(updated.number));
        } else {
            // Ticket senza mail chiuso (risolto o annullato): si torna all'elenco, dove non c'è più.
            setView('tickets');
            showStatus(TEXTS.mail.closedTicket(updated.number));
        }
    }
    // keepOpen: ci sono altri ticket aperti sul thread, quindi il principale non si può ancora chiudere (Risolvi, Annulla):
    // prima si chiudono gli altri, poi con l'ultimo si risponde alla mail.
    function onSendCard(ticket, config, isCurrentMail, { timer, keepOpen = false }) {
        const draft = actionDraft(ticket, config);
        const confirmedType = ticket.pendingAction?.type;
        if (keepOpen && CLOSING_ACTIONS.includes(draft.type) && draft.type !== confirmedType) draft.type = 'comment';
        const rerender = () => render();
        const key = combinationKey(ticket.companyId, ticket.activityId);
        const templates = [
            { value: '', label: TEXTS.mail.replyEmpty },
            ...(config.combinationResponses[key]?.response ? [{ value: 'combination', label: TEXTS.mail.replyCombination }] : []),
            ...Object.entries(config.usefulResponses).sort(byName).map(([id, item]) => ({ value: `useful:${id}`, label: item.title }))
        ];
        // Prima con cosa rispondere, poi cosa fare del ticket all'invio. "Salva e compila" conferma l'azione e compila
        // la risposta in Outlook (vuota se non si sceglie un testo). Confermata, la scelta resta accesa e bloccata;
        // il pulsante diventa "Annulla" (toglie l'azione e sblocca).
        const confirmed = Boolean(ticket.pendingAction);
        return h('section', { class: 'ao-card' },
            workTitle(timer),
            h('div', { class: `ao-form-fields${confirmed ? ' ao-locked' : ''}`, inert: confirmed ? '' : null },
                selectField(TEXTS.mail.replyText, templates, draft.template, value => { draft.template = value; }),
                checkboxField(TEXTS.mail.includeReference(ticket.number), draft.includeReference, value => { draft.includeReference = value; }),
                ticketActionBar(ticket, draft, keepOpen ? CLOSING_ACTIONS : []),
                choiceZone(`ticket:${ticket.number}`, actionFields(draft))),
            h('div', { class: 'ao-actions' }, h('button', {
                type: 'button', class: `ao-button ao-primary${ACTION_FILLS[draft.type] || ''}`,
                onclick: confirmed ? async () => { if (await confirmMorph(ticket.number, () => savePendingAction(ticket, null))) showStatus(TEXTS.mail.pendingCleared); } : () => prepareReply(ticket, isCurrentMail)
            }, confirmed ? TEXTS.mail.changeChoice : TEXTS.mail.prepareReply)));
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
        );
    }
    function buildPendingAction(draft, by) {
        const action = { type: draft.type, by }; // anche Aggiorna si conferma; il testo è quello della mail
        if (draft.type === 'resolve') action.resolutionCode = draft.resolutionCode;
        if (draft.type === 'hold') {
            const date = parseDate(draft.holdExpiration.replace('T', ' '));
            action.holdReason = draft.holdReason;
            action.holdExpiration = date ? formatServiceNowDate(date) : '';
        }
        return action;
    }
    // Conferma e annullamento dell'azione: la scelta si accende (o si spegne) e i campi si bloccano (o si sbloccano)
    // con un passaggio morbido dallo stato di prima; il pulsante cambia in dissolvenza.
    async function confirmMorph(number, change) {
        const attivita = () => [...root.querySelectorAll('.ao-ticket-card')].find(card => card.querySelector('.ao-item-number')?.textContent === number)?.querySelector('.ao-form-fields')?.closest('.ao-card');
        const before = attivita();
        const thumb = before?.querySelector('.ao-segment-thumb');
        const old = thumb ? { boxShadow: getComputedStyle(thumb).boxShadow, filter: getComputedStyle(thumb).filter } : null;
        const fields = [...(before?.querySelectorAll('.ao-form-fields > :not(.ao-segmented)') || [])].map(field => getComputedStyle(field).opacity);
        const result = await change();
        const after = attivita();
        if (!after || reducedMotion()) return result;
        const timing = { duration: MOTION.swap * 1.6, easing: 'cubic-bezier(.3, .8, .2, 1)' };
        const newThumb = after.querySelector('.ao-segment-thumb');
        if (newThumb && old) newThumb.animate([old, { boxShadow: getComputedStyle(newThumb).boxShadow, filter: getComputedStyle(newThumb).filter }], timing);
        [...after.querySelectorAll('.ao-form-fields > :not(.ao-segmented)')].forEach((field, index) => {
            if (fields[index] !== undefined) field.animate([{ opacity: fields[index] }, { opacity: getComputedStyle(field).opacity }], timing);
        });
        after.querySelector('.ao-actions > .ao-button')?.animate([{ color: 'transparent' }, { color: getComputedStyle(after.querySelector('.ao-actions > .ao-button')).color }], timing); // stesso pulsante: cambia solo la scritta
        return result;
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
        // Il riferimento all'Incident nella risposta (in fondo) rende sicuro il collegamento mail ↔ ticket.
        const text = [template || '', draft.includeReference ? TEXTS.mail.reference(ticket.number) : ''].filter(Boolean).join('\n\n'); // il riferimento in fondo
        const updated = await confirmMorph(ticket.number, () => savePendingAction(ticket, buildPendingAction(draft, currentUser(config).displayValue)));
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

    // ---------------------------- Impostazioni ----------------------------
    // Profilo e configurazioni in schede separate (barra secondaria). Ogni configurazione mostra
    // subito tutti i suoi valori; la modifica avviene in linea. EDITORS descrive i campi.
    // I riferimenti a ServiceNow (Company, Assignment group, Category, Assigned to) si scrivono
    // per nome: il sys_id resta interno (arriva dall'import o dalla verifica del servizio) e non
    // si mostra né si modifica. Se il nome cambia, il sys_id si azzera e la voce è "da verificare".
    const EDITORS = {
        companies: {
            // scheda compatta: sotto il nome company e gruppo; accanto il primo dominio
            summary: item => [normalizeReference(item.company).displayValue, normalizeReference(item.assignmentGroup).displayValue].filter(Boolean).join(' · '),
            badge: item => ((item.domains || []).length ? `@${item.domains[0]}${item.domains.length > 1 ? ` +${item.domains.length - 1}` : ''}` : null),
            unverified: item => [item.company, item.assignmentGroup].some(reference => normalizeReference(reference).displayValue && !normalizeReference(reference).value),
            fields: [
                { key: 'name', label: TEXTS.settings.companyName, required: true },
                { key: 'company', label: TEXTS.settings.company, reference: true },
                { key: 'assignmentGroup', label: TEXTS.settings.assignmentGroup, reference: true },
                { key: 'domains', label: TEXTS.settings.domains, hint: TEXTS.settings.domainsHint, list: true }
            ]
        },
        activities: {
            summary: item => [normalizeReference(item.category).displayValue, `${SEVERITY_LABELS[item.impact] || '?'} / ${SEVERITY_LABELS[item.urgency] || '?'}`, TEXTS.mail.minutes(Number(item.timeWorked) || 0)].filter(Boolean).join(' · '),
            badge: item => item.shortDescriptionPrefix || null,
            unverified: item => Boolean(normalizeReference(item.category).displayValue && !normalizeReference(item.category).value),
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
            summary: item => item.response || '',
            fields: [
                { key: 'title', label: TEXTS.settings.responseTitle, required: true },
                { key: 'response', label: TEXTS.settings.response, multiline: true }
            ]
        }
    };
    const collectionOf = tab => EDITORS[tab].collection || tab;

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
    // Agenzie, Attività e Risposte: le stesse schede del resto del pannello. Compatte (nome, riassunto, etichetta);
    // un clic le apre come in In coda (si sollevano, le altre si sfocano) con il modulo già compilato, da modificare lì;
    // la freccia le richiude. "+ Nuova …" ha lo stile di "+ Nuovo ticket" e si trasforma nel modulo vuoto.
    function collectionSection(tab, config) {
        const editor = EDITORS[tab];
        const items = Object.entries(config[collectionOf(tab)]).sort(byName);
        const editing = state.editing?.kind === tab ? state.editing : null;
        const nameKey = editor.nameKey || 'name';
        return [
            editing?.id === null
                ? h('section', { class: 'ao-card ao-new-ticket-card ao-config-new' }, h('div', { class: 'ao-inline-form' }, editorForm(tab, null, {})))
                : h('button', { type: 'button', class: 'ao-new-ticket', onclick: () => openNewItem(tab) }, TEXTS.settings.newItem[tab]),
            items.length ? h('ul', { class: 'ao-list' }, items.map(([id, item]) => {
                const expanded = editing?.id === id;
                const badge = editor.badge?.(item);
                return h('li', {
                    class: `ao-item ao-config-item${expanded ? ' ao-expanded' : ''}${expanded && state.liftConfig === id ? ' ao-unlifted' : ''}`,
                    dataset: { id }, tabindex: '0', 'aria-expanded': String(expanded),
                    onclick: event => pickConfigItem(event, tab, id),
                    onkeydown: event => { if (event.key === 'Enter' && event.target === event.currentTarget) pickConfigItem(event, tab, id); }
                },
                    h('div', { class: 'ao-config-face' },
                        h('div', { class: 'ao-item-head' },
                            h('span', { class: 'ao-item-title' }, item[nameKey]),
                            editor.unverified?.(item) ? chip(TEXTS.settings.toVerify, 'warning', TEXTS.settings.toVerifyHint) : null,
                            badge ? chip(badge, 'neutral') : null),
                        editor.summary(item) ? h('div', { class: `ao-config-summary${tab === 'responses' ? ' ao-config-text' : ''}` }, editor.summary(item)) : null),
                    expanded ? h('div', { class: 'ao-config-more' }, editorForm(tab, id, item)) : null);
            })) : (editing ? null : h('p', { class: 'ao-muted' }, TEXTS.settings.empty))
        ];
    }
    // Clic su una scheda: se un'altra è aperta prima si richiude quella; altrimenti si apre questa.
    function pickConfigItem(event, tab, id) {
        if (state.holdRender || event.target.closest('.ao-config-more, button, a, input, textarea, select')) return;
        const open = root.querySelector('.ao-config-item.ao-expanded');
        if (open && open !== event.currentTarget) return collapseConfigItem(open);
        if (!open) openConfigItem(tab, id);
    }
    async function openConfigItem(tab, id) {
        await whenIdle();
        state.editing = { kind: tab, id };
        state.liftConfig = id; // si disegna ancora a terra, poi si solleva con la transizione
        render();
        state.liftConfig = null;
        const card = configCardOf(id);
        card?.getBoundingClientRect();
        card?.classList.remove('ao-unlifted');
        await holdingRender(() => unfold(card?.querySelector('.ao-config-more')));
    }
    async function collapseConfigItem(card) {
        if (!card || state.holdRender) return;
        card.classList.add('ao-unlifted'); // torna a terra mentre il modulo risale
        holdRender();
        await fold(card.querySelector('.ao-config-more'));
        state.editing = null;
        releaseRender();
        render();
    }
    const configCardOf = id => [...root.querySelectorAll('.ao-config-item')].find(card => card.dataset.id === id);
    async function openNewItem(tab) {
        await whenIdle();
        const button = root.querySelector('.ao-scroll .ao-new-ticket');
        state.newItemShape = button && shape(button);
        state.editing = { kind: tab, id: null };
        render();
        const card = root.querySelector('.ao-config-new');
        if (card && state.newItemShape && !reducedMotion()) await holdingRender(() => growFrom(card, state.newItemShape));
    }
    async function closeNewItem() {
        const card = root.querySelector('.ao-config-new');
        if (card && state.newItemShape && !reducedMotion()) await holdingRender(() => shrinkTo(card, state.newItemShape));
        state.editing = null;
        render();
    }
    // Campi del modulo (già compilati con i valori della voce) e i comandi: Salva ed Elimina sulla voce aperta,
    // Salva e Annulla su quella nuova. Salvata, la voce aperta si richiude; quella nuova diventa la sua scheda.
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
        const save = async event => {
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
            showStatus(TEXTS.settings.saved);
            if (id) return collapseConfigItem(event.currentTarget.closest('.ao-config-item'));
            // voce nuova: il modulo compilato diventa la sua scheda, poi "+ Nuova …" torna al suo posto
            const form = root.querySelector('.ao-config-new');
            const from = form && !reducedMotion() ? shape(form) : null;
            if (from) await holdingRender(() => fadeAll([...form.children], 1, 0, { duration: MOTION.hold, easing: 'ease-out', fill: 'forwards' }));
            state.editing = null;
            render();
            const button = aside(root.querySelector('.ao-scroll .ao-new-ticket'));
            const card = configCardOf(targetId);
            card?.scrollIntoView({ block: 'nearest' });
            if (from && card) await holdingRender(() => morphInto(card, from));
            appear(button);
        };
        return h('div', { class: 'ao-editor' },
            fields,
            h('div', { class: 'ao-actions' },
                h('button', { type: 'button', class: 'ao-button ao-primary', onclick: save }, TEXTS.common.save),
                h('span', { class: 'ao-spacer' }),
                id ? confirmButton(TEXTS.common.delete, 'ao-button ao-subtle ao-danger', () => deleteItem(tab, id))
                    : h('button', { type: 'button', class: 'ao-button ao-subtle', onclick: closeNewItem }, TEXTS.common.cancel)),
            id ? h('button', { type: 'button', class: 'ao-collapse', 'aria-label': TEXTS.common.close, title: TEXTS.common.close, onclick: event => collapseConfigItem(event.currentTarget.closest('.ao-config-item')) }, svgIcon(CHEVRON_UP_ICON)) : null);
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
        state.editing = null;
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
#ao-root .ao-tabs { position: relative; flex: none; display: flex; align-items: center; gap: 2px; padding: 0 4px; background: #0078d4; color: var(--ao-nav-ink); -webkit-backdrop-filter: var(--ao-nav-blur); backdrop-filter: var(--ao-nav-blur); } /* sfondo: continua l'intestazione (blendTabsWithHeader) */
#ao-root .ao-tabs .ao-icon-button { flex: none; font-size: var(--ao-fs-lg); color: var(--ao-nav-ink-soft); }
/* Freccia per tornare all'add-in: invisibile finché non la si raggiunge con il mouse, così le schede hanno tutta la larghezza */
#ao-root .ao-tabs .ao-back { position: absolute; left: 4px; top: 50%; z-index: 1; transform: translateY(-50%); opacity: 0; transition: opacity .2s; }
#ao-root .ao-tabs .ao-back:is(:hover, :focus-visible) { opacity: 1; }
#ao-root .ao-tabs .ao-icon-button:hover { background: rgba(255, 255, 255, .14); border-color: transparent; color: var(--ao-nav-ink); }
#ao-root .ao-tab { flex: 1; min-width: 0; display: flex; align-items: center; justify-content: center; gap: 0; padding: 9px 4px 7px; border: 0; border-bottom: 2px solid transparent; border-radius: var(--ao-radius-base) var(--ao-radius-base) 0 0; background: none; color: var(--ao-nav-ink-soft); font-size: var(--ao-fs-lg); font-weight: var(--ao-fw-normal); cursor: pointer; }
#ao-root .ao-tab:hover { color: var(--ao-nav-ink); }
#ao-root .ao-tab.ao-active { color: var(--ao-nav-ink); }
#ao-root .ao-tab-icon { position: relative; display: inline-flex; opacity: .8; transition: color .25s, opacity .25s; }
#ao-root .ao-tab.ao-active .ao-tab-icon { color: var(--ao-nav-ink); opacity: 1; }
/* Ogni scheda ha il suo posto: chiusa solo l'icona, aperta anche nome e contatore (si allargano nel suo posto) */
#ao-root .ao-tab-label { max-width: 0; margin-left: 0; overflow: hidden; white-space: nowrap; opacity: 0; transition: max-width var(--ao-selection-motion), margin-left var(--ao-selection-motion), opacity .25s; }
#ao-root .ao-tab.ao-active .ao-tab-label { max-width: 120px; margin-left: ${TAB_GAP.label}px; opacity: 1; } /* spazio anche per il segnalino */
@media (prefers-reduced-motion: reduce) { #ao-root .ao-tab-label { transition: none; } }
#ao-root .ao-tab-indicator { position: absolute; left: 0; bottom: 0; height: 3px; width: 0; border-radius: 3px 3px 0 0; background: var(--ao-nav-ink); pointer-events: none; opacity: 0; }
#ao-root .ao-tab-indicator.ao-ready { opacity: 1; }
#ao-root .ao-tab-indicator.ao-moving { transition: transform .32s cubic-bezier(.3, .8, .2, 1), width .32s cubic-bezier(.3, .8, .2, 1); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-tab-indicator.ao-moving { transition: none; } }
/* Scheda aperta: leggero zoom della scritta (e del numero), che resta finché è selezionata */
#ao-root .ao-tab > * { transition: transform .35s ease-in-out; }
#ao-root .ao-tab.ao-active > * { transform: scale(1.04); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-tab > * { transition: none; } }
/* Schede: il conteggio è un numero pulito, solo sulla scheda aperta. */
#ao-root .ao-counter { display: inline-block; min-width: 18px; height: 18px; padding: 0 4px; border-radius: var(--ao-radius-badge); color: var(--ao-text-muted); font-size: var(--ao-fs-md); font-weight: var(--ao-fw-medium); line-height: 18px; text-align: center; }
#ao-root .ao-tab.ao-active .ao-counter { color: var(--ao-nav-ink); }
/* Segnalino della scheda: sparisce rimpicciolendosi e svanendo, e la scritta si ricentra senza scatti (e il contrario quando torna) */
#ao-root .ao-tab .ao-counter { min-width: 0; margin-left: ${TAB_GAP.counter}px; padding: 0; max-width: 40px; color: var(--ao-nav-ink-soft); overflow: hidden; transition: opacity .2s ease, max-width .28s ease, margin-left .28s ease, transform .28s ease; } /* solo il numero: la sottolineatura si centra sulle cifre */
#ao-root .ao-tab .ao-counter.ao-counter-off { opacity: 0; max-width: 0; margin-left: 0; transform: scale(.6); }
@media (prefers-reduced-motion: reduce) { #ao-root .ao-tab .ao-counter { transition: none; } }
/* Aggiornamenti: segnalino tondo con il numero sull'icona, sempre visibile, del colore del testo intorno. */
#ao-root .ao-notify { position: absolute; top: -5px; right: -6px; min-width: 12px; height: 12px; padding: 0 3px; border-radius: 999px; background: currentColor; font-size: 8px; font-weight: var(--ao-fw-bold); line-height: 12px; text-align: center; pointer-events: auto; transition: opacity .2s, transform .2s; }
#ao-root .ao-notify-number { color: var(--ao-notify-ink); }
#ao-root .ao-tab .ao-notify-number { color: #005a9e; } /* blu dell'intestazione */
#ao-root .ao-notify-off { opacity: 0; transform: scale(.6); pointer-events: none; }
#ao-root .ao-segment:not(.ao-active) .ao-notify:not(.ao-notify-off) { opacity: .5; } /* selettore: sulla voce non scelta il segnalino è più discreto */
#ao-root .ao-notify-anchor { position: relative; display: inline-flex; margin-right: 5px; }
/* Card: numero dei messaggi nuovi su sfondo blu. */
#ao-root .ao-news { display: inline-block; min-width: 18px; height: 18px; padding: 0 4px; border-radius: 999px; background: var(--ao-accent); color: var(--ao-white); font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-bold); line-height: 18px; text-align: center; }
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
#ao-root .ao-item { display: grid; gap: 3px; padding: 9px 10px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); box-shadow: var(--ao-shadow-card), 0 0 0 transparent, inset 0 0 0 1px transparent; cursor: pointer; transition: background-color .35s ease, box-shadow .4s ease, filter .3s, transform .32s cubic-bezier(.3, .8, .2, 1); } /* stessa struttura d'ombra del passaggio del mouse: sfuma invece di scattare */
/* Ticket che si apre o si richiude: sollevato come la scheda aperta di In coda (openFromCard, closeToList) */
#ao-root .ao-item.ao-opening { position: relative; z-index: 1; transform: translateY(-2px) scale(1.015); box-shadow: var(--ao-raised-shadow); }
/* Il blu (titolo e bordo) è riservato alla scheda selezionata (modulo o dettagli aperti), che non reagisce all'hover
   (il passaggio del mouse, uguale per tutte le superfici, è nel blocco Superfici). */
#ao-root .ao-expanded, #ao-root .ao-expanded:hover { background: var(--ao-surface); } /* aperta: niente schiarimento al passaggio del mouse */
#ao-root .ao-item-top { display: flex; align-items: baseline; justify-content: space-between; gap: 8px; }
#ao-root .ao-item-title { font-weight: var(--ao-fw-medium); overflow-wrap: anywhere; }
#ao-root .ao-primary-mark { display: inline-flex; align-items: center; justify-content: center; width: 18px; height: 18px; margin-right: 5px; vertical-align: -3px; border-radius: 5px; background: var(--ao-accent-bg); color: var(--ao-text-accent); } /* principale del thread */
#ao-root .ao-primary-mark svg { width: 12px; height: 12px; }
#ao-root .ao-item-number { font-weight: var(--ao-fw-bold); color: var(--ao-text-accent); }
#ao-root .ao-item-sub { display: flex; align-items: center; justify-content: space-between; gap: 8px; color: var(--ao-text-muted); font-size: var(--ao-fs-md); overflow-wrap: anywhere; } /* a destra l'etichetta dell'operatore */
#ao-root .ao-item-time { color: var(--ao-text-subtle); font-size: var(--ao-fs-sm); white-space: nowrap; }
#ao-root .ao-item-meta { display: flex; flex-wrap: wrap; align-items: center; gap: 6px; margin-top: 3px; }
#ao-root .ao-empty, #ao-root .ao-footnote { margin: 8px 0; color: var(--ao-text-subtle); text-align: center; }
#ao-root .ao-footnote { font-size: var(--ao-fs-sm); }

/* Etichette colorate */
#ao-root .ao-chip { display: inline-flex; align-items: center; padding: 1px 8px; border-radius: var(--ao-radius-pill); box-shadow: inset 0 0 0 1px var(--ao-chip-line); font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-medium); white-space: nowrap; }
#ao-root .ao-chip.ao-tone-neutral { background: var(--ao-line-light); color: var(--ao-text-muted); }
#ao-root .ao-chip.ao-tone-person { background: var(--ao-surface); color: var(--ao-text); box-shadow: inset 0 0 0 1px var(--ao-text-subtle); }
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
#ao-root .ao-mail-head { display: grid; gap: 2px; min-width: 0; padding: 2px 2px 0; }
/* Testata del thread: sempre visibile. Scorrendo prende il fondo e l'ombra, e l'oggetto si stringe a una riga */
#ao-root .ao-fixed.ao-pinned { will-change: auto; }
#ao-root .ao-thread-head { display: grid; gap: 8px; margin: -12px -8px 0 -12px; padding: 10px 8px 8px 12px;
    background: color-mix(in srgb, var(--ao-surface) calc(100% * var(--ao-lift)), transparent);
    box-shadow: 0 3px 10px rgba(0,0,0, calc(.12 * var(--ao-lift))); transition: box-shadow .25s ease-out, background-color .25s ease-out; }
#ao-root .ao-thread-title { display: flex; align-items: flex-start; gap: 4px; }
#ao-root .ao-thread-back { flex: none; display: inline-flex; align-items: center; justify-content: center; width: 26px; height: 26px; margin-left: -6px; padding: 0; border: 0; border-radius: var(--ao-radius-base); background: none; color: var(--ao-text-muted); cursor: pointer; transition: background-color .25s, color .25s; }
#ao-root .ao-thread-back:hover { background: var(--ao-hover); color: var(--ao-text); }
#ao-root .ao-thread-head .ao-mail-subject { transition: font-size .25s; }
#ao-root .ao-thread-head .ao-mail-from, #ao-root .ao-thread-head .ao-mail-open { max-height: 40px; overflow: hidden; transition: max-height .25s, opacity .25s; }
#ao-root .ao-scrolled .ao-thread-head .ao-mail-subject { font-size: var(--ao-fs-base); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
#ao-root .ao-scrolled .ao-thread-head :is(.ao-mail-from, .ao-mail-open) { max-height: 0; opacity: 0; }
#ao-root .ao-thread-bar { display: flex; align-items: center; justify-content: space-between; gap: 8px; padding-top: 6px; border-top: 1px solid var(--ao-line); }
#ao-root .ao-thread-count { color: var(--ao-text-muted); font-size: var(--ao-fs-sm); }
#ao-root .ao-pill { flex: none; display: inline-flex; align-items: center; height: 26px; padding: 0 10px; border: var(--ao-call-border); border-radius: var(--ao-radius-pill); background: var(--ao-surface); box-shadow: var(--ao-call-glow); color: var(--ao-text-accent); font: inherit; font-size: var(--ao-fs-md); font-weight: var(--ao-fw-medium); cursor: pointer; transition: border-color .4s ease, box-shadow .4s ease; } /* come + Nuovo ticket */
#ao-root .ao-pill.ao-away { visibility: hidden; }
#ao-root .ao-morph { position: fixed; z-index: 20; pointer-events: none; border: 1px solid var(--ao-call-open-border); background: var(--ao-surface); box-shadow: var(--ao-call-glow), 0 6px 16px rgba(0,0,0,.12); } /* "＋ Ticket" che diventa il modulo */
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
#ao-root .ao-form-fields .ao-fav { padding: 4px 10px; border-color: transparent; background: var(--ao-fav-bg); box-shadow: 0 1px 2px rgba(0,0,0,.06); color: var(--ao-text); transition: background-color .2s, color .2s; }
#ao-root .ao-form-fields .ao-fav:hover { border-color: transparent; background: var(--ao-hover); }
#ao-root .ao-form-fields .ao-fav.ao-active { border-color: transparent; box-shadow: none; background: color-mix(in srgb, var(--ao-accent) 14%, var(--ao-surface)); color: var(--ao-text-accent); font-weight: var(--ao-fw-normal); }
#ao-root .ao-form-fields .ao-fav-company { color: var(--ao-text-subtle); }
#ao-root .ao-form-fields .ao-fav.ao-active .ao-fav-company { color: inherit; }
#ao-root .ao-new-ticket { display: flex; align-items: center; justify-content: center; gap: 6px; margin-bottom: 6px; padding: 8px; border: var(--ao-call-border); border-radius: var(--ao-radius-base); background: var(--ao-surface); box-shadow: var(--ao-call-glow); color: var(--ao-text-accent); font: inherit; font-weight: var(--ao-fw-medium); cursor: pointer; transition: border-color .4s ease, box-shadow .4s ease; }
/* Aperto: come una scheda di In coda aperta (sollevato, gli altri ticket sfocati), col suo bagliore */
#ao-root .ao-new-ticket-card { position: relative; z-index: 1; border-color: var(--ao-call-open-border); transform: translateY(-2px) scale(1.015); box-shadow: var(--ao-call-glow), 0 6px 16px rgba(0,0,0,.12), 0 1px 3px rgba(0,0,0,.08); } /* nasce dal pulsante (growFrom, morphInto) */
#ao-root .ao-view:has(.ao-new-ticket-card) :is(.ao-list .ao-item, .ao-ticket-card) { filter: blur(5px); }
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

/* Storico */

/* Torna all'elenco (vista di lavoro del ticket) */
#ao-root .ao-back-row { display: flex; align-items: center; min-height: 22px; }
#ao-root .ao-back-link { font-weight: var(--ao-fw-medium); }

/* Impostazioni */
#ao-root .ao-editor { display: grid; gap: 8px; } /* modulo dentro la scheda (Impostazioni) */

/* In coda: riga con azioni e modulo compatto */
#ao-root .ao-open-first { display: grid; justify-items: start; gap: 8px; padding: 12px; border: 1px dashed var(--ao-line-strong); border-radius: var(--ao-radius-base); color: var(--ao-text-muted); }
#ao-root .ao-open-first p { margin: 0; }
#ao-root .ao-item-main { display: grid; grid-template-columns: minmax(0, 1fr) auto; gap: 3px 10px; cursor: pointer; border-radius: calc(var(--ao-radius-base) - 2px); }
#ao-root .ao-item-left { display: grid; gap: 3px; align-content: start; min-width: 0; }
#ao-root .ao-inline-form { display: grid; gap: 8px; padding-top: 8px; border-top: 1px solid var(--ao-line-light); }
/* Scheda di In coda aperta: testo della mail */
#ao-root .ao-mail-preview { max-height: 230px; overflow: hidden; padding: 6px 8px; border-radius: calc(var(--ao-radius-base) - 2px); background: var(--ao-quote-bg); color: var(--ao-text-muted); font-size: var(--ao-fs-md); line-height: 1.45; white-space: pre-line; overflow-wrap: anywhere; scrollbar-width: thin; } /* come la description nel thread */
#ao-root .ao-mail-preview.ao-scrollable { overflow-y: auto; } /* testo oltre 230px (scheda di sola lettura grande circa quanto il modulo di Senza mail): solo allora la barra */
/* Più messaggi nel thread: frecce per scorrerli; quello che si guarda è quello da cui nasce il ticket */
#ao-root .ao-message-nav { display: flex; align-items: center; gap: 4px; min-width: 0; }
#ao-root .ao-message-step { display: inline-flex; align-items: center; justify-content: center; width: 26px; height: 26px; padding: 0; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); color: var(--ao-text); cursor: pointer; }
#ao-root .ao-message-step:hover:where(:not(:disabled)) { background: var(--ao-hover); }
#ao-root .ao-message-step:disabled { opacity: .4; cursor: default; }
#ao-root .ao-message-where { color: var(--ao-text-muted); font-size: var(--ao-fs-sm); font-variant-numeric: tabular-nums; white-space: nowrap; }
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
#ao-root .ao-choice-accent { --ao-choice-bg: color-mix(in srgb, var(--ao-accent) 16%, var(--ao-surface)); --ao-choice-glow: var(--ao-accent); --ao-choice-text: var(--ao-text-accent); } /* più deciso dell'azzurro dei filtri: si stacca dal fondo grigio */
#ao-root .ao-choice-forward { --ao-choice-bg: var(--ao-choice-forward-bg); --ao-choice-glow: var(--ao-tone-forward); --ao-choice-text: var(--ao-choice-forward-text); }
#ao-root .ao-choice-danger { --ao-choice-bg: var(--ao-choice-danger-bg); --ao-choice-glow: var(--ao-tone-ignore); --ao-choice-text: var(--ao-choice-danger-text); }
#ao-root .ao-choice-resolve { --ao-choice-bg: var(--ao-choice-resolve-bg); --ao-choice-glow: var(--ao-tone-resolve); --ao-choice-text: var(--ao-choice-resolve-text); }
#ao-root .ao-choice-hold { --ao-choice-bg: var(--ao-choice-hold-bg); --ao-choice-glow: var(--ao-tone-hold); --ao-choice-text: var(--ao-choice-hold-text); }
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
    box-shadow: var(--ao-open-shadow);
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
/* Category: spazio per due righe, etichetta centrata in altezza */

/* Risposte nuove e più ticket sullo stesso thread */
/* Contenitore del ticket nella pagina del thread */
/* Testata del ticket (lista e thread): numero con stato ed etichetta affiancati; agenzia e attività; short description */
/* Ticket nel thread: una sola scheda. Sotto la testata (uguale alla lista) la description come citazione, le date,
   e da estesa Attività; estesa si solleva con l'ombra di In coda */
#ao-root .ao-ticket-card { display: grid; gap: 0; cursor: pointer; }
#ao-root .ao-ticket-card.ao-static { cursor: default; }
/* Aperto: come In coda, in primo piano (si solleva e si ingrandisce appena); gli altri ticket si sfocano */
#ao-root :is(.ao-ticket-card, .ao-config-item).ao-expanded { position: relative; z-index: 1; cursor: default; transform: translateY(-2px) scale(1.015); box-shadow: var(--ao-open-shadow); }
/* Appena aperta, ancora a terra: lo stesso aspetto del passaggio del mouse (bagliore e bordo), così sollevandosi il
   bagliore prosegue senza spegnersi; richiudendo torna allo stesso aspetto */
#ao-root :is(.ao-queue-item, .ao-ticket-card, .ao-config-item).ao-expanded.ao-unlifted { transform: none; box-shadow: var(--ao-item-hover-glow), var(--ao-shadow-card), inset 0 0 0 1px var(--ao-item-hover-edge); }
#ao-root .ao-view:has(.ao-ticket-card.ao-expanded:not(.ao-unlifted)) .ao-ticket-card:not(.ao-expanded) { filter: blur(5px); }
#ao-root .ao-view:has(.ao-config-item.ao-expanded:not(.ao-unlifted)) .ao-config-item:not(.ao-expanded) { filter: blur(5px); }
/* Voci di Impostazioni (agenzie, attività, risposte): schede compatte come le altre; aperte, il modulo sotto */
#ao-root .ao-config-face { display: grid; gap: 2px; }
#ao-root .ao-item-head { display: flex; align-items: center; gap: 6px; }
#ao-root .ao-item-head .ao-item-title { flex: 1; min-width: 0; }
#ao-root .ao-config-summary { color: var(--ao-text-muted); font-size: var(--ao-fs-md); overflow-wrap: anywhere; }
#ao-root .ao-config-text { white-space: pre-line; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; } /* risposta: le prime due righe */
#ao-root .ao-config-more { margin-top: 8px; padding-top: 8px; border-top: 1px solid var(--ao-line); overflow: hidden; }
#ao-root .ao-ticket-card.ao-away { visibility: hidden; } /* la sua scheda sta ancora arrivando */
#ao-root .ao-ticket-more { display: grid; gap: 6px; padding-top: 6px; }
#ao-root .ao-ticket-description { padding: 6px 8px; border-radius: calc(var(--ao-radius-base) - 2px); background: var(--ao-quote-bg); color: var(--ao-text-muted); font-size: var(--ao-fs-md); line-height: 1.45; white-space: pre-line; overflow-wrap: anywhere; }
#ao-root .ao-ticket-dates { display: flex; gap: 12px; color: var(--ao-text-subtle); font-size: var(--ao-fs-sm); }
#ao-root .ao-ticket-dates b { color: var(--ao-text-muted); font-weight: var(--ao-fw-medium); }
#ao-root .ao-ticket-work { display: grid; gap: 6px; }
#ao-root .ao-ticket-work > .ao-card { padding: 8px 0 0; border: 0; border-top: 1px solid var(--ao-line); border-radius: 0; background: none; box-shadow: none; -webkit-backdrop-filter: none; backdrop-filter: none; cursor: default; } /* Attività dentro la scheda */
#ao-root .ao-item-tags { display: flex; align-items: center; gap: 4px; flex: none; }
#ao-root .ao-item-where { display: flex; align-items: center; justify-content: space-between; gap: 10px; }
#ao-root .ao-item-place { min-width: 0; }
#ao-root .ao-item-company { color: var(--ao-text); font-weight: var(--ao-fw-medium); }
#ao-root .ao-item-activity { color: var(--ao-text); font-size: var(--ao-fs-md); }
/* Tempo lavorato: larghezza fissa, così − e + restano sempre allo stesso posto (serve per tenere premuto).
   Il tempo sta incassato nella traccia dei selettori; − e + sono in rilievo come la voce scelta. */
#ao-root .ao-timer { flex: none; display: inline-flex; align-items: center; gap: 2px; height: 28px; padding: 2px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-segment-track); }
#ao-root .ao-timer-step { width: 24px; height: 100%; padding: 0; border: 0; border-radius: calc(var(--ao-radius-base) - 2px); background: var(--ao-selection); box-shadow: inset 0 0 0 1px var(--ao-selection-line), var(--ao-selection-shadow); color: var(--ao-text); font: inherit; font-size: var(--ao-fs-lg); cursor: pointer; touch-action: none; user-select: none; transition: filter .3s, box-shadow .35s; }
#ao-root .ao-timer-step:hover:where(:not(:disabled)) { filter: var(--ao-hover-filter); box-shadow: inset 0 0 0 1px var(--ao-selection-line), var(--ao-hover-glow); }
#ao-root .ao-timer-step:disabled { opacity: .4; cursor: default; }
#ao-root .ao-timer-value { width: 76px; text-align: center; color: var(--ao-text); font-size: var(--ao-fs-md); font-weight: var(--ao-fw-medium); font-variant-numeric: tabular-nums; white-space: nowrap; cursor: text; }
#ao-root .ao-timer-input { width: 100%; height: 20px; padding: 0; border: 0; border-radius: var(--ao-radius-badge); outline: none; background: var(--ao-surface); color: inherit; font: inherit; text-align: center; }
/* Azione confermata: la scelta resta accesa (come una scheda sotto il mouse) e non si può cambiare */
#ao-root .ao-locked .ao-segment-thumb { filter: var(--ao-hover-filter); box-shadow: inset 0 0 0 1.5px var(--ao-lock-line),
    0 0 18px color-mix(in srgb, var(--ao-choice-glow) var(--ao-lock-glow), transparent), 0 0 4px color-mix(in srgb, var(--ao-choice-glow) calc(var(--ao-lock-glow) / 2), transparent),
    var(--ao-selection-shadow); } /* bagliore forte come + Nuovo ticket, del colore dell'azione scelta */
#ao-root .ao-locked { cursor: default; }
#ao-root .ao-locked > :not(.ao-segmented) { opacity: .6; } /* bloccati: i campi si attenuano, la scelta resta accesa */
#ao-root .ao-locked .ao-segment:not(.ao-active) { opacity: var(--ao-disabled-opacity); } /* le altre azioni: spente come quelle non disponibili */
/* Ticket chiusi del thread */
#ao-root .ao-closed { display: grid; gap: 6px; }
#ao-root .ao-closed-toggle { display: flex; align-items: center; justify-content: space-between; padding: 7px 10px; border: 1px dashed var(--ao-line-strong); border-radius: var(--ao-radius-base); background: none; color: var(--ao-text-muted); font: inherit; font-size: var(--ao-fs-md); cursor: pointer; }
#ao-root .ao-closed.ao-open .ao-select-caret { transform: rotate(180deg); }
#ao-root .ao-closed-item { display: grid; gap: 2px; padding: 7px 10px; border: 1px solid var(--ao-line); border-radius: var(--ao-radius-base); background: var(--ao-surface); opacity: .8; }

/* Archivio ignorate, gruppi di ticket, anteprima dei messaggi nuovi */
/* Mazzo: thread con più ticket. In alto le fasce dei ticket dietro (con il numero), sotto la scheda davanti */
#ao-root .ao-deck { display: grid; cursor: pointer; }
#ao-root .ao-deck-strip { margin: 0 var(--ao-deck-inset) -8px; padding: 3px 10px 11px; border: 1px solid var(--ao-line); border-bottom: 0; border-radius: var(--ao-radius-base) var(--ao-radius-base) 0 0; background: linear-gradient(rgba(0, 0, 0, var(--ao-deck-shade)), rgba(0, 0, 0, var(--ao-deck-shade))), var(--ao-surface); /* la scheda davanti, un tono più scura per ogni passo indietro */ color: var(--ao-text-accent); font-size: var(--ao-fs-sm); font-weight: var(--ao-fw-medium); }
#ao-root .ao-deck > .ao-item { position: relative; }
#ao-root .ao-journey { position: fixed; z-index: 30; margin: 0; pointer-events: none; list-style: none; } /* copia della scheda in viaggio */
/* mazzo in viaggio: l'ombra va sulla scheda davanti, non sull'intero mazzo (sopra c'è lo spazio delle fasce) */
#ao-root .ao-deck.ao-journey { box-shadow: none !important; }
#ao-root .ao-deck.ao-journey > .ao-item { box-shadow: var(--ao-raised-shadow); }
#ao-root .ao-ghost { position: fixed; z-index: 29; margin: 0; overflow: hidden; pointer-events: none; will-change: filter, opacity; } /* sfocatura preparata in anticipo */ /* copia ferma della lista durante l'andata */
#ao-root .ao-item-late { color: var(--ao-danger); font-weight: var(--ao-fw-medium); }
#ao-root .ao-new-preview { display: grid; gap: 2px; padding: 6px 8px; border-radius: var(--ao-radius-base); background: var(--ao-surface); border: 1px solid var(--ao-accent-border); }
#ao-root .ao-item .ao-new-preview { background: var(--ao-accent-bg); }
#ao-root .ao-new-head { display: flex; align-items: baseline; gap: 6px; font-size: var(--ao-fs-sm); }
#ao-root .ao-new-from { flex: 1; font-weight: var(--ao-fw-medium); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
#ao-root .ao-new-text { color: var(--ao-text-muted); font-size: var(--ao-fs-md); display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; }

/* Collegamento mail ↔ ticket */
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
#ao-root .ao-item, #ao-root .ao-deck-strip, #ao-root .ao-card, #ao-root .ao-segmented, #ao-root .ao-bar-btn, #ao-root .ao-new-ticket, #ao-root .ao-candidate, #ao-root .ao-statusbar { -webkit-backdrop-filter: var(--ao-blur); backdrop-filter: var(--ao-blur); }
#ao-root .ao-item:not(.ao-expanded, .ao-static):hover { background: var(--ao-item-hover); box-shadow: var(--ao-shadow-card), var(--ao-item-hover-glow), inset 0 0 0 1px var(--ao-item-hover-edge); } /* solo la scheda: niente filtro sul contenuto (timer, etichette) */
#ao-root :is(.ao-new-ticket, .ao-thread-bar .ao-pill):hover { border-color: var(--ao-call-edge-hover); box-shadow: var(--ao-call-glow-hover); animation: ao-call-pulse .45s ease-out; } /* non si schiarisce: si accende */
@keyframes ao-call-pulse { /* impulso breve e misurato (nel flat il bagliore non c'è) */
    0% { box-shadow: var(--ao-call-glow); }
    30% { box-shadow: var(--ao-call-glow-pulse); }
    100% { box-shadow: var(--ao-call-glow-hover); } }
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
    const glassHeaderRule = (look, ink = '#fff') => `background: ${look.background}; box-shadow: ${look.light}; color: ${ink};`;
    const OUTSIDE_STYLES = `
:has(> .app-header):has(#amp-view:not([hidden])) > .app-header { padding-bottom: 6px; }
:has(> .app-header):has(#amp-view:not([hidden]) #ao-root[data-style="glass"]) { background: ${THEME.glassBackdrop}; }
:has(> .app-header):has(#amp-view:not([hidden]) #ao-root[data-style="glass"][data-theme="dark"]) { background: ${THEME.glassDarkBackdrop}; }
@media (prefers-color-scheme: dark) { :has(> .app-header):has(#amp-view:not([hidden]) #ao-root[data-style="glass"]:not([data-theme="light"])) { background: ${THEME.glassDarkBackdrop}; } }
:has(> .app-header):has(#amp-view:not([hidden]) #ao-root[data-style="glass"]) > .app-header { -webkit-backdrop-filter: ${THEME.glassColors.navBlur}; backdrop-filter: ${THEME.glassColors.navBlur}; ${glassHeaderRule(THEME.glassHeader.light)} }
:has(> .app-header):has(#amp-view:not([hidden]) #ao-root[data-style="glass"][data-theme="dark"]) > .app-header { ${glassHeaderRule(THEME.glassHeader.dark, THEME.glassDarkColors.navInk)} }
@media (prefers-color-scheme: dark) { :has(> .app-header):has(#amp-view:not([hidden]) #ao-root[data-style="glass"]:not([data-theme="light"])) > .app-header { ${glassHeaderRule(THEME.glassHeader.dark, THEME.glassDarkColors.navInk)} } }
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
