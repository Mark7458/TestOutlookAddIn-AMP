/*
 * AMP - SNOW Assistant: pannello per l'area AMP dell'add-in Outlook.
 *
 * Replica del pannello dello script Tampermonkey (stesse sezioni, funzioni e testi),
 * con lo stile dell'add-in. Generato da verifica/make-addin.py a partire dallo script.
 *
 * Collegamento con l'add-in: espone window.AmpModule { initialize, show };
 * il pulsante "Apri area AMP" (#open-amp-button) chiama show(), la freccia
 * in alto a sinistra torna alla schermata principale (TestPluginAPI.showMainView).
 *
 * Le operazioni sul ticket passano dal blocco 3 (Adattatore ServiceNow). Nell'anteprima
 * (amp/preview.html) lavorano su un modulo Incident simulato; nell'add-in vanno
 * collegate ai connettori della Function App.
 */
(function () {
    'use strict';

    // Indice dei blocchi
    //  0. Aspetto: tema, testi, icone          8. Chiusura e On Hold
    //  1. Configurazione e stato          9. Import / export
    //  2. Salvataggio dati               10. Interfaccia: pannello
    //  3. Adattatore ServiceNow          11. Interfaccia: configurazioni salvate
    //  4. Configurazioni                 12. Interfaccia: elenchi a discesa con ricerca
    //  5. Preferiti                      13. Tour guidato
    //  6. Risposte                       14. Stili (usano i valori del blocco 0)
    //  7. Ticket aperti e On Hold        15. Avvio e collegamento con l'add-in
    //
    // Regola: solo il blocco 3 legge o scrive la pagina ServiceNow
    // (g_form, snDocument, snWindow). Gli altri blocchi passano da lì.

    // ======================================================================
    // 0. ASPETTO: TEMA, TESTI E ICONE
    // Il punto da cui cambiare colori, caratteri, dimensioni, angoli, ombre, scritte e icone.
    // ======================================================================

    // ----------------------------------------------------------------------
    // TEMA: tutti i valori di aspetto del pannello in un unico punto.
    // Il CSS li usa come variabili --snow-* (vedi themeCssVariables()):
    // cambiando un valore qui cambia ovunque.
    // ----------------------------------------------------------------------
    const THEME = {
        colors: {
            // Colore principale (blu dell'add-in)
            accent: '#0078d4', // pulsanti principali, evidenze, voci selezionate
            accentDark: '#106ebe', // colore principale al passaggio del mouse
            accentBorder: '#c7e0f4', // bordo azzurro chiaro (passaggio del mouse, tour)
            accentBg: '#eff6fc', // sfondo azzurro chiaro (passaggio del mouse, voci selezionate)
            headerBackground: 'linear-gradient(135deg, #005a9e, #0078d4)', // intestazione, come quella dell'add-in
            // Testo
            white: '#fff',
            text: '#242424', // testo principale
            textMuted: '#605e5c', // testo secondario, etichette, pulsanti
            textSubtle: '#8a8886', // icone e testi attenuati, elenchi vuoti
            textDisabled: '#a19f9d', // segnaposto
            // Grigi per bordi e sfondi
            bgSoft: '#faf9f8', // sfondo tenue
            lineLight: '#edebe9', // separatori e bordi leggeri
            line: '#e1dfdd', // bordi di campi e riquadri
            lineStrong: '#c8c6c4', // bordi dei pulsanti
            // Esiti
            success: '#107c10', // verde "predisposto → Save"
            successDark: '#0b5a0b', // verde al passaggio del mouse
            danger: '#a4262c', // rosso: eliminazione, On Hold scaduto
            dangerBg: '#fde7e9', // sfondo rosso chiaro
            // Messaggi in fondo al pannello e avviso "Verifica il testo"
            statusSuccessBg: '#dff6dd',
            statusSuccessBorder: '#9fd89f',
            statusSuccessText: '#107c10',
            statusErrorBg: '#fde7e9',
            statusErrorBorder: '#f1bbbc',
            statusErrorText: '#a4262c',
            statusWarningBg: '#fff4ce',
            statusWarningBorder: '#f2d58c',
            statusWarningText: '#8a5a00',
            // Tooltip ticket (riquadro scuro)
            tooltipBg: '#323130',
            tooltipText: '#ffffff'
        },
        font: {
            family: '"Segoe UI", system-ui, -apple-system, BlinkMacSystemFont, sans-serif' // carattere dell'add-in
        },
        fontSize: {
            xs: '10px', // etichette minime
            sm: '11px', // testi secondari
            md: '12px', // testo normale
            base: '13px', // testo del pannello
            lg: '14px', // titoli
            xl: '15px' // titolo finestra
        },
        fontWeight: {
            normal: '400',
            medium: '600',
            bold: '700',
            heavy: '800'
        },
        radius: {
            base: '6px', // angoli di pulsanti, campi, riquadri, finestre, voci (come l'add-in)
            pill: '999px' // estremità tonde (cursore durata On Hold)
        },
        shadow: {
            panel: '-4px 0 12px rgba(0,0,0,.10)', // pannello laterale
            button: '0 2px 4px rgba(0,0,0,.10)', // pulsanti in rilievo
            buttonHover: '0 3px 7px rgba(0,0,0,.14)', // pulsanti in rilievo al passaggio
            focus: '0 0 0 1px rgba(0,120,212,.25)', // campo attivo
            dropdown: '0 5px 14px rgba(0,0,0,.16)' // elenchi a discesa
        },
        state: {
            disabledOpacity: '.45' // trasparenza di tutto ciò che non è cliccabile
        },
        layout: {
            headerHeight: '52px', // altezza intestazione
            controlHeight: '34px', // altezza campi e pulsanti standard
            actionHeight: '38px', // altezza Compila Ticket, stella, barra copia
            resolutionHeight: '40px' // altezza Predisponi Chiusura / On Hold
        }
    };

    // ----------------------------------------------------------------------
    // TESTI: tutte le scritte del pannello in un unico punto
    // (etichette, messaggi, conferme, tooltip, tour). Le voci con (n) o (nome)
    // sono funzioni: il valore tra parentesi viene inserito nel testo.
    // ----------------------------------------------------------------------
    const TEXTS = {
        panel: {
            title: 'SNOW Assistant',
            back: 'Torna all’add-in',
            profile: name => `● ${name}`,
            profileMissing: '○ Profilo non configurato',
            profileMissingInitial: 'Profilo non configurato',
            favorites: 'Preferiti',
            company: 'Agenzia',
            activity: 'Attività',
            companyPlaceholder: 'Seleziona agenzia',
            activityPlaceholder: 'Seleziona attività',
            apply: 'Compila Ticket',
            usefulResponses: 'Risposte salvate',
            preview: 'Anteprima risposta',
            reviewNotice: 'Verifica il testo prima di confermare.',
            copyResponse: 'Risposta',
            copyIncident: 'Incident',
            openTickets: 'Ticket aperti',
            openTicketsCount: count => `Ticket aperti (${count})`,
            toggleList: 'Apri o chiudi elenco',
            selectPlaceholder: 'Seleziona',
            noMatches: 'Nessuna corrispondenza',
            emptyFavorites: 'Nessun preferito',
            emptyResponses: 'Nessuna risposta salvata',
            emptyTickets: 'Nessun ticket aperto memorizzato'
        },
        settings: {
            favoritesTitle: 'Gestisci Preferiti',
            groupBy: 'Raggruppa per',
            groupNone: 'Nessuno',
            groupCompany: 'Agenzia',
            groupActivity: 'Attività',
            savedTitle: 'Configurazioni salvate',
            savedOpen: 'Visualizza configurazioni salvate',
            configTitle: 'Gestisci Configurazione',
            importJson: 'Importa JSON',
            exportJson: 'Esporta JSON',
            tour: 'Tour dell’Assistant'
        },
        closure: {
            prepare: 'Predisponi Chiusura',
            confirm: 'Conferma chiusura',
            prepared: 'Chiusura predisposta',
            cancelPrepared: 'Cancellazione predisposta',
            ticketClosed: 'Ticket Chiuso',
            statusLabel: 'Stato',
            statusPlaceholder: 'Seleziona stato',
            codeLabel: 'Resolution Code',
            codePlaceholder: 'Seleziona codice',
            noCode: 'Nessun codice', // ultima voce dei Resolution Code
            saveSuffix: 'Save' // dopo la freccia: "Chiusura predisposta → Save"
        },
        hold: {
            configure: 'Configura stato On Hold',
            prepare: 'Predisponi On Hold',
            prepared: 'On Hold predisposto',
            reasonLabel: 'Motivo On Hold',
            reasonPlaceholder: 'Seleziona motivo',
            durationLabel: 'Termine On Hold',
            durationAria: 'Durata On Hold',
            days: count => (count === 1 ? '1 giorno' : `${count} giorni`),
            cancel: 'Annulla'
        },
        tickets: {
            companyMissing: 'Agenzia non disponibile',
            activityMissing: 'Attività non disponibile',
            holdExpired: 'On Hold scaduto: richiede intervento.',
            holdExpiredAria: 'On Hold scaduto: richiede intervento',
            holdNoExpiration: 'Ticket in On Hold (scadenza non disponibile)',
            holdRemaining: time => `On Hold: ${time} rimanenti`,
            holdAria: 'Ticket in On Hold',
            days: count => `${count} ${count === 1 ? 'giorno' : 'giorni'}`,
            hours: count => `${count} ${count === 1 ? 'ora' : 'ore'}`,
            minutes: count => `${count} ${count === 1 ? 'minuto' : 'minuti'}`,
            openNewTab: 'Apri in nuova scheda',
            copyLink: 'Copia numero e link',
            remove: 'Elimina riferimento',
            invalidLink: 'Link non valido'
        },
        responses: {
            rename: 'Modifica nome',
            remove: 'Elimina risposta',
            // Risposta proposta quando non ce n'è una salvata per la combinazione (una riga per elemento).
            defaultText: ['Salve,', '', 'attività completata.', '', 'Saluti,']
        },
        summary: {
            title: 'Configurazioni salvate',
            close: 'Chiudi',
            tabsLabel: 'Tipo di configurazione',
            companyTab: 'Agenzia',
            activityTab: 'Attività',
            companyCount: count => (count === 1 ? '1 agenzia salvata' : `${count} agenzie salvate`),
            activityCount: count => (count === 1 ? '1 attività salvata' : `${count} attività salvate`),
            emptyCompany: 'Nessuna Agenzia salvata.',
            emptyActivity: 'Nessuna Attività salvata.',
            unnamedCompany: 'Agenzia senza nome',
            unnamedActivity: 'Attività senza nome',
            fieldCompany: 'Company',
            fieldAssignmentGroup: 'Assignment Group',
            fieldCategory: 'Category',
            fieldImpact: 'Impact',
            fieldUrgency: 'Urgency',
            fieldTimeWorked: 'Time Worked',
            timeWorkedValue: minutes => `${minutes} minuti`,
            fieldPrefix: 'Tag Short Description',
            none: 'Nessuno'
        },
        // Messaggi nella barra in fondo al pannello
        status: {
            fillCategory: 'Compila prima Category.',
            fillImpact: 'Compila prima Impact.',
            fillUrgency: 'Compila prima Urgency.',
            fillTimeWorked: 'Compila prima Time Worked.',
            fillAssignedTo: 'Compila prima Assigned to nel ticket.',
            fillCompany: 'Compila prima Company.',
            fillAssignmentGroup: 'Compila prima Assignment Group.',
            profileSet: name => `Profilo impostato: ${name}`,
            profileMissing: 'Configura prima il profilo personale.',
            companyExists: 'Agenzia già presente. Selezionala e usa la matita.',
            companySaved: name => `Agenzia salvata: ${name}`,
            companySelectToOverwrite: 'Seleziona l’agenzia da sovrascrivere.',
            companyNotFound: 'Agenzia selezionata non trovata.',
            companyOverwritten: name => `Agenzia sovrascritta: ${name}`,
            companySelectToDelete: 'Seleziona un’agenzia da eliminare.',
            companyDeleted: 'Agenzia eliminata.',
            activityNameEmpty: 'Inserisci un nome per l’attività.',
            activityExists: 'Attività già presente. Selezionala e usa la matita.',
            activitySaved: name => `Attività salvata: ${name}`,
            activitySelectToOverwrite: 'Seleziona l’attività da sovrascrivere.',
            activityNotFound: 'Attività selezionata non trovata.',
            activityOverwritten: name => `Attività sovrascritta: ${name}`,
            activitySelectFirst: 'Seleziona prima un’attività.',
            activityNameRequired: 'Il nome dell’attività non può essere vuoto.',
            activityDuplicate: 'Esiste già un’attività con questo nome.',
            activityRenamed: name => `Attività rinominata: ${name}`,
            activitySelectToDelete: 'Seleziona un’attività da eliminare.',
            activityDeleted: 'Attività eliminata.',
            selectCompany: 'Seleziona un’agenzia.',
            selectActivity: 'Seleziona un’attività.',
            selectBoth: 'Seleziona Agenzia e Attività.',
            configurationNotFound: 'Configurazione non trovata.',
            applyPartial: failures => `Compilazione parziale: ${failures} campo/i non compilato/i.`,
            applied: (company, activity) => `${company} + ${activity} applicato.`,
            favoriteRemoved: 'Preferito rimosso.',
            favoriteAdded: 'Preferito aggiunto.',
            favoriteCleared: 'Selezione preferito azzerata.',
            favoriteSelected: 'Preferito selezionato. Premi Compila Ticket per applicarlo.',
            incidentNotFound: 'Numero Incident non trovato.',
            previewEmpty: 'L’anteprima è vuota.',
            previewEmptyRefresh: 'L’anteprima è vuota. Premi ↻ per generare la risposta.',
            previewEmptyHold: 'L’anteprima risposta è vuota: inserisci i commenti visibili al cliente.',
            previewEmptyClosure: 'L’anteprima risposta è vuota.',
            combinationResponseSaved: 'Risposta salvata per la combinazione.',
            titleMissing: 'Short Description vuota: titolo non disponibile.',
            responseSaved: title => `Risposta salvata: ${title}`,
            responseNotFound: 'Risposta salvata non trovata.',
            responseNameRequired: 'Il nome della risposta non può essere vuoto.',
            responseDuplicate: 'Esiste già una risposta salvata con questo nome.',
            responseRenamed: title => `Risposta rinominata: ${title}`,
            responseLoaded: title => `Risposta caricata: ${title}`,
            copyFailed: 'Impossibile copiare il testo.',
            responseCopied: 'Risposta copiata negli appunti.',
            incidentCopied: 'Incident copiato negli appunti.',
            linkCopied: 'Numero e link copiati.',
            referenceDeleted: number => `Riferimento ${number} eliminato.`,
            invalidTicketLink: 'Link del ticket non valido o non appartenente a ServiceNow.',
            holdDurationInvalid: 'Seleziona una durata On Hold valida.',
            holdReasonMissing: 'Seleziona il motivo On Hold.',
            holdPrepared: expiration => `On Hold predisposto fino al ${expiration}. Verifica i campi e premi SAVE in AMP. Per annullare, clicca di nuovo sul pulsante.`,
            closureStatusInvalid: 'Seleziona uno stato di chiusura valido.',
            closureCodeMissing: 'Seleziona un codice di risoluzione per lo stato Resolved.',
            closurePrepared: (status, code) => `Ticket predisposto per lo stato ${status}${code ? ` con codice di risoluzione “${code}”` : ''}. Verifica i campi e premi SAVE in AMP. Per annullare, clicca di nuovo sul pulsante.`,
            ticketFilledKnown: state => `Ticket ${state}: configurazione ripristinata. Compila Ticket chiede conferma prima di sovrascrivere.`,
            ticketFilledUnknown: state => `Ticket ${state}: Compila Ticket chiede conferma prima di sovrascrivere i campi.`,
            ticketClosedKnown: state => `Ticket ${state}: configurazione ripristinata. Compila Ticket disabilitato.`,
            ticketClosedUnknown: state => `Ticket ${state}: gestione configurazioni disponibile. Compila Ticket disabilitato.`,
            exported: 'Configurazione esportata.',
            imported: 'Configurazione importata.',
            importInvalid: 'File JSON non valido.'
        },
        // Finestre di conferma (OK / Annulla)
        confirm: {
            replaceProfile: (current, next) => `Sostituire il profilo "${current}" con "${next}"?`,
            overwriteCompany: (name, company, group) => `Sovrascrivere l'agenzia "${name}" con i dati presenti nel ticket?\n\n`
                + `Company: ${company}\nAssignment Group: ${group}`,
            overwriteActivity: (name, values) => `Sovrascrivere l'attività "${name}" con i dati presenti nel ticket?\n\n`
                + `Category: ${values.category.displayValue}\nImpact: ${values.impact}\n`
                + `Urgency: ${values.urgency}\nTime Worked: ${values.timeWorked} minuti\n`
                + `TAG Short Description: ${values.shortDescriptionPrefix || 'Nessuno'}`,
            overwriteTicket: (company, activity) => `Il ticket è già stato compilato.\n\nSovrascrivere i campi con "${company} + ${activity}"?`,
            deleteCompany: name => `Eliminare l'agenzia "${name}"?`,
            deleteActivity: name => `Eliminare l'attività "${name}"?`,
            deleteResponse: title => `Eliminare la risposta "${title}"?`,
            deleteReference: number => `Eliminare il riferimento a ${number}?`,
            undoHold: 'Vuoi annullare la predisposizione dell’On Hold e ripristinare i valori precedenti dei campi? La pagina verrà ricaricata',
            undoClosure: 'Vuoi annullare la predisposizione della chiusura e ripristinare i valori precedenti dei campi? La pagina verrà ricaricata',
            importReplace: 'L’importazione sostituirà la configurazione locale.\n\nContinuare?'
        },
        // Richieste di un testo
        prompt: {
            activityName: 'Nome attività:',
            activityRename: 'Nuovo nome attività:',
            responseRename: 'Nuovo nome risposta:'
        },
        tour: {
            close: 'Chiudi',
            prev: 'Indietro',
            next: 'Avanti',
            finish: 'Fine',
            demoCompany: 'Agenzia dimostrativa',
            demoActivity: 'Attività dimostrativa',
            demoResponse: 'Risposta dimostrativa',
            // Titolo e testo di ogni passo (il numero viene aggiunto in automatico).
            // L'ordine dei passi e l'elemento indicato sono in TOUR_STEPS.
            steps: {
                profile: { title: 'Profilo personale', text: 'Il pulsante con l’omino memorizza il tuo nome da Assigned to: serve per la presa in carico del ticket e per la firma delle risposte.' },
                company: { title: 'Agenzia', text: 'Scegli un’Agenzia memorizzata (puoi scrivere per cercarla). Con + ne salvi una nuova da Company e Assignment Group del ticket; quando è selezionata, i pulsanti accanto la aggiornano con i valori del ticket o la eliminano.' },
                activity: { title: 'Attività', text: 'Funziona come l’Agenzia e memorizza Category, Impact, Urgency, Time Worked e l’eventuale tag FWD della Short Description. Quando è selezionata puoi anche rinominarla, senza perdere Preferiti e risposte collegate.' },
                apply: { title: 'Compila Ticket', text: 'Applica al ticket la combinazione Agenzia–Attività selezionata. Su un ticket già compilato chiede conferma prima di sovrascrivere; sui ticket chiusi non è disponibile.' },
                favorites: { title: 'Preferiti', text: 'La stella aggiunge o toglie dai Preferiti la combinazione selezionata. Dall’elenco Preferiti la richiami con un clic; poi premi Compila Ticket per applicarla.' },
                settings: { title: 'Impostazioni', text: 'Dall’ingranaggio puoi raggruppare i Preferiti per Agenzia o Attività, consultare tutte le configurazioni salvate ed esportarle o importarle in un file JSON.' },
                preview: { title: 'Anteprima risposta', text: 'Il testo è sempre modificabile. I pulsanti in alto lo salvano per la combinazione Agenzia–Attività, lo archiviano tra le Risposte salvate o ricaricano la risposta della combinazione.' },
                useful: { title: 'Risposte salvate', text: 'L’archivio delle risposte riutilizzabili: un clic carica il testo nell’Anteprima. Passando il cursore su una risposta compaiono il cestino per eliminarla e la matita per rinominarla.' },
                copy: { title: 'Copia', text: 'Risposta copia il testo con la tua firma, il numero Incident e la severity; Incident copia solo il riferimento con la severity.' },
                closure: { title: 'Predisponi Chiusura', text: 'Apre il riquadro di chiusura: scegli lo stato (Resolved o Canceled) e il Resolution Code, obbligatorio per Resolved.' },
                closureConfirm: { title: 'Conferma e annulla chiusura', text: 'Conferma chiusura compila stato, codice, note di risoluzione e commento al cliente. Il pulsante diventa “→ Save”: il SAVE in AMP resta manuale e un nuovo clic sul pulsante annulla, ripristinando i valori precedenti.', buttonLabel: 'Chiusura predisposta → Save' },
                hold: { title: 'On Hold', text: 'Il pulsante pausa apre il riquadro On Hold: scegli il motivo della sospensione e la durata, da 1 a 7 giorni.' },
                holdConfirm: { title: 'Conferma e annulla On Hold', text: 'Predisponi On Hold compila stato, motivo, scadenza, commento al cliente e note interne. Come per la chiusura il pulsante diventa “→ Save”: il SAVE resta manuale e un nuovo clic annulla.', buttonLabel: 'On Hold predisposto → Save' },
                tickets: { title: 'Ticket aperti', text: 'La coda dei ticket in lavorazione: un clic apre il ticket; passando il cursore compaiono copia, apertura in nuova scheda ed eliminazione del riferimento. La pausa indica un On Hold, il punto esclamativo rosso un On Hold scaduto. Sull’elenco Incident il pannello mostra solo questa coda.' }
            }
        }
    };

    // Icone SVG usate nel pannello (inserite con ICONS.nome).
    const ICONS = {
        user: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="8" r="4"></circle><path d="M4 21c0-4.4 3.6-8 8-8s8 3.6 8 8"></path></svg>',
        save: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 5v14M5 12h14"></path></svg>',
        preparedArrow: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 12h15"></path><path d="m13 6 6 6-6 6"></path></svg>',
        edit: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M18.8 7.2A8 8 0 0 0 5.5 8"></path><path d="M18.8 7.2L18 3.5M18.8 7.2l-3.8.6"></path><path d="M5.2 16.8A8 8 0 0 0 18.5 16"></path><path d="M5.2 16.8l.8 3.7M5.2 16.8l3.8-.6"></path></svg>',
        rename: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 20l4.5-1 10-10-3.5-3.5-10 10z"></path><path d="M13.8 6.7l3.5 3.5M4 20h6"></path></svg>',
        remove: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"></path></svg>',
        trash: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 4h8"></path><path d="M5 7h14"></path><path d="M7 7l1 12h8l1-12"></path></svg>',
        external: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M14 4h6v6M20 4l-9 9"></path><path d="M18 13v6H5V6h6"></path></svg>',
        download: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v12M7 10l5 5 5-5M4 20h16"></path></svg>',
        upload: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 16V4M7 9l5-5 5 5M4 20h16"></path></svg>',
        star: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3.5l2.7 5.5 6.1.9-4.4 4.3 1 6.1-5.4-2.9-5.4 2.9 1-6.1-4.4-4.3 6.1-.9z"></path></svg>',
        archive: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 5h16v4H4z"></path><path d="M5 9v10h14V9"></path><path d="M9 13h6"></path></svg>',
        listChecks: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m3.5 6.5 1.5 1.5 2.5-3"></path><path d="M10 7h11"></path><path d="m3.5 12.5 1.5 1.5 2.5-3"></path><path d="M10 13h11"></path><path d="m3.5 18.5 1.5 1.5 2.5-3"></path><path d="M10 19h11"></path></svg>',
        refresh: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M20 6v5h-5"></path><path d="M19 11a7.5 7.5 0 1 0-1.7 5"></path></svg>',
        copy: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="8" y="8" width="11" height="11" rx="2"></rect><path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h2"></path></svg>',
        check: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 12.5l4.2 4.2L19 7"></path></svg>',
        help: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9"></circle><path d="M9.8 9a2.4 2.4 0 1 1 3.7 2c-1 .6-1.5 1.1-1.5 2.2"></path><path d="M12 17h.01"></path></svg>',
        gear: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="3"></circle><path d="M19 12a7 7 0 0 0-.1-1l2-1.5-2-3.4-2.4 1a7 7 0 0 0-1.7-1L14.5 3h-5l-.4 3.1a7 7 0 0 0-1.7 1l-2.4-1-2 3.4L5 11a7 7 0 0 0 0 2l-2 1.5 2 3.4 2.4-1a7 7 0 0 0 1.7 1l.4 3.1h5l.4-3.1a7 7 0 0 0 1.7-1l2.4 1 2-3.4-2-1.5a7 7 0 0 0 .1-1z"></path></svg>',
        left: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M15 5l-7 7 7 7"></path></svg>',
        right: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 5l7 7-7 7"></path></svg>',
        up: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 15l7-7 7 7"></path></svg>',
        down: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 9l7 7 7-7"></path></svg>',
        warning: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3 22 21H2L12 3Z"></path><path d="M12 9v5M12 17h.01"></path></svg>',
        pause: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="5.5" y="4" width="5" height="16"></rect><rect x="13.5" y="4" width="5" height="16"></rect></svg>'
    };

    // ======================================================================
    // 1. CONFIGURAZIONE E STATO
    // Costanti dell'applicazione, mappa dei campi ServiceNow e stato condiviso.
    // ======================================================================

    const APP = {
        name: 'SNOW Assistant',
        version: 3,
        storageKey: 'snow_ticket_assistant_v1',
        openTicketsStorageKey: 'snow_ticket_assistant_open_tickets_v1',
        pendingTicketSessionKey: 'snow_ticket_assistant_pending_ticket_v1',
        panelId: 'snow-assistant-panel',
        styleId: 'snow-assistant-style'
    };
    const FIELDS = {
        company: { name: 'company', displayId: 'sys_display.incident.company' },
        assignmentGroup: { name: 'assignment_group', displayId: 'sys_display.incident.assignment_group' },
        assignedTo: { name: 'assigned_to', displayId: 'sys_display.incident.assigned_to' },
        category: { name: 'x_avana_platform_category_value', displayId: 'sys_display.incident.x_avana_platform_category_value' },
        impact: { name: 'impact', elementId: 'incident.impact' },
        urgency: { name: 'urgency', elementId: 'incident.urgency' },
        shortDescription: { name: 'short_description', elementId: 'incident.short_description' },
        incidentNumber: { elementId: 'sys_readonly.incident.number' },
        state: { name: 'state', elementId: 'incident.state' },
        holdReason: { name: 'hold_reason', elementId: 'incident.hold_reason' },
        holdExpiration: { name: 'u_on_hold_expiration', elementId: 'incident.u_on_hold_expiration' },
        resolutionCode: { name: 'close_code', elementId: 'incident.close_code' },
        resolutionNotes: { name: 'close_notes', elementId: 'incident.close_notes' },
        additionalComments: { elementId: 'activity-stream-comments-textarea', fallbackElementId: 'incident.comments' },
        workNotes: { name: 'work_notes', elementId: 'activity-stream-work_notes-textarea', fallbackElementId: 'incident.work_notes' },
        timeWorkedMinutes: {
            selector: 'input[data-type="timer"][id$="_min"]',
            storedElementId: 'incident.time_worked'
        }
    };
    const SEVERITY_MAP = {
        '3-3': 'std',
        '1-1': 'p1',
        '1-2': 'p2'
    };
    // Nomi degli stati del ticket mostrati nei messaggi.
    const TICKET_STATE_LABELS = { '1': 'New', '2': 'In Progress', '3': 'On Hold', '6': 'Resolved', '7': 'Closed', '8': 'Canceled' };
    // Ticket già compilati ma aperti (Compila Ticket chiede conferma) e ticket chiusi (Compila Ticket disabilitato).
    const FILLED_OPEN_STATES = ['2', '3'];
    const CLOSED_STATES = ['6', '7', '8'];
    // Voci delle tendine di chiusura e On Hold: value = codice in ServiceNow, label = testo mostrato.
    const CLOSURE_STATUSES = [
        { value: '6', label: 'Resolved', selected: true },
        { value: '8', label: 'Canceled' }
    ];
    const RESOLUTION_CODES = [
        { value: 'Solved (Permanently)', label: 'Solved (Permanently)', selected: true },
        { value: 'Solved by Third Party Vendor', label: 'Solved by Third Party Vendor' },
        { value: 'Closed on Initial', label: 'Closed on Initial' },
        { value: 'Not Reproducible (Alert Cleared)', label: 'Not Reproducible (Alert Cleared)' },
        { value: 'Not Solved (Not Reproducible)', label: 'Not Solved (Not Reproducible)' },
        { value: 'Caused by change', label: 'Caused by change' },
        { value: 'Not Solved (Too Costly)', label: 'Not Solved (Too Costly)' },
        { value: 'No Client Response', label: 'No Client Response' },
        { value: 'Solved (Work Around)', label: 'Solved (Work Around)' },
        { value: 'Duplicate', label: 'Duplicate' },
        { value: 'Out of Scope', label: 'Out of Scope' },
        { value: '', label: TEXTS.closure.noCode }
    ];
    // Le voci disabled sono visibili ma non selezionabili.
    const HOLD_REASONS = [
        { value: '5', label: 'Awaiting Change', disabled: true },
        { value: '3', label: 'Awaiting Problem Resolution', disabled: true },
        { value: '6', label: 'Awaiting Third Party', disabled: true },
        { value: '7', label: 'Awaiting Client Action/Update' },
        { value: '8', label: 'Awaiting Client Resolution' },
        { value: '9', label: 'Awaiting Review' }
    ];
    const pendingResolutionSessionKey = `${APP.storageKey}:pending-resolution`;
    let snWindow = null;
    let snDocument = null;

    // ======================================================================
    // 2. SALVATAGGIO DATI
    // Lettura e scrittura di configurazioni e ticket in localStorage e sessionStorage.
    // ======================================================================

    function createDefaultData() {
        const now = new Date().toISOString();
        return {
            version: APP.version,
            created: now,
            updated: now,
            profile: { assignedTo: { value: '', displayValue: '' } },
            companies: {},
            activities: {},
            favorites: [],
            recent: [],
            combinationResponses: {},
            usefulResponses: {},
            preferences: { favoritesGroupBy: 'none' }
        };
    }
    function normalizeReference(reference) {
        if (!reference) return { value: '', displayValue: '' };
        if (typeof reference === 'string') return { value: '', displayValue: reference };
        return {
            value: reference.value || '',
            displayValue: reference.displayValue || ''
        };
    }
    function normalizeData(data) {
        const defaults = createDefaultData();
        if (!data || typeof data !== 'object') return defaults;
        return {
            version: Number(data.version) || APP.version,
            created: data.created || defaults.created,
            updated: data.updated || defaults.updated,
            profile: { assignedTo: normalizeReference(data.profile?.assignedTo) },
            companies: data.companies && typeof data.companies === 'object' ? data.companies : {},
            activities: data.activities && typeof data.activities === 'object' ? data.activities : {},
            favorites: Array.isArray(data.favorites) ? data.favorites : [],
            recent: Array.isArray(data.recent) ? data.recent : [],
            combinationResponses: data.combinationResponses && typeof data.combinationResponses === 'object' ? data.combinationResponses : {},
            usefulResponses: data.usefulResponses && typeof data.usefulResponses === 'object' ? data.usefulResponses : {},
            preferences: { favoritesGroupBy: ['company', 'activity'].includes(data.preferences?.favoritesGroupBy) ? data.preferences.favoritesGroupBy : 'none' }
        };
    }
    function loadData() {
        try {
            const raw = localStorage.getItem(APP.storageKey);
            if (!raw) {
                const data = createDefaultData();
                saveData(data);
                return data;
            }
            return normalizeData(JSON.parse(raw));
        } catch (error) {
            console.error(`[${APP.name}] Errore lettura storage:`, error);
            const data = createDefaultData();
            saveData(data);
            return data;
        }
    }
    function saveData(data) {
        data.version = APP.version;
        data.updated = new Date().toISOString();
        localStorage.setItem(APP.storageKey, JSON.stringify(data));
    }
    function loadOpenTickets() {
        try {
            const raw = localStorage.getItem(APP.openTicketsStorageKey);
            return raw ? JSON.parse(raw) : {};
        } catch {
            return {};
        }
    }
    function saveOpenTickets(data) {
        localStorage.setItem(APP.openTicketsStorageKey, JSON.stringify(data));
    }
    function savePendingTicketSelection(companyId, activityId) {
        if (!companyId || !activityId) return;
        sessionStorage.setItem(APP.pendingTicketSessionKey, JSON.stringify({ companyId, activityId }));
    }
    function loadPendingTicketSelection() {
        try {
            const raw = sessionStorage.getItem(APP.pendingTicketSessionKey);
            return raw ? JSON.parse(raw) : null;
        } catch {
            return null;
        }
    }
    function clearPendingTicketSelection() {
        sessionStorage.removeItem(APP.pendingTicketSessionKey);
    }
    function markPendingResolutionSave() {
        const number = readIncidentNumber();
        if (number) sessionStorage.setItem(pendingResolutionSessionKey, number);
    }
    function clearPendingResolutionSave() {
        sessionStorage.removeItem(pendingResolutionSessionKey);
    }

    // ======================================================================
    // 3. ADATTATORE SERVICENOW
    // Unico punto che legge e scrive la pagina AMP (g_form, campi del form, finestra).
    // Nell'add-in Outlook questo blocco verrà sostituito da uno che legge la mail
    // e comunica con la Function App: il resto dello script resta invariato.
    // ======================================================================

    function findServiceNowContext() {
        if (document.getElementById(FIELDS.company.displayId) || typeof window.g_form !== 'undefined') {
            return { win: window, doc: document };
        }
        const frame = document.getElementById('gsft_main') || document.querySelector('iframe[name="gsft_main"]');
        if (!frame) return null;
        try {
            return { win: frame.contentWindow, doc: frame.contentDocument };
        } catch (error) {
            console.warn(`[${APP.name}] Accesso iframe fallito:`, error);
            return null;
        }
    }
    function getGForm() {
        try { return snWindow?.g_form || null; } catch { return null; }
    }
    // I riferimenti ai ticket devono puntare esclusivamente all'istanza ServiceNow prevista.
    // I dati in localStorage non sono considerati attendibili: possono essere modificati manualmente.
    function getSafeServiceNowTicketUrl(value) {
        if (typeof value !== 'string' || !value.trim()) return '';
        try {
            const url = new URL(value);
            if (url.protocol !== 'https:' || url.origin !== 'https://amp.service-now.com' ||
                url.username || url.password) return '';
            return url.href;
        } catch {
            return '';
        }
    }
    function readReference(field) {
        let value = '';
        const gForm = getGForm();
        if (gForm) {
            try { value = gForm.getValue(field.name) || ''; } catch (error) { console.warn(error); }
        }
        const displayValue = snDocument.getElementById(field.displayId)?.value || '';
        return { value, displayValue: displayValue.trim() };
    }
    function readChoice(field) {
        const gForm = getGForm();
        if (gForm) {
            try { return gForm.getValue(field.name) || ''; } catch (error) { console.warn(error); }
        }
        return snDocument.getElementById(field.elementId)?.value || '';
    }
    function readIncidentNumber() {
        return String(snDocument.getElementById(FIELDS.incidentNumber.elementId)?.value || '').trim();
    }
    function readShortDescription() {
        return String(snDocument.getElementById(FIELDS.shortDescription.elementId)?.value || '').trim();
    }
    function readHoldExpirationValue() {
        const gForm = getGForm();
        if (gForm) {
            try { return String(gForm.getValue(FIELDS.holdExpiration.name) || '').trim(); }
            catch (error) { console.warn(error); }
        }
        return String(snDocument.getElementById(FIELDS.holdExpiration.elementId)?.value || '').trim();
    }
    function readTimeWorkedMinutes() {
        const state = readChoice(FIELDS.state);
        const useStoredTimer = state === '2' || state === '6';

        if (useStoredTimer) {
            const storedTimer = snDocument.getElementById(FIELDS.timeWorkedMinutes.storedElementId);
            const storedValue = String(storedTimer?.value || '').trim();
            const match = storedValue.match(/^(\d+):(\d{2}):(\d{2})$/);
            if (!match) return '';

            const minutes = Number.parseInt(match[2], 10);

            if (Number.isNaN(minutes)) return '';

            return minutes;
        }

        const element = snDocument.querySelector(FIELDS.timeWorkedMinutes.selector);
        if (!element) return '';
        const value = String(element.value || '').trim();
        if (!value) return '';
        const parsed = Number.parseInt(value, 10);
        return Number.isNaN(parsed) ? value : parsed;
    }
    function extractShortDescriptionPrefix() {
        const match = readShortDescription().match(/^\s*(\[FWD\])/i);
        return match ? '[FWD]' : '';
    }
    function dispatchEvents(element) {
        ['input', 'change', 'blur'].forEach(name => {
            element.dispatchEvent(new Event(name, { bubbles: true }));
        });
    }
    function writeReference(field, reference) {
        const normalized = normalizeReference(reference);
        const gForm = getGForm();
        if (gForm && normalized.value) {
            try {
                gForm.setValue(field.name, normalized.value, normalized.displayValue);
                return true;
            } catch (error) { console.warn(error); }
        }
        const element = snDocument.getElementById(field.displayId);
        if (!element) return false;
        element.value = normalized.displayValue;
        dispatchEvents(element);
        return true;
    }
    function writeChoice(field, value) {
        const normalized = String(value ?? '');
        const gForm = getGForm();
        if (gForm) {
            try {
                gForm.setValue(field.name, normalized);
                return true;
            } catch (error) { console.warn(error); }
        }
        const element = snDocument.getElementById(field.elementId);
        if (!element) return false;
        element.value = normalized;
        dispatchEvents(element);
        return true;
    }
    function writeTimeWorkedMinutes(value) {
        const element = snDocument.querySelector(FIELDS.timeWorkedMinutes.selector);
        const parsed = Number.parseInt(value, 10);
        if (!element || Number.isNaN(parsed)) return false;
        element.value = String(parsed).padStart(2, '0');
        dispatchEvents(element);
        return true;
    }
    function applyShortDescriptionPrefix(prefix) {
        if (!prefix) return true;
        const element = snDocument.getElementById(FIELDS.shortDescription.elementId);
        if (!element) return false;
        const clean = String(element.value || '').replace(/^\s*\[[^\]]+\]\s*/, '').trim();
        element.value = `${prefix} ${clean}`.trim();
        dispatchEvents(element);
        return true;
    }
    function writeTextareaElement(elementId, value) {
        const element = snDocument.getElementById(elementId);
        if (!element) return false;
        element.value = value;
        dispatchEvents(element);
        return true;
    }
    function writeAdditionalComments(value) {
        const element = snDocument.getElementById(FIELDS.additionalComments.elementId) || snDocument.getElementById(FIELDS.additionalComments.fallbackElementId);
        if (!element) return false;
        element.value = value;
        dispatchEvents(element);
        return true;
    }
    function writeWorkNotes(value) {
        const normalized = String(value ?? '');
        const gForm = getGForm();
        let written = false;
        if (gForm) {
            try { gForm.setValue(FIELDS.workNotes.name, normalized); written = true; }
            catch (error) { console.warn(error); }
        }
        const element = snDocument.getElementById(FIELDS.workNotes.elementId) || snDocument.getElementById(FIELDS.workNotes.fallbackElementId);
        if (element) {
            try { element.value = normalized; dispatchEvents(element); written = true; }
            catch (error) { console.warn(error); }
        }
        return written;
    }
    function writeOnHoldComments(value) {
        // On Hold compila entrambe le attività: note interne e commenti visibili al cliente.
        const commentsWritten = writeAdditionalComments(value);
        const workNotesWritten = writeWorkNotes(value);
        return commentsWritten && workNotesWritten;
    }
    function writeHoldExpiration(value) {
        const gForm = getGForm();
        if (gForm) {
            try {
                gForm.setValue(FIELDS.holdExpiration.name, value);
                return true;
            } catch (error) { console.warn(error); }
        }
        const element = snDocument.getElementById(FIELDS.holdExpiration.elementId);
        if (!element) return false;
        element.value = value;
        dispatchEvents(element);
        return true;
    }
    function undoPreparedAction(type) {
        const isHold = type === 'hold';
        const message = isHold ? TEXTS.confirm.undoHold : TEXTS.confirm.undoClosure;
        if (!snWindow.confirm(message)) return;

        // Il refresh ripristina i valori presenti sul server prima della predisposizione.
        snWindow.location.reload();
    }
    function bindServiceNowEvents() {
        snDocument.getElementById(FIELDS.impact.elementId)?.addEventListener('change', regenerateResponsePreview);
        snDocument.getElementById(FIELDS.urgency.elementId)?.addEventListener('change', regenerateResponsePreview);

        const captureBeforeTicketSave = event => {
            // Mantiene il trigger già usato per l'inserimento in coda:
            // intercetta il comando di salvataggio per conservare la selezione.
            const saveControl = event.target?.closest?.(
                '[id^="sysverb_update"], [id^="sysverb_insert"], ' +
                '[name^="sysverb_update"], [name^="sysverb_insert"], ' +
                'button[type="submit"], input[type="submit"], [data-action*="update"]'
            );
            const saveId = saveControl?.id || '';
            const saveName = saveControl?.getAttribute?.('name') || '';
            const isSaveAction = /^sysverb_(update|insert)(_|$)/i.test(saveId) ||
                /^sysverb_(update|insert)(_|$)/i.test(saveName) ||
                Boolean(saveControl?.matches?.('button[type="submit"], input[type="submit"], [data-action*="update"]'));

            if (isSaveAction || event.type === 'submit') captureTicketSelectionFromSelectors();
        };
        const capturePendingTicketChangesBeforeUnload = () => {
            // Stesso trigger che completa il flusso di inserimento in coda:
            // al cambio pagina conserva la selezione e, se la chiusura è predisposta,
            // registra l'Incident da verificare dopo il reload.
            captureTicketSelectionFromSelectors();
            if (readChoice(FIELDS.state) === '6') markPendingResolutionSave();
        };
        snDocument.addEventListener('click', captureBeforeTicketSave, true);
        snDocument.addEventListener('submit', captureBeforeTicketSave, true);
        snWindow?.addEventListener('beforeunload', capturePendingTicketChangesBeforeUnload);
    }

    // ======================================================================
    // 4. CONFIGURAZIONI: PROFILO, AGENZIE, ATTIVITÀ
    // Salvataggio, modifica ed eliminazione delle configurazioni e compilazione del ticket.
    // ======================================================================

    function createId(label) {
        const base = String(label || '')
            .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
            .toLowerCase().trim().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
        return `${base || 'config'}_${Date.now().toString(36)}`;
    }
    function combinationKey(companyId, activityId) {
        return `${companyId}::${activityId}`;
    }
    function readCurrentActivityValues(showErrors = true) {
        const values = {
            category: readReference(FIELDS.category),
            impact: readChoice(FIELDS.impact),
            urgency: readChoice(FIELDS.urgency),
            timeWorked: readTimeWorkedMinutes()
        };
        let message = '';
        if (!values.category.displayValue) message = TEXTS.status.fillCategory;
        else if (!values.impact) message = TEXTS.status.fillImpact;
        else if (!values.urgency) message = TEXTS.status.fillUrgency;
        else if (values.timeWorked === '') message = TEXTS.status.fillTimeWorked;
        if (message) {
            if (showErrors) showStatus(message, 'error');
            return null;
        }
        return {
            category: values.category,
            impact: String(values.impact),
            urgency: String(values.urgency),
            timeWorked: Number(values.timeWorked),
            shortDescriptionPrefix: extractShortDescriptionPrefix()
        };
    }
    function configureProfile() {
        const assignedTo = readReference(FIELDS.assignedTo);
        if (!assignedTo.displayValue) return showStatus(TEXTS.status.fillAssignedTo, 'error');
        const data = loadData();
        const current = data.profile.assignedTo.displayValue;
        if (current && current !== assignedTo.displayValue) {
            if (!confirm(TEXTS.confirm.replaceProfile(current, assignedTo.displayValue))) return;
        }
        data.profile.assignedTo = assignedTo;
        saveData(data);
        refreshPanel();
        regenerateResponsePreview();
        showStatus(TEXTS.status.profileSet(assignedTo.displayValue), 'success');
    }
    function saveNewCompany() {
        const company = readReference(FIELDS.company);
        const assignmentGroup = readReference(FIELDS.assignmentGroup);
        if (!company.displayValue) return showStatus(TEXTS.status.fillCompany, 'error');
        if (!assignmentGroup.displayValue) return showStatus(TEXTS.status.fillAssignmentGroup, 'error');
        const data = loadData();
        const sameName = Object.values(data.companies).some(item =>
            item.name.toLowerCase() === company.displayValue.toLowerCase()
        );
        if (sameName) return showStatus(TEXTS.status.companyExists, 'error');
        const id = createId(company.displayValue);
        data.companies[id] = { name: company.displayValue, company, assignmentGroup };
        saveData(data);
        refreshPanel(id, null);
        showStatus(TEXTS.status.companySaved(company.displayValue), 'success');
    }
    function overwriteSelectedCompany() {
        const companyId = getPanelElement('snow-company-select')?.value;
        if (!companyId) return showStatus(TEXTS.status.companySelectToOverwrite, 'error');
        const company = readReference(FIELDS.company);
        const assignmentGroup = readReference(FIELDS.assignmentGroup);
        if (!company.displayValue) return showStatus(TEXTS.status.fillCompany, 'error');
        if (!assignmentGroup.displayValue) return showStatus(TEXTS.status.fillAssignmentGroup, 'error');
        const data = loadData();
        const saved = data.companies[companyId];
        if (!saved) return showStatus(TEXTS.status.companyNotFound, 'error');
        if (!confirm(TEXTS.confirm.overwriteCompany(saved.name, company.displayValue, assignmentGroup.displayValue))) return;
        data.companies[companyId] = {
            name: company.displayValue,
            company,
            assignmentGroup
        };
        saveData(data);
        refreshPanel(companyId, null);
        showStatus(TEXTS.status.companyOverwritten(company.displayValue), 'success');
    }
    function deleteSelectedCompany() {
        const select = getPanelElement('snow-company-select');
        if (!select?.value) return showStatus(TEXTS.status.companySelectToDelete, 'error');
        const data = loadData();
        const company = data.companies[select.value];
        if (!company || !confirm(TEXTS.confirm.deleteCompany(company.name))) return;
        const id = select.value;
        delete data.companies[id];
        data.favorites = data.favorites.filter(item => item.company !== id);
        data.recent = data.recent.filter(item => item.company !== id);
        Object.keys(data.combinationResponses).forEach(key => {
            if (data.combinationResponses[key].company === id) delete data.combinationResponses[key];
        });
        saveData(data);
        refreshPanel();
        showStatus(TEXTS.status.companyDeleted, 'success');
    }
    function saveNewActivity() {
        const values = readCurrentActivityValues(true);
        if (!values) return;
        const nameInput = prompt(TEXTS.prompt.activityName, values.category.displayValue);
        if (nameInput === null) return;
        const name = nameInput.trim();
        if (!name) return showStatus(TEXTS.status.activityNameEmpty, 'error');
        const data = loadData();
        const sameName = Object.values(data.activities).some(item =>
            item.name.toLowerCase() === name.toLowerCase()
        );
        if (sameName) return showStatus(TEXTS.status.activityExists, 'error');
        const id = createId(name);
        data.activities[id] = { name, ...values };
        saveData(data);
        refreshPanel(null, id);
        showStatus(TEXTS.status.activitySaved(name), 'success');
    }
    function overwriteSelectedActivity() {
        const activityId = getPanelElement('snow-activity-select')?.value;
        if (!activityId) return showStatus(TEXTS.status.activitySelectToOverwrite, 'error');
        const values = readCurrentActivityValues(true);
        if (!values) return;
        const data = loadData();
        const activity = data.activities[activityId];
        if (!activity) return showStatus(TEXTS.status.activityNotFound, 'error');
        if (!confirm(TEXTS.confirm.overwriteActivity(activity.name, values))) return;
        data.activities[activityId] = { name: activity.name, ...values };
        saveData(data);
        refreshPanel(null, activityId);
        regenerateResponsePreview();
        showStatus(TEXTS.status.activityOverwritten(activity.name), 'success');
    }
    function renameSelectedActivity() {
        const activityId = getPanelElement('snow-activity-select')?.value;
        if (!activityId) return showStatus(TEXTS.status.activitySelectFirst, 'error');
        const data = loadData();
        const activity = data.activities[activityId];
        if (!activity) return showStatus(TEXTS.status.activityNotFound, 'error');
        const newName = prompt(TEXTS.prompt.activityRename, activity.name);
        if (newName === null) return;
        const cleanName = newName.trim();
        if (!cleanName) return showStatus(TEXTS.status.activityNameRequired, 'error');
        const duplicate = Object.entries(data.activities).some(([id, item]) =>
            id !== activityId && item.name.toLowerCase() === cleanName.toLowerCase()
        );
        if (duplicate) return showStatus(TEXTS.status.activityDuplicate, 'error');
        activity.name = cleanName;
        saveData(data);
        refreshPanel(null, activityId);
        showStatus(TEXTS.status.activityRenamed(cleanName), 'success');
    }
    function deleteSelectedActivity() {
        const select = getPanelElement('snow-activity-select');
        if (!select?.value) return showStatus(TEXTS.status.activitySelectToDelete, 'error');
        const data = loadData();
        const activity = data.activities[select.value];
        if (!activity || !confirm(TEXTS.confirm.deleteActivity(activity.name))) return;
        const id = select.value;
        delete data.activities[id];
        data.favorites = data.favorites.filter(item => item.activity !== id);
        data.recent = data.recent.filter(item => item.activity !== id);
        Object.keys(data.combinationResponses).forEach(key => {
            if (data.combinationResponses[key].activity === id) delete data.combinationResponses[key];
        });
        saveData(data);
        refreshPanel();
        showStatus(TEXTS.status.activityDeleted, 'success');
    }
    function applySelectedConfiguration() {
        const companyId = getPanelElement('snow-company-select')?.value;
        const activityId = getPanelElement('snow-activity-select')?.value;
        if (!companyId) return showStatus(TEXTS.status.selectCompany, 'error');
        if (!activityId) return showStatus(TEXTS.status.selectActivity, 'error');
        const data = loadData();
        const company = data.companies[companyId];
        const activity = data.activities[activityId];
        if (!company || !activity) return showStatus(TEXTS.status.configurationNotFound, 'error');
        if (!data.profile.assignedTo.displayValue) return showStatus(TEXTS.status.profileMissing, 'error');
        if (FILLED_OPEN_STATES.includes(readChoice(FIELDS.state))
            && !confirm(TEXTS.confirm.overwriteTicket(company.name, activity.name))) return;
        const results = [
            writeReference(FIELDS.company, company.company),
            writeReference(FIELDS.assignmentGroup, company.assignmentGroup),
            writeReference(FIELDS.assignedTo, data.profile.assignedTo),
            writeReference(FIELDS.category, activity.category),
            writeChoice(FIELDS.impact, activity.impact),
            writeChoice(FIELDS.urgency, activity.urgency),
            writeTimeWorkedMinutes(activity.timeWorked),
            applyShortDescriptionPrefix(activity.shortDescriptionPrefix || '')
        ];
        data.recent = data.recent.filter(item => !(item.company === companyId && item.activity === activityId));
        data.recent.unshift({ company: companyId, activity: activityId, timestamp: new Date().toISOString() });
        data.recent = data.recent.slice(0, 10);
        saveData(data);
        regenerateResponsePreview();
        const failures = results.filter(result => !result).length;
        if (!failures) savePendingTicketSelection(companyId, activityId);
        showStatus(
            failures ? TEXTS.status.applyPartial(failures) : TEXTS.status.applied(company.name, activity.name),
            failures ? 'warning' : 'success'
        );
    }

    // ======================================================================
    // 5. PREFERITI
    // Combinazioni Agenzia–Attività preferite e loro raggruppamento.
    // ======================================================================

    function toggleFavorite() {
        const company = getPanelElement('snow-company-select')?.value;
        const activity = getPanelElement('snow-activity-select')?.value;
        if (!company || !activity) return showStatus(TEXTS.status.selectBoth, 'error');
        const data = loadData();
        const index = data.favorites.findIndex(item => item.company === company && item.activity === activity);
        if (index >= 0) {
            data.favorites.splice(index, 1);
            showStatus(TEXTS.status.favoriteRemoved, 'success');
        } else {
            data.favorites.push({ company, activity });
            showStatus(TEXTS.status.favoriteAdded, 'success');
        }
        saveData(data);
        refreshFavoriteButton();
        refreshFavorites();
        refreshFavoritesGroupingControls();
        refreshUsefulResponses();
    }
    function refreshFavoriteButton() {
        const button = getPanelElement('snow-favorite-button');
        if (!button) return;
        const company = getPanelElement('snow-company-select')?.value;
        const activity = getPanelElement('snow-activity-select')?.value;
        const active = loadData().favorites.some(item => item.company === company && item.activity === activity);
        button.innerHTML = ICONS.star;
        button.classList.toggle('active', active);
    }
    function setFavoritesGrouping(value) {
        const data = loadData();
        data.preferences = data.preferences || {};
        data.preferences.favoritesGroupBy = ['company', 'activity'].includes(value) ? value : 'none';
        saveData(data);
        refreshFavoritesGroupingControls();
        refreshFavorites();
        closeSettingsMenu();
    }
    function refreshFavoritesGroupingControls() {
        const value = loadData().preferences?.favoritesGroupBy || 'none';
        ['none', 'company', 'activity'].forEach(option => {
            getPanelElement(`snow-favorites-group-${option}`)?.classList.toggle('active', value === option);
        });
    }
    function createFavoriteButton(favorite, company, activity, groupBy = 'none') {
        const button = document.createElement('button');
        button.className = 'snow-favorite-entry snow-style-list-item snow-style-button';
        const selectedCompany = getPanelElement('snow-company-select')?.value;
        const selectedActivity = getPanelElement('snow-activity-select')?.value;
        button.classList.toggle('selected', selectedCompany === favorite.company && selectedActivity === favorite.activity);
        const label = document.createElement('span');
        const strong = document.createElement('strong');
        if (groupBy === 'company') {
            strong.textContent = String(activity.name ?? '');
            label.appendChild(strong);
        } else if (groupBy === 'activity') {
            strong.textContent = String(company.name ?? '');
            label.appendChild(strong);
        } else {
            strong.textContent = String(company.name ?? '');
            const small = document.createElement('small');
            small.textContent = String(activity.name ?? '');
            label.append(strong, small);
        }
        button.appendChild(label);
        button.onclick = () => {
            const currentCompany = getPanelElement('snow-company-select')?.value;
            const currentActivity = getPanelElement('snow-activity-select')?.value;
            if (currentCompany === favorite.company && currentActivity === favorite.activity) {
                clearSelect('snow-company-select');
                clearSelect('snow-activity-select');
                showStatus(TEXTS.status.favoriteCleared, 'success');
            } else {
                getPanelElement('snow-company-select').value = favorite.company;
                getPanelElement('snow-activity-select').value = favorite.activity;
                refreshCustomSelect('snow-company-select');
                refreshCustomSelect('snow-activity-select');
                captureTicketSelectionFromSelectors();
                showStatus(TEXTS.status.favoriteSelected, 'success');
            }
            refreshSectionActions();
            refreshClearButtons();
            refreshFavoriteButton();
            refreshFavorites();
            regenerateResponsePreview();
            getPanelElement('snow-favorites-section')?.classList.remove('open');
            const favoritesArrow = getPanelElement('snow-favorites-arrow');
            if (favoritesArrow) favoritesArrow.innerHTML = ICONS.down;
        };
        return button;
    }
    function refreshFavorites() {
        const container = getPanelElement('snow-favorites-list');
        if (!container) return;
        const data = loadData();
        container.innerHTML = '';
        const valid = data.favorites.map(favorite => ({
            favorite,
            company: data.companies[favorite.company],
            activity: data.activities[favorite.activity]
        })).filter(item => item.company && item.activity);
        if (!valid.length) {
            container.innerHTML = `<div class="snow-empty">${TEXTS.panel.emptyFavorites}</div>`;
            return;
        }
        const groupBy = data.preferences?.favoritesGroupBy || 'none';
        if (groupBy === 'none') {
            valid.sort((a, b) => {
                const byCompany = a.company.name.localeCompare(b.company.name, 'it');
                return byCompany || a.activity.name.localeCompare(b.activity.name, 'it');
            }).forEach(item => container.appendChild(createFavoriteButton(item.favorite, item.company, item.activity, groupBy)));
            return;
        }
        const groups = new Map();
        valid.forEach(item => {
            const key = groupBy === 'company' ? item.favorite.company : item.favorite.activity;
            const label = groupBy === 'company' ? item.company.name : item.activity.name;
            if (!groups.has(key)) groups.set(key, { label, items: [] });
            groups.get(key).items.push(item);
        });
        [...groups.values()]
            .sort((a, b) => a.label.localeCompare(b.label, 'it'))
            .forEach(group => {
                group.items.sort((a, b) => {
                    const aLabel = groupBy === 'company' ? a.activity.name : a.company.name;
                    const bLabel = groupBy === 'company' ? b.activity.name : b.company.name;
                    return aLabel.localeCompare(bLabel, 'it');
                });
                const section = document.createElement('div');
                section.className = 'snow-favorite-group';
                const toggle = document.createElement('button');
                toggle.className = 'snow-favorite-group-toggle snow-style-group-toggle snow-style-button';
                const groupLabel = document.createElement('span');
                groupLabel.textContent = String(group.label ?? '');
                const groupArrow = document.createElement('span');
                groupArrow.className = 'snow-favorite-group-arrow';
                groupArrow.innerHTML = ICONS.down;
                toggle.append(groupLabel, groupArrow);
                const items = document.createElement('div');
                items.className = 'snow-favorite-group-items';
                toggle.addEventListener('click', () => {
                    const open = section.classList.toggle('open');
                    toggle.querySelector('.snow-favorite-group-arrow').innerHTML = open ? ICONS.up : ICONS.down;
                });
                group.items.forEach(item => items.appendChild(createFavoriteButton(item.favorite, item.company, item.activity, groupBy)));
                section.append(toggle, items);
                container.appendChild(section);
            });
    }

    // ======================================================================
    // 6. RISPOSTE
    // Anteprima risposta, risposte per combinazione, risposte salvate e copia negli appunti.
    // ======================================================================

    function getIncidentReference(showErrors = true) {
        const incident = readIncidentNumber();
        if (!incident) {
            if (showErrors) showStatus(TEXTS.status.incidentNotFound, 'error');
            return '';
        }
        const impact = readChoice(FIELDS.impact);
        const urgency = readChoice(FIELDS.urgency);
        const severity = SEVERITY_MAP[`${impact}-${urgency}`] || '';

        // La severity è facoltativa: se la combinazione Impact/Urgency non è mappata,
        // copia comunque il numero dell'Incident senza suffisso.
        return severity ? `${incident}-${severity}` : incident;
    }
    function buildBaseResponse(showErrors = true) {
        const profileName = loadData().profile.assignedTo.displayValue;
        if (!profileName) {
            if (showErrors) showStatus(TEXTS.status.profileMissing, 'error');
            return '';
        }
        return TEXTS.responses.defaultText.join('\n');
    }
    function getResponseForSelection() {
        const companyId = getPanelElement('snow-company-select')?.value;
        const activityId = getPanelElement('snow-activity-select')?.value;
        if (companyId && activityId) {
            const saved = loadData().combinationResponses[combinationKey(companyId, activityId)];
            if (saved?.response) return saved.response;
        }
        return buildBaseResponse(false);
    }
    function regenerateResponsePreview() {
        const preview = getPanelElement('snow-response-preview');
        if (preview) preview.value = getResponseForSelection();
    }
    function saveCombinationResponse() {
        const company = getPanelElement('snow-company-select')?.value;
        const activity = getPanelElement('snow-activity-select')?.value;
        const response = String(getPanelElement('snow-response-preview')?.value || '').trim();
        if (!company || !activity) return showStatus(TEXTS.status.selectBoth, 'error');
        if (!response) return showStatus(TEXTS.status.previewEmpty, 'error');
        const data = loadData();
        data.combinationResponses[combinationKey(company, activity)] = {
            company, activity, response, updated: new Date().toISOString()
        };
        saveData(data);
        showStatus(TEXTS.status.combinationResponseSaved, 'success');
    }
    function saveUsefulResponse() {
        const title = readShortDescription();
        const response = String(getPanelElement('snow-response-preview')?.value || '').trim();
        if (!title) return showStatus(TEXTS.status.titleMissing, 'error');
        if (!response) return showStatus(TEXTS.status.previewEmpty, 'error');
        const data = loadData();
        const existing = Object.entries(data.usefulResponses).find(([, item]) => item.title.toLowerCase() === title.toLowerCase());
        const id = existing?.[0] || createId(title);
        const now = new Date().toISOString();
        data.usefulResponses[id] = {
            title, response,
            created: existing?.[1]?.created || now,
            updated: now
        };
        saveData(data);
        refreshUsefulResponses();
        showStatus(TEXTS.status.responseSaved(title), 'success');
    }
    function renameUsefulResponse(id) {
        const data = loadData();
        const item = data.usefulResponses[id];
        if (!item) return showStatus(TEXTS.status.responseNotFound, 'error');
        const newTitle = prompt(TEXTS.prompt.responseRename, item.title);
        if (newTitle === null) return;
        const cleanTitle = newTitle.trim();
        if (!cleanTitle) return showStatus(TEXTS.status.responseNameRequired, 'error');
        const duplicate = Object.entries(data.usefulResponses).some(([otherId, response]) =>
            otherId !== id && response.title.toLowerCase() === cleanTitle.toLowerCase()
        );
        if (duplicate) return showStatus(TEXTS.status.responseDuplicate, 'error');
        item.title = cleanTitle;
        item.updated = new Date().toISOString();
        saveData(data);
        refreshUsefulResponses();
        showStatus(TEXTS.status.responseRenamed(cleanTitle), 'success');
    }
    function deleteUsefulResponse(id) {
        const data = loadData();
        const item = data.usefulResponses[id];
        if (!item || !confirm(TEXTS.confirm.deleteResponse(item.title))) return;
        delete data.usefulResponses[id];
        saveData(data);
        refreshUsefulResponses();
    }
    function refreshUsefulResponses() {
        const container = getPanelElement('snow-useful-list');
        if (!container) return;
        const data = loadData();
        const entries = Object.entries(data.usefulResponses).sort(([, a], [, b]) => a.title.localeCompare(b.title, 'it'));
        container.innerHTML = '';
        if (!entries.length) {
            container.innerHTML = `<div class="snow-empty">${TEXTS.panel.emptyResponses}</div>`;
            return;
        }
        entries.forEach(([id, item]) => {
            const row = document.createElement('div');
            row.className = 'snow-useful-row';
            const loadButton = document.createElement('button');
            loadButton.className = 'snow-useful-entry snow-style-list-item snow-style-button';
            loadButton.innerHTML = ICONS.copy;
            const title = document.createElement('span');
            title.textContent = String(item.title ?? '');
            loadButton.appendChild(title);
            // Comportamento one-shot: carica la risposta senza mantenere alcuno stato .selected toggle
            loadButton.addEventListener('click', () => {
                getPanelElement('snow-response-preview').value = item.response;
                showStatus(TEXTS.status.responseLoaded(item.title), 'success');
                getPanelElement('snow-useful-section')?.classList.remove('open');
                const usefulArrow = getPanelElement('snow-useful-arrow');
                if (usefulArrow) usefulArrow.innerHTML = ICONS.down;
            });
            const actions = document.createElement('div');
            actions.className = 'snow-useful-actions';
            const renameButton = document.createElement('button');
            renameButton.className = 'snow-useful-rename snow-style-button snow-style-icon-button';
            renameButton.innerHTML = ICONS.rename;
            renameButton.title = TEXTS.responses.rename;
            renameButton.setAttribute('aria-label', TEXTS.responses.rename);
            renameButton.addEventListener('click', () => renameUsefulResponse(id));
            const deleteButton = document.createElement('button');
            deleteButton.className = 'snow-useful-delete snow-style-button snow-style-icon-button snow-style-danger';
            deleteButton.innerHTML = ICONS.trash;
            deleteButton.title = TEXTS.responses.remove;
            deleteButton.setAttribute('aria-label', TEXTS.responses.remove);
            deleteButton.addEventListener('click', () => deleteUsefulResponse(id));
            // Le azioni seguono il comportamento dei ticket: appaiono al passaggio del cursore.
            // Ordine: cestino, matita.
            actions.append(deleteButton, renameButton);
            row.append(loadButton, actions);
            container.appendChild(row);
        });
    }
    async function copyText(text, successMessage) {
        if (!text) return;
        try {
            await navigator.clipboard.writeText(text);
            showStatus(successMessage, 'success');
        } catch {
            const textarea = document.createElement('textarea');
            textarea.value = text;
            textarea.style.position = 'fixed';
            textarea.style.opacity = '0';
            document.body.appendChild(textarea);
            textarea.focus();
            textarea.select();
            const copied = document.execCommand('copy');
            textarea.remove();
            showStatus(copied ? successMessage : TEXTS.status.copyFailed, copied ? 'success' : 'error');
        }
    }
    function copyBaseResponse() {
        const text = String(getPanelElement('snow-response-preview')?.value || '').trim();
        if (!text) return showStatus(TEXTS.status.previewEmptyRefresh, 'error');
        const profileName =loadData().profile.assignedTo.displayValue;
        const incident = getIncidentReference(true);
        if (!incident) return;
        const finalText = `${text}\n${profileName}\n${incident}`;
        copyText(finalText, TEXTS.status.responseCopied);
    }
    function copyIncidentWithSeverity() {
        copyText(getIncidentReference(true), TEXTS.status.incidentCopied);
    }

    // ======================================================================
    // 7. TICKET APERTI E MONITORAGGIO ON HOLD
    // Memoria dei ticket in lavorazione, sincronizzazione con il form e indicatori di scadenza.
    // ======================================================================

    function parseServiceNowDateTime(value) {
        const match = String(value || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/);
        if (!match) return null;
        const [, year, month, day, hour, minute, second = '0'] = match;
        const date = new Date(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second));
        return Number.isNaN(date.getTime()) ? null : date;
    }
    function formatServiceNowDateTime(date) {
        const pad = value => String(value).padStart(2, '0');
        return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
    }
    function captureTicketSelectionFromSelectors() {
        const companyId = getPanelElement('snow-company-select')?.value;
        const activityId = getPanelElement('snow-activity-select')?.value;
        if (companyId && activityId) savePendingTicketSelection(companyId, activityId);
        else clearPendingTicketSelection();
    }
    function rememberOpenTicket(companyId, activityId) {
        const number = readIncidentNumber();
        if (!number || !companyId || !activityId) return;
        const tickets = loadOpenTickets();
        tickets[number] = {
            number,
            url: snWindow?.location?.href || window.location.href,
            companyId,
            activityId,
            state: readChoice(FIELDS.state),
            holdExpiration: readHoldExpirationValue(),
            resolutionPrepared: false
        };
        saveOpenTickets(tickets);
        refreshOpenTickets();
    }
    function commitPendingTicketAfterSave() {
        const state = readChoice(FIELDS.state);
        if (!['2', '3'].includes(state)) return false;
        const pending = loadPendingTicketSelection();
        const number = readIncidentNumber();
        if (!pending || !number) return false;
        rememberOpenTicket(pending.companyId, pending.activityId);
        clearPendingTicketSelection();
        return true;
    }
    function forgetCurrentTicket() {
        const number = readIncidentNumber();
        if (!number) return;
        const tickets = loadOpenTickets();
        if (!tickets[number]) return;
        delete tickets[number];
        saveOpenTickets(tickets);
        refreshOpenTickets();
    }
    function restoreKnownTicketSelection() {
        const number = readIncidentNumber();
        const ticket = number ? loadOpenTickets()[number] : null;
        if (!ticket) return false;
        const data = loadData();
        if (data.companies[ticket.companyId]) getPanelElement('snow-company-select').value = ticket.companyId;
        if (data.activities[ticket.activityId]) getPanelElement('snow-activity-select').value = ticket.activityId;
        refreshCustomSelect('snow-company-select');
        refreshCustomSelect('snow-activity-select');
        refreshFavoriteButton();
        regenerateResponsePreview();
        return true;
    }
    function syncTrackedTicketState() {
        const number = readIncidentNumber();
        const state = readChoice(FIELDS.state);
        if (!number || !['2', '3'].includes(state)) return;

        const tickets = loadOpenTickets();
        const ticket = tickets[number];
        if (!ticket) return;

        // La pagina ServiceNow è la fonte attendibile: sincronizza anche i ticket
        // già in memoria prima dell'introduzione del monitoraggio della scadenza.
        const pageExpiration = state === '3' ? readHoldExpirationValue() : '';
        const normalizedStoredExpiration = String(ticket.holdExpiration ?? '').trim();
        const stateChanged = String(ticket.state ?? '') !== state;
        const expirationChanged = normalizedStoredExpiration !== pageExpiration;
        const parsedPageExpiration = parseServiceNowDateTime(pageExpiration);
        const acknowledgeExpiredAlert = state === '3' &&
            Boolean(parsedPageExpiration && parsedPageExpiration.getTime() <= Date.now()) &&
            (ticket.expiredAlertAcknowledged !== true || expirationChanged);

        if (!stateChanged && !expirationChanged && !acknowledgeExpiredAlert) return;

        ticket.state = state;
        // Persisti il valore letto dal form. Se il campo è vuoto, non mantenere
        // una vecchia scadenza potenzialmente errata: il tooltip segnalerà che manca.
        ticket.holdExpiration = pageExpiration;
        // La visita del ticket scaduto prende in carico l'avviso: da questo momento
        // l'indicatore resta visibile ma smette di lampeggiare.
        if (state !== '3') ticket.expiredAlertAcknowledged = false;
        else if (acknowledgeExpiredAlert) ticket.expiredAlertAcknowledged = true;
        else if (expirationChanged) ticket.expiredAlertAcknowledged = false;
        tickets[number] = ticket;
        saveOpenTickets(tickets);
        refreshOpenTickets();
    }
    function deleteOpenTicketRecord(number) {
        const tickets = loadOpenTickets();
        if (!tickets[number]) return;
        delete tickets[number];
        saveOpenTickets(tickets);
        refreshOpenTickets();
        showStatus(TEXTS.status.referenceDeleted(number), 'success');
    }
    function getTicketStatusTooltip() {
        let tooltip = document.getElementById('snow-ticket-status-tooltip');
        if (tooltip) return tooltip;
        tooltip = document.createElement('div');
        tooltip.id = 'snow-ticket-status-tooltip';
        Object.assign(tooltip.style, {
            position: 'fixed',
            display: 'none',
            zIndex: '2147483647',
            maxWidth: 'min(320px, calc(100vw - 16px))',
            boxSizing: 'border-box',
            padding: '7px 9px',
            border: `1px solid ${THEME.colors.textMuted}`,
            borderRadius: THEME.radius.base,
            background: THEME.colors.tooltipBg,
            color: THEME.colors.tooltipText,
            boxShadow: '0 4px 14px rgba(0,0,0,.4)',
            font: `${THEME.fontSize.md}/1.4 ${THEME.font.family}`,
            whiteSpace: 'normal',
            overflowWrap: 'anywhere',
            pointerEvents: 'none'
        });
        document.body.appendChild(tooltip);
        return tooltip;
    }
    function positionTicketStatusTooltip(indicator, event) {
        const tooltip = getTicketStatusTooltip();
        const message = indicator.dataset.tooltipText;
        if (!message) {
            tooltip.style.display = 'none';
            return;
        }
        tooltip.textContent = message;
        tooltip.style.display = 'block';
        const rect = indicator.getBoundingClientRect();
        const tipRect = tooltip.getBoundingClientRect();
        const gap = 8;
        let left = event && Number.isFinite(event.clientX) ? event.clientX + 12 : rect.left + rect.width / 2 - tipRect.width / 2;
        let top = rect.top - tipRect.height - gap;
        if (top < 8) top = rect.bottom + gap;
        left = Math.max(8, Math.min(left, window.innerWidth - tipRect.width - 8));
        top = Math.max(8, Math.min(top, window.innerHeight - tipRect.height - 8));
        tooltip.style.left = `${left}px`;
        tooltip.style.top = `${top}px`;
    }
    function hideTicketStatusTooltip() {
        const tooltip = document.getElementById('snow-ticket-status-tooltip');
        if (tooltip) tooltip.style.display = 'none';
    }
    function refreshOpenTickets() {
        const container = getPanelElement('snow-open-tickets-list');
        if (!container) return;
        const tickets = Object.values(loadOpenTickets());
        const title = getPanelElement('snow-assistant-panel')?.querySelector('.snow-open-tickets-title');
        if (title) title.textContent = TEXTS.panel.openTicketsCount(tickets.length);
        const config = loadData();
        container.innerHTML = '';
        if (!tickets.length) {
            container.innerHTML = `<div class="snow-empty">${TEXTS.panel.emptyTickets}</div>`;
            return;
        }
        tickets.forEach(ticket => {
            const company = config.companies[ticket.companyId];
            const activity = config.activities[ticket.activityId];
            const row = document.createElement('div');
            row.className = 'snow-open-ticket-row snow-style-record';
            const info = document.createElement('button');
            info.className = 'snow-open-ticket-info snow-style-button';
            const ticketNumberLine = document.createElement('span');
            ticketNumberLine.className = 'snow-ticket-number-line';
            const ticketNumber = document.createElement('span');
            ticketNumber.className = 'snow-ticket-number';
            ticketNumber.textContent = String(ticket.number ?? '');
            const statusIndicator = document.createElement('span');
            statusIndicator.className = 'snow-ticket-status-indicator';
            ticketNumberLine.append(ticketNumber, statusIndicator);
            const ticketCompany = document.createElement('span');
            ticketCompany.className = 'snow-ticket-company';
            ticketCompany.textContent = String(company?.name || TEXTS.tickets.companyMissing);
            const ticketActivity = document.createElement('span');
            ticketActivity.className = 'snow-ticket-activity';
            ticketActivity.textContent = String(activity?.name || TEXTS.tickets.activityMissing);
            info.append(ticketNumberLine, ticketCompany, ticketActivity);
            if (String(ticket.state ?? '') === '3') {
                const expirationDate = parseServiceNowDateTime(ticket.holdExpiration);
                const holdExpired = Boolean(expirationDate && expirationDate.getTime() <= Date.now());
                if (holdExpired) {
                    statusIndicator.classList.add('expired');
                    if (ticket.expiredAlertAcknowledged) statusIndicator.classList.add('acknowledged');
                    statusIndicator.textContent = '!';
                    statusIndicator.dataset.tooltipText = TEXTS.tickets.holdExpired;
                    statusIndicator.setAttribute('aria-label', TEXTS.tickets.holdExpiredAria);
                } else {
                    statusIndicator.innerHTML = ICONS.pause;
                    const updateHoldTooltip = () => {
                        const currentExpiration = parseServiceNowDateTime(ticket.holdExpiration);
                        if (!currentExpiration) {
                            statusIndicator.dataset.tooltipText = TEXTS.tickets.holdNoExpiration;
                            return;
                        }
                        const remainingMs = currentExpiration.getTime() - Date.now();
                        if (remainingMs <= 0) {
                            statusIndicator.classList.add('expired');
                            statusIndicator.textContent = '!';
                            statusIndicator.dataset.tooltipText = TEXTS.tickets.holdExpired;
                            statusIndicator.setAttribute('aria-label', TEXTS.tickets.holdExpiredAria);
                            statusIndicator.removeEventListener('mouseenter', updateHoldTooltip);
                            statusIndicator.removeEventListener('focus', updateHoldTooltip);
                            return;
                        }
                        const remainingMinutes = Math.ceil(remainingMs / 60000);
                        let remainingLabel;
                        if (remainingMs >= 24 * 60 * 60 * 1000) {
                            const days = Math.floor(remainingMs / (24 * 60 * 60 * 1000));
                            remainingLabel = TEXTS.tickets.days(days);
                        } else if (remainingMs >= 60 * 60 * 1000) {
                            const wholeHours = Math.floor(remainingMs / (60 * 60 * 1000));
                            const minutesIntoHour = Math.floor((remainingMs % (60 * 60 * 1000)) / 60000);
                            const roundedHours = wholeHours + (minutesIntoHour >= 30 ? 1 : 0);
                            remainingLabel = TEXTS.tickets.hours(Math.max(1, roundedHours));
                        } else {
                            remainingLabel = TEXTS.tickets.minutes(remainingMinutes);
                        }
                        statusIndicator.dataset.tooltipText = TEXTS.tickets.holdRemaining(remainingLabel);
                    };
                    updateHoldTooltip();
                    statusIndicator.addEventListener('mouseenter', updateHoldTooltip);
                    statusIndicator.addEventListener('focus', updateHoldTooltip);
                    statusIndicator.setAttribute('aria-label', TEXTS.tickets.holdAria);
                }
            } else {
                statusIndicator.hidden = true;
                statusIndicator.setAttribute('aria-hidden', 'true');
            }
            if (!statusIndicator.hidden) {
                statusIndicator.addEventListener('mouseenter', event => positionTicketStatusTooltip(statusIndicator, event));
                statusIndicator.addEventListener('mousemove', event => positionTicketStatusTooltip(statusIndicator, event));
                statusIndicator.addEventListener('mouseleave', hideTicketStatusTooltip);
                statusIndicator.addEventListener('focus', () => positionTicketStatusTooltip(statusIndicator));
                statusIndicator.addEventListener('blur', hideTicketStatusTooltip);
            }
            info.onclick = () => {
                const safeUrl = getSafeServiceNowTicketUrl(ticket.url);
                if (!safeUrl) {
                    showStatus(TEXTS.status.invalidTicketLink, 'error');
                    return;
                }
                window.open(safeUrl, '_blank', 'noopener,noreferrer');
            };
            const actions = document.createElement('div');
            actions.className = 'snow-open-ticket-actions';
            const safeTicketUrl = getSafeServiceNowTicketUrl(ticket.url);

            const open = document.createElement('button');
            open.className = 'snow-ticket-action snow-style-button snow-style-icon-button';
            open.innerHTML = ICONS.external;
            open.title = TEXTS.tickets.openNewTab;
            open.onclick = e => {
                e.stopPropagation();
                const safeUrl = getSafeServiceNowTicketUrl(ticket.url);
                if (!safeUrl) {
                    showStatus(TEXTS.status.invalidTicketLink, 'error');
                    return;
                }
                window.open(safeUrl, '_blank', 'noopener,noreferrer');
            };

            const copy = document.createElement('button');
            copy.className = 'snow-ticket-action snow-style-button snow-style-icon-button';
            copy.innerHTML = ICONS.copy;
            copy.title = TEXTS.tickets.copyLink;
            copy.onclick = e => {
                e.stopPropagation();
                copyText(`${ticket.number}\n${safeTicketUrl || TEXTS.tickets.invalidLink}`, TEXTS.status.linkCopied);
            };

            const del = document.createElement('button');
            del.className = 'snow-ticket-action delete snow-style-button snow-style-icon-button snow-style-danger';
            del.innerHTML = ICONS.trash;
            del.title = TEXTS.tickets.remove;
            del.onclick = e => {
                e.stopPropagation();
                if (confirm(TEXTS.confirm.deleteReference(ticket.number))) deleteOpenTicketRecord(ticket.number);
            };

            actions.append(open, copy, del);
            row.append(info, actions);
            container.appendChild(row);
        });
    }

    // ======================================================================
    // 8. CHIUSURA E ON HOLD
    // Predisposizione di chiusura e sospensione e modalità del pulsante in base allo stato del ticket.
    // ======================================================================

    function getHoldConfigurationElements() {
        return {
            actions: getPanelElement('snow-resolution-actions'),
            button: getPanelElement('snow-prepare-resolution-button'),
            buttonLabel: getPanelElement('snow-resolution-action-label'),
            check: getPanelElement('snow-resolution-check'),
            pause: getPanelElement('snow-resolution-pause'),
            trigger: getPanelElement('snow-prepare-on-hold-button'),
            panel: getPanelElement('snow-hold-config-panel'),
            closurePanel: getPanelElement('snow-closure-config-panel'),
            reason: getPanelElement('snow-hold-reason'),
            reasonField: getPanelElement('snow-hold-reason')?.closest('.snow-hold-field'),
            duration: getPanelElement('snow-hold-duration'),
            durationField: getPanelElement('snow-hold-duration-field')
        };
    }
    function setReviewNotice(visible) {
        const notice = getPanelElement('snow-review-notice');
        if (!notice) return;
        notice.classList.toggle('is-visible', Boolean(visible));
        notice.setAttribute('aria-hidden', String(!visible));
    }
    function setHoldConfiguration(open, options = {}) {
        const el = getHoldConfigurationElements();
        if (!el.actions || !el.button || !el.panel) return;
        setReviewNotice(open);
        el.actions.classList.toggle('hold-configuring', open);
        el.actions.classList.remove('resolution-configuring');
        if (el.closurePanel) el.closurePanel.hidden = true;
        getPanelElement(APP.panelId)?.classList.toggle('hold-config-active', open);
        el.actions.classList.remove('resolution-ready', 'hold-ready', 'ticket-closed');
        el.panel.hidden = !open;
        if (el.durationField) el.durationField.hidden = !open;
        if (el.buttonLabel) el.buttonLabel.textContent = open ? TEXTS.hold.prepare : TEXTS.closure.prepare;
        if (el.check) el.check.hidden = true;
        if (el.pause) el.pause.hidden = true;
        el.button.classList.remove('ready', 'hold-ready');
        if (el.trigger) el.trigger.setAttribute('aria-expanded', String(open));
        if (!options.fromTour && !open) {
            if (el.reason) el.reason.selectedIndex = -1;
            refreshCustomSelect('snow-hold-reason', true);
            if (el.duration) el.duration.value = '1';
            const durationLabel = getPanelElement('snow-hold-duration-value');
            if (durationLabel) durationLabel.textContent = TEXTS.hold.days(1);
        }
    }
    function calculateHoldExpiration() {
        const el = getHoldConfigurationElements();
        const duration = Number(el.duration?.value || 1);
        if (!Number.isInteger(duration) || duration < 1 || duration > 7) {
            return { error: TEXTS.status.holdDurationInvalid };
        }
        // Acquisisce una sola volta l'ora corrente e calcola la scadenza a partire da quell'istante.
        const currentTime = new Date();
        const expiration = new Date(
            currentTime.getTime() + duration * 24 * 60 * 60 * 1000 + 5 * 60 * 1000
        );
        return { value: formatServiceNowDateTime(expiration) };
    }
    function setPreparedSaveLabel(element, label) {
        if (!element) return;
        element.replaceChildren(document.createTextNode(`${label} `));
        const arrow = document.createElement('span');
        arrow.className = 'snow-prepared-save-arrow';
        arrow.innerHTML = ICONS.preparedArrow;
        element.appendChild(arrow);
        element.appendChild(document.createTextNode(` ${TEXTS.closure.saveSuffix}`));
    }
    function prepareOnHold() {
        const el = getHoldConfigurationElements();
        const preview = String(getPanelElement('snow-response-preview')?.value || '').trim();
        if (!preview) return showStatus(TEXTS.status.previewEmptyHold, 'error');
        if (!el.reason?.value) return showStatus(TEXTS.status.holdReasonMissing, 'error');
        const expiration = calculateHoldExpiration();
        if (expiration.error) return showStatus(expiration.error, 'error');

        writeChoice(FIELDS.state, '3');
        writeChoice(FIELDS.holdReason, el.reason.value);
        writeHoldExpiration(expiration.value);
        writeOnHoldComments(preview);
        setReviewNotice(false);
        el.actions?.classList.remove('hold-configuring', 'resolution-ready');
        el.actions?.classList.add('hold-ready');
        el.panel.hidden = true;
        el.button?.classList.remove('ready');
        el.button?.classList.add('hold-ready');
        setPreparedSaveLabel(el.buttonLabel, TEXTS.hold.prepared);
        if (el.check) el.check.hidden = true;
        if (el.pause) el.pause.hidden = true;
        if (el.trigger) el.trigger.setAttribute('aria-expanded', 'false');
        showStatus(TEXTS.status.holdPrepared(expiration.value), 'success');
    }
    function setClosureConfiguration(open) {
        const el = getHoldConfigurationElements();
        if (!el.actions || !el.closurePanel) return;
        setReviewNotice(open);
        el.actions.classList.toggle('resolution-configuring', open);
        el.actions.classList.remove('hold-configuring');
        if (el.panel) el.panel.hidden = true;
        el.closurePanel.hidden = !open;
        if (el.buttonLabel) el.buttonLabel.textContent = open ? TEXTS.closure.confirm : TEXTS.closure.prepare;
        if (el.check) el.check.hidden = true;
        if (el.pause) el.pause.hidden = true;
        el.button?.classList.remove('ready', 'hold-ready');
    }
    function prepareResolution() {
        const el = getHoldConfigurationElements();
        if (readChoice(FIELDS.state) === '6' && !el.actions?.classList.contains('resolution-ready')) {
            setTicketClosedMode();
            return;
        }
        if (el.actions?.classList.contains('resolution-ready')) {
            return undoPreparedAction('resolution');
        }
        if (el.actions?.classList.contains('hold-ready')) {
            return undoPreparedAction('hold');
        }
        if (el.actions?.classList.contains('hold-configuring')) return prepareOnHold();
        if (!el.actions?.classList.contains('resolution-configuring')) {
            setClosureConfiguration(true);
            return;
        }

        const status = String(getPanelElement('snow-closure-status')?.value || '6');
        const resolutionCode = String(getPanelElement('snow-closure-code')?.value || '');
        if (status !== '6' && status !== '8') return showStatus(TEXTS.status.closureStatusInvalid, 'error');
        if (status === '6' && !resolutionCode) return showStatus(TEXTS.status.closureCodeMissing, 'error');

        const preview = String(getPanelElement('snow-response-preview')?.value || '').trim();
        if (!preview) return showStatus(TEXTS.status.previewEmptyClosure, 'error');
        const profileName = String(loadData().profile.assignedTo.displayValue || '').trim();
        if (!profileName) return showStatus(TEXTS.status.profileMissing, 'error');

        const resolutionText = `${preview}\n${profileName}`;
        setReviewNotice(false);
        writeChoice(FIELDS.state, status);
        writeAdditionalComments(resolutionText);
        writeChoice(FIELDS.resolutionCode, resolutionCode);
        writeTextareaElement(FIELDS.resolutionNotes.elementId, resolutionText);
        el.actions?.classList.add('resolution-ready');
        el.actions?.classList.remove('hold-ready', 'hold-configuring', 'resolution-configuring');
        el.button?.classList.add('ready');
        el.button?.classList.remove('hold-ready');
        setPreparedSaveLabel(el.buttonLabel, status === '8' ? TEXTS.closure.cancelPrepared : TEXTS.closure.prepared);
        if (el.check) el.check.hidden = true;
        if (el.pause) el.pause.hidden = true;
        if (el.panel) el.panel.hidden = true;
        if (el.closurePanel) el.closurePanel.hidden = true;
        if (el.trigger) el.trigger.setAttribute('aria-expanded', 'false');
        const statusLabel = TICKET_STATE_LABELS[status];
        showStatus(TEXTS.status.closurePrepared(statusLabel, resolutionCode), 'success');
    }
    function setTicketClosedMode() {
        const el = getHoldConfigurationElements();
        if (!el.actions || !el.button) return;
        el.actions.classList.remove('resolution-ready', 'hold-ready', 'hold-configuring', 'resolution-configuring');
        el.actions.classList.add('ticket-closed');
        if (el.panel) el.panel.hidden = true;
        if (el.closurePanel) el.closurePanel.hidden = true;
        if (el.durationField) el.durationField.hidden = true;
        el.button.classList.remove('ready', 'hold-ready');
        if (el.buttonLabel) el.buttonLabel.textContent = TEXTS.closure.ticketClosed;
        if (el.check) el.check.hidden = false;
        if (el.pause) el.pause.hidden = true;
        if (el.trigger) el.trigger.setAttribute('aria-expanded', 'false');
    }
    function updateTicketMode() {
        const state = readChoice(FIELDS.state);
        const number = readIncidentNumber();
        const pendingResolutionNumber = sessionStorage.getItem(pendingResolutionSessionKey);
        const gForm = getGForm();

        // Non usare category, agenzia o attività per identificare il ticket:
        // la chiave del monitoraggio è il numero Incident. Dopo il SAVE, il reload
        // conferma che lo stato 6 è stato effettivamente persistito.
        if (pendingResolutionNumber) {
            if (number === pendingResolutionNumber && state === '6') {
                forgetCurrentTicket();
                clearPendingResolutionSave();
            } else if (number === pendingResolutionNumber && state !== '6') {
                // Il salvataggio non ha prodotto uno stato di risoluzione persistito.
                clearPendingResolutionSave();
            }
        }

        // Fallback per ticket già risolti quando il form risulta pulito.
        // Non rimuovere durante la predisposizione: in quel momento il form è modificato.
        if (state === '6' && number && loadOpenTickets()[number] &&
            gForm && typeof gForm.isModified === 'function') {
            try {
                if (gForm.isModified() === false) forgetCurrentTicket();
            } catch (error) {
                console.warn(`[${APP.name}] Impossibile verificare lo stato di salvataggio del ticket:`, error);
            }
        }

        syncTrackedTicketState();
        commitPendingTicketAfterSave();
        const ticketFilled = FILLED_OPEN_STATES.includes(state);
        const ticketClosed = CLOSED_STATES.includes(state);
        const ticketNotNew = ticketFilled || ticketClosed;
        const known = ticketNotNew ? restoreKnownTicketSelection() : false;
        const company = getPanelElement('snow-company-select');
        const activity = getPanelElement('snow-activity-select');
        const apply = getPanelElement('snow-apply-button');
        const favorite = getPanelElement('snow-favorite-button');

        // Agenzia, Attivita e relativi comandi restano sempre disponibili
        // per consultare e gestire le configurazioni, anche sui ticket non modificabili.
        if (company) company.disabled = false;
        if (activity) activity.disabled = false;
        if (favorite) favorite.disabled = false;

        // Compila Ticket è disponibile finché il ticket è aperto; su un ticket già compilato
        // chiede conferma prima di sovrascrivere (vedi applySelectedConfiguration).
        if (apply) apply.disabled = ticketClosed;

        refreshSectionActions();
        refreshClearButtons();

        if (ticketNotNew) {
            const stateLabel = TICKET_STATE_LABELS[state];
            const message = ticketClosed
                ? (known ? TEXTS.status.ticketClosedKnown(stateLabel) : TEXTS.status.ticketClosedUnknown(stateLabel))
                : (known ? TEXTS.status.ticketFilledKnown(stateLabel) : TEXTS.status.ticketFilledUnknown(stateLabel));
            showStatus(message, 'warning');
        }

        // Un ticket già Resolved mostra solo lo stato di chiusura.
        // La classe resolution-ready distingue la predisposizione non ancora salvata.
        const actions = getPanelElement('snow-resolution-actions');
        if (state === '6' && !actions?.classList.contains('resolution-ready')) {
            setTicketClosedMode();
        } else if (state !== '6' && actions?.classList.contains('ticket-closed')) {
            actions.classList.remove('ticket-closed');
            setHoldConfiguration(false);
        }
    }

    // ======================================================================
    // 9. IMPORT / EXPORT
    // Esportazione e importazione della configurazione in JSON.
    // ======================================================================

    function exportConfiguration() {
        closeSettingsMenu();
        const blob = new Blob([JSON.stringify(loadData(), null, 2)], { type: 'application/json;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `snow-assistant-config-${new Date().toISOString().slice(0, 10)}.json`;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
        showStatus(TEXTS.status.exported, 'success');
    }
    function startImport() {
        closeSettingsMenu();
        const input = getPanelElement('snow-import-input');
        input.value = '';
        input.click();
    }
    function importConfiguration(event) {
        const file = event.target.files?.[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = () => {
            try {
                const imported = normalizeData(JSON.parse(reader.result));
                if (!confirm(TEXTS.confirm.importReplace)) return;
                saveData(imported);
                refreshPanel();
                regenerateResponsePreview();
                showStatus(TEXTS.status.imported, 'success');
            } catch (error) {
                console.error(error);
                showStatus(TEXTS.status.importInvalid, 'error');
            }
        };
        reader.readAsText(file);
    }

    // ======================================================================
    // 10. INTERFACCIA: PANNELLO
    // Costruzione del pannello, eventi, messaggi di stato e aggiornamento generale.
    // ======================================================================

    // Elenco a discesa con ricerca: <select> nascosto (contiene i valori) + campo di ricerca + elenco visibile.
    function customSelectHtml({ id, placeholder, label, options = '', wrapClass = '', required = false, clearButtonId = '' }) {
        return `
            <div class="snow-select-wrap${wrapClass ? ` ${wrapClass}` : ''}" data-placeholder="${placeholder}">
              <select id="${id}" tabindex="-1" aria-hidden="true"${required ? ' required' : ''}>${options}</select>
              <input type="text" class="snow-select-trigger" id="${id}-trigger" placeholder="${placeholder}" aria-label="${label}" aria-haspopup="listbox" aria-expanded="false" autocomplete="off" spellcheck="false">
              <button type="button" class="snow-select-trigger-arrow" aria-label="${TEXTS.panel.toggleList}">${ICONS.down}</button>
              <div class="snow-select-dropdown" id="${id}-dropdown" role="listbox" aria-labelledby="${id}-trigger"></div>
              ${clearButtonId ? `<button class="snow-select-clear" id="${clearButtonId}" hidden>${ICONS.remove}</button>` : ''}
            </div>`;
    }
    function optionsHtml(list) {
        return list
            .map(option => `<option value="${option.value}"${option.selected ? ' selected' : ''}${option.disabled ? ' disabled' : ''}>${option.label}</option>`)
            .join('');
    }
    function panelHtml() {
        const T = TEXTS;
        return `
<div class="snow-header">
  <button class="snow-header-button" id="snow-back-button" title="${T.panel.back}" aria-label="${T.panel.back}">${ICONS.left}</button>
  <div class="snow-header-info">
    <div class="snow-title">${T.panel.title}</div>
    <div class="snow-profile" id="snow-profile-label">${T.panel.profileMissingInitial}</div>
  </div>
  <div class="snow-header-actions">
    <div class="snow-settings-wrapper">
      <button class="snow-header-button" id="snow-advanced-button">${ICONS.gear}</button>
      <div class="snow-settings-menu" id="snow-advanced-popover">
        <div class="snow-settings-section">
          <div class="snow-settings-title">${T.settings.favoritesTitle}</div>
          <div class="snow-settings-label">${T.settings.groupBy}</div>
          <div class="snow-settings-segment" id="snow-favorites-group-controls">
            <button id="snow-favorites-group-none">${T.settings.groupNone}</button>
            <button id="snow-favorites-group-company">${T.settings.groupCompany}</button>
            <button id="snow-favorites-group-activity">${T.settings.groupActivity}</button>
          </div>
        </div>
        <div class="snow-settings-section">
          <div class="snow-settings-title">${T.settings.savedTitle}</div>
          <button class="snow-settings-tour snow-style-list-item snow-style-button" id="snow-activity-summary-button">${ICONS.listChecks}<span>${T.settings.savedOpen}</span></button>
        </div>
        <div class="snow-settings-section">
          <div class="snow-settings-title">${T.settings.configTitle}</div>
          <div class="snow-settings-actions">
            <button id="snow-import-button">${ICONS.upload}<span>${T.settings.importJson}</span></button>
            <button id="snow-export-button">${ICONS.download}<span>${T.settings.exportJson}</span></button>
          </div>
        </div>
        <div class="snow-settings-section">
          <button class="snow-settings-tour snow-style-list-item snow-style-button" id="snow-start-tour-button">${ICONS.help}<span>${T.settings.tour}</span></button>
        </div>
      </div>
    </div>
    <button class="snow-header-button snow-profile-button" id="snow-profile-button">${ICONS.user}</button>
  </div>
</div>

<div class="snow-body">
  <!-- Preferiti -->
  <div class="snow-dropdown-section" id="snow-favorites-dropdown">
    <button class="snow-toggle snow-style-section-toggle snow-style-button" id="snow-favorites-toggle">
      <span class="snow-toggle-label">${ICONS.star}<span>${T.panel.favorites}</span></span>
      <span class="snow-toggle-arrow" id="snow-favorites-arrow">${ICONS.down}</span>
    </button>
    <div class="snow-collapsible snow-dropdown-panel" id="snow-favorites-section">
      <div class="snow-favorites" id="snow-favorites-list"></div>
    </div>
  </div>

  <!-- Agenzia -->
  <div class="snow-field">
    <div class="snow-label-row">
      <span class="snow-label">${T.panel.company}</span>
      <span class="snow-field-actions">
        <button class="snow-icon-button" id="snow-save-company-button">${ICONS.save}</button>
        <button class="snow-icon-button" id="snow-edit-company-button" hidden>${ICONS.edit}</button>
        <button class="snow-icon-button delete" id="snow-delete-company-button" hidden>${ICONS.trash}</button>
      </span>
    </div>
    ${customSelectHtml({ id: 'snow-company-select', placeholder: T.panel.companyPlaceholder, label: T.panel.company, clearButtonId: 'snow-clear-company' })}
  </div>

  <!-- Attività -->
  <div class="snow-field">
    <div class="snow-label-row">
      <span class="snow-label">${T.panel.activity}</span>
      <span class="snow-field-actions">
        <button class="snow-icon-button" id="snow-save-activity-button">${ICONS.save}</button>
        <button class="snow-icon-button" id="snow-rename-activity-button" hidden>${ICONS.rename}</button>
        <button class="snow-icon-button" id="snow-edit-activity-button" hidden>${ICONS.edit}</button>
        <button class="snow-icon-button delete" id="snow-delete-activity-button" hidden>${ICONS.trash}</button>
      </span>
    </div>
    ${customSelectHtml({ id: 'snow-activity-select', placeholder: T.panel.activityPlaceholder, label: T.panel.activity, clearButtonId: 'snow-clear-activity' })}
  </div>

  <!-- Compila Ticket e stella dei Preferiti -->
  <div class="snow-main-actions">
    <button class="snow-button snow-primary" id="snow-apply-button">${T.panel.apply}</button>
    <button class="snow-button snow-star" id="snow-favorite-button">${ICONS.star}</button>
  </div>
  <div class="snow-separator"></div>

  <!-- Risposte salvate -->
  <div class="snow-dropdown-section" id="snow-useful-dropdown">
    <button class="snow-toggle snow-style-section-toggle snow-style-button" id="snow-useful-toggle">
      <span class="snow-toggle-label">${ICONS.archive}<span>${T.panel.usefulResponses}</span></span>
      <span class="snow-toggle-arrow" id="snow-useful-arrow">${ICONS.down}</span>
    </button>
    <div class="snow-collapsible snow-dropdown-panel" id="snow-useful-section">
      <div class="snow-useful-list" id="snow-useful-list"></div>
    </div>
  </div>

  <!-- Anteprima risposta -->
  <div class="snow-preview-header">
    <div class="snow-preview-title">${T.panel.preview}</div>
    <div class="snow-preview-actions">
      <button class="snow-button snow-refresh" id="snow-save-combination-response">${ICONS.save}</button>
      <button class="snow-button snow-refresh" id="snow-save-useful-response">${ICONS.archive}</button>
      <button class="snow-button snow-refresh" id="snow-refresh-response-button">${ICONS.refresh}</button>
    </div>
  </div>
  <div class="snow-preview-wrap">
    <textarea class="snow-preview" id="snow-response-preview" spellcheck="true"></textarea>
    <div class="snow-review-notice-slot" aria-live="polite">
      <div class="snow-review-notice" id="snow-review-notice" role="status" aria-hidden="true">${ICONS.warning}<span>${T.panel.reviewNotice}</span></div>
    </div>
  </div>

  <!-- Copia -->
  <div class="snow-copy-bar">
    <button class="snow-copy-action snow-copy-primary" id="snow-copy-response-button">${ICONS.copy}<span>${T.panel.copyResponse}</span></button>
    <button class="snow-copy-action snow-copy-secondary" id="snow-copy-incident-button">${ICONS.copy}<span>${T.panel.copyIncident}</span></button>
  </div>

  <!-- Chiusura e On Hold -->
  <div class="snow-resolution-wrapper" id="snow-resolution-wrapper">
    <div class="snow-resolution-actions" id="snow-resolution-actions">
      <button class="snow-resolution-action" id="snow-prepare-resolution-button">
        <span id="snow-resolution-action-label">${T.closure.prepare}</span>
        <span class="snow-resolution-check" id="snow-resolution-check" hidden>${ICONS.check}</span>
        <span class="snow-resolution-pause" id="snow-resolution-pause" hidden>${ICONS.pause}</span>
      </button>
      <button class="snow-button snow-on-hold" id="snow-prepare-on-hold-button" title="${T.hold.configure}" aria-label="${T.hold.configure}" aria-expanded="false">${ICONS.pause}</button>
    </div>

    <div class="snow-closure-config-panel snow-hold-config-panel" id="snow-closure-config-panel" hidden>
      <div class="snow-hold-field">
        <label for="snow-closure-status-trigger">${T.closure.statusLabel}</label>
        ${customSelectHtml({ id: 'snow-closure-status', placeholder: T.closure.statusPlaceholder, label: T.closure.statusLabel, options: optionsHtml(CLOSURE_STATUSES), wrapClass: 'snow-closure-select-wrap' })}
      </div>
      <div class="snow-hold-field">
        <label for="snow-closure-code-trigger">${T.closure.codeLabel}</label>
        ${customSelectHtml({ id: 'snow-closure-code', placeholder: T.closure.codePlaceholder, label: T.closure.codeLabel, options: optionsHtml(RESOLUTION_CODES), wrapClass: 'snow-closure-select-wrap' })}
      </div>
    </div>

    <div class="snow-hold-config-panel" id="snow-hold-config-panel" hidden>
      <div class="snow-hold-field">
        <label for="snow-hold-reason-trigger">${T.hold.reasonLabel}</label>
        ${customSelectHtml({ id: 'snow-hold-reason', placeholder: T.hold.reasonPlaceholder, label: T.hold.reasonLabel, options: optionsHtml(HOLD_REASONS), wrapClass: 'snow-hold-reason-wrap', required: true })}
      </div>
      <div class="snow-hold-field" id="snow-hold-duration-field" hidden>
        <div class="snow-hold-duration-heading">
          <label for="snow-hold-duration">${T.hold.durationLabel}</label>
          <output id="snow-hold-duration-value" for="snow-hold-duration">${T.hold.days(1)}</output>
        </div>
        <input type="range" id="snow-hold-duration" min="1" max="7" step="1" value="1" aria-label="${T.hold.durationAria}">
        <div class="snow-hold-range-labels" aria-hidden="true">${[1, 3, 5, 7].map(days => `<span>${T.hold.days(days)}</span>`).join('')}</div>
      </div>
      <div class="snow-hold-config-actions">
        <button type="button" class="snow-hold-cancel" id="snow-hold-cancel">${T.hold.cancel}</button>
      </div>
    </div>
  </div>

  <!-- Ticket aperti -->
  <div class="snow-open-tickets-title">${T.panel.openTickets}</div>
  <div class="snow-open-tickets-list" id="snow-open-tickets-list"></div>

  <!-- Finestra Configurazioni salvate -->
  <div class="snow-activity-summary-overlay" id="snow-activity-summary-overlay" hidden role="dialog" aria-modal="true" aria-labelledby="snow-activity-summary-title">
    <div class="snow-activity-summary-panel">
      <div class="snow-activity-summary-header">
        <div>
          <div class="snow-activity-summary-title" id="snow-activity-summary-title">${T.summary.title}</div>
          <div class="snow-activity-summary-subtitle" id="snow-activity-summary-count"></div>
        </div>
        <button type="button" class="snow-activity-summary-close" id="snow-activity-summary-close" aria-label="${T.summary.close}">×</button>
      </div>
      <div class="snow-activity-summary-tabs" role="tablist" aria-label="${T.summary.tabsLabel}">
        <button type="button" class="snow-activity-summary-tab active" id="snow-config-tab-company" role="tab" aria-selected="true" aria-controls="snow-activity-summary-list">${T.summary.companyTab}</button>
        <button type="button" class="snow-activity-summary-tab" id="snow-config-tab-activity" role="tab" aria-selected="false" aria-controls="snow-activity-summary-list">${T.summary.activityTab}</button>
      </div>
      <div class="snow-activity-summary-list" id="snow-activity-summary-list" role="tabpanel" aria-labelledby="snow-config-tab-company"></div>
    </div>
  </div>

  <input type="file" class="snow-file-input" id="snow-import-input" accept=".json,application/json" hidden>
</div>

<div class="snow-status-area"><div class="snow-status" id="snow-status"></div></div>`;
    }
    function buildPanel(container) {
        if (document.getElementById(APP.panelId)) return;
        addStyles();
        const panel = document.createElement('div');
        panel.id = APP.panelId;
        panel.innerHTML = panelHtml();
        container.appendChild(panel);
        bindPanelEvents();
        refreshPanel();
        bindServiceNowEvents();
        regenerateResponsePreview();
    }
    function getPanelElement(id) {
        return document.getElementById(id);
    }
    function bindPanelEvents() {
        const on = (id, event, fn) => getPanelElement(id).addEventListener(event, fn);
        const holdReasonSelect = getPanelElement('snow-hold-reason');
        if (holdReasonSelect) holdReasonSelect.selectedIndex = -1;
        refreshCustomSelect('snow-hold-reason', true);
        on('snow-advanced-button', 'click', event => { event.stopPropagation(); getPanelElement('snow-advanced-popover').classList.toggle('open'); });
        on('snow-activity-summary-button', 'click', event => { event.stopPropagation(); openSavedConfigurations(); });
        on('snow-config-tab-company', 'click', () => selectSavedConfigurationTab('company'));
        on('snow-config-tab-activity', 'click', () => selectSavedConfigurationTab('activity'));
        on('snow-activity-summary-close', 'click', closeActivitySummary);
        on('snow-activity-summary-overlay', 'click', event => { if (event.target === event.currentTarget) closeActivitySummary(); });
        on('snow-profile-button', 'click', configureProfile);
        on('snow-back-button', 'click', () => window.TestPluginAPI?.showMainView?.());
        on('snow-favorites-toggle', 'click', () => {
            const section = getPanelElement('snow-favorites-section');
            const open = !section.classList.contains('open');
            section.classList.toggle('open', open);
            getPanelElement('snow-favorites-arrow').innerHTML = open ? ICONS.up : ICONS.down;
        });
        on('snow-useful-toggle', 'click', () => {
            refreshUsefulResponses();
            const section = getPanelElement('snow-useful-section');
            const open = !section.classList.contains('open');
            section.classList.toggle('open', open);
            getPanelElement('snow-useful-arrow').innerHTML = open ? ICONS.up : ICONS.down;
        });
        ['snow-company-select', 'snow-activity-select'].forEach(selectId => on(selectId, 'change', () => {
            refreshCustomSelect(selectId);
            captureTicketSelectionFromSelectors();
            refreshSectionActions();
            refreshClearButtons();
            refreshFavoriteButton();
            refreshFavorites();
            regenerateResponsePreview();
        }));
        ['snow-company-select', 'snow-activity-select', 'snow-hold-reason', 'snow-closure-status', 'snow-closure-code'].forEach(selectId => {
            const trigger = getPanelElement(`${selectId}-trigger`);
            const arrow = trigger?.closest('.snow-select-wrap')?.querySelector('.snow-select-trigger-arrow');
            arrow?.addEventListener('click', event => {
                event.preventDefault();
                event.stopPropagation();
                const dropdown = getPanelElement(`${selectId}-dropdown`);
                if (dropdown?.classList.contains('open')) {
                    dropdown.classList.remove('open');
                    trigger?.setAttribute('aria-expanded', 'false');
                    trigger.dataset.searchActive = 'false';
                    const select = getPanelElement(selectId);
                    const selectedOption = select?.selectedIndex >= 0 ? select.options[select.selectedIndex] : null;
                    trigger.value = selectedOption?.textContent || '';
                    trigger.blur();
                } else {
                    openCustomSelect(selectId);
                    trigger?.focus({ preventScroll: true });
                }
            });
            trigger.addEventListener('focus', () => {
                if (trigger.dataset.preventOpenOnce === 'true') {
                    trigger.dataset.preventOpenOnce = 'false';
                    return;
                }
                openCustomSelect(selectId);
            });
            trigger.addEventListener('input', () => {
                const dropdown = getPanelElement(`${selectId}-dropdown`);
                if (!dropdown?.classList.contains('open')) openCustomSelect(selectId);
                trigger.dataset.searchActive = 'true';
                filterCustomSelectOptions(selectId);
            });
            trigger.addEventListener('keydown', event => {
                if (event.key === 'ArrowDown') {
                    event.preventDefault();
                    openCustomSelect(selectId);
                    dropdownFocusFirst(selectId);
                } else if (event.key === 'Enter') {
                    const firstVisible = getPanelElement(`${selectId}-dropdown`)?.querySelector('.snow-select-option:not([hidden]):not(:disabled)');
                    if (firstVisible && getPanelElement(`${selectId}-dropdown`)?.classList.contains('open')) {
                        event.preventDefault();
                        firstVisible.click();
                    }
                }
            });
        });
        on('snow-clear-company', 'click', () => clearSelect('snow-company-select'));
        on('snow-clear-activity', 'click', () => clearSelect('snow-activity-select'));
        on('snow-save-company-button', 'click', saveNewCompany);
        on('snow-edit-company-button', 'click', overwriteSelectedCompany);
        on('snow-delete-company-button', 'click', deleteSelectedCompany);
        on('snow-save-activity-button', 'click', saveNewActivity);
        on('snow-rename-activity-button', 'click', renameSelectedActivity);
        on('snow-edit-activity-button', 'click', overwriteSelectedActivity);
        on('snow-delete-activity-button', 'click', deleteSelectedActivity);
        on('snow-apply-button', 'click', applySelectedConfiguration);
        on('snow-favorite-button', 'click', toggleFavorite);
        on('snow-save-combination-response', 'click', saveCombinationResponse);
        on('snow-save-useful-response', 'click', saveUsefulResponse);
        on('snow-refresh-response-button', 'click', regenerateResponsePreview);
        on('snow-copy-response-button', 'click', copyBaseResponse);
        on('snow-copy-incident-button', 'click', copyIncidentWithSeverity);
        on('snow-prepare-resolution-button', 'click', prepareResolution);
        on('snow-prepare-on-hold-button', 'click', () => {
            const actions = getPanelElement('snow-resolution-actions');
            if (readChoice(FIELDS.state) === '6' && !actions?.classList.contains('resolution-ready')) {
                setTicketClosedMode();
                return;
            }
            const isOpen = actions?.classList.contains('hold-configuring');
            if (isOpen) setHoldConfiguration(false);
            else {
                setHoldConfiguration(true);
                getPanelElement('snow-hold-reason-trigger')?.focus();
            }
        });
        on('snow-closure-status', 'change', () => refreshCustomSelect('snow-closure-status'));
        on('snow-closure-code', 'change', () => refreshCustomSelect('snow-closure-code'));
        on('snow-hold-reason', 'change', event => {
            refreshCustomSelect('snow-hold-reason');
            const durationField = getPanelElement('snow-hold-duration-field');
            if (durationField) durationField.hidden = false;
        });
        on('snow-hold-duration', 'input', event => {
            const value = Number(event.target.value);
            const output = getPanelElement('snow-hold-duration-value');
            if (output) output.textContent = TEXTS.hold.days(value);
            const progress = ((value - 1) / 6) * 100;
            event.target.style.setProperty('--snow-hold-progress', `${progress}%`);
            event.target.style.setProperty('--snow-hold-bright-start', `${progress * 0.35}%`);
        });
        const holdDurationSlider = getPanelElement('snow-hold-duration');
        if (holdDurationSlider) {
            const initialDuration = Number(holdDurationSlider.value || 1);
            const progress = ((initialDuration - 1) / 6) * 100;
            holdDurationSlider.style.setProperty('--snow-hold-progress', `${progress}%`);
            holdDurationSlider.style.setProperty('--snow-hold-bright-start', `${progress * 0.35}%`);
        }
        on('snow-hold-cancel', 'click', () => setHoldConfiguration(false));
        on('snow-export-button', 'click', exportConfiguration);
        on('snow-import-button', 'click', startImport);
        on('snow-import-input', 'change', importConfiguration);
        on('snow-favorites-group-none', 'click', () => setFavoritesGrouping('none'));
        on('snow-favorites-group-company', 'click', () => setFavoritesGrouping('company'));
        on('snow-favorites-group-activity', 'click', () => setFavoritesGrouping('activity'));
        on('snow-start-tour-button', 'click', startTour);
        document.addEventListener('click', event => {
            if (
                !event.target.closest(`#${APP.panelId} .snow-settings-wrapper`) &&
                !event.target.closest('#snow-tour-balloon')
            ) {
                closeSettingsMenu();
            }
            if (!event.target.closest(`#${APP.panelId} #snow-favorites-dropdown`)) {
                getPanelElement('snow-favorites-section')?.classList.remove('open');
                const arrow = getPanelElement('snow-favorites-arrow');
                if (arrow) arrow.innerHTML = ICONS.down;
            }
            if (!event.target.closest(`#${APP.panelId} #snow-useful-dropdown`)) {
                getPanelElement('snow-useful-section')?.classList.remove('open');
                const arrow = getPanelElement('snow-useful-arrow');
                if (arrow) arrow.innerHTML = ICONS.down;
            }
            if (!event.target.closest(`#${APP.panelId} #snow-resolution-wrapper`)) {
                const actions = getPanelElement('snow-resolution-actions');
                if (actions?.classList.contains('hold-configuring')) setHoldConfiguration(false);
                if (actions?.classList.contains('resolution-configuring')) setClosureConfiguration(false);
            }
            if (!event.target.closest(`#${APP.panelId} .snow-select-wrap`)) closeCustomSelects();
        }, true);
        document.addEventListener('keydown', event => {
            if (event.key === 'Escape') {
                closeCustomSelects();
                closeActivitySummary();
            }
        });
    }
    function closeSettingsMenu() {
        getPanelElement('snow-advanced-popover')?.classList.remove('open');
    }
    function refreshSectionActions() {
        const companySelected = Boolean(getPanelElement('snow-company-select')?.value);
        const activitySelected = Boolean(getPanelElement('snow-activity-select')?.value);
        getPanelElement('snow-save-company-button').hidden = companySelected;
        getPanelElement('snow-edit-company-button').hidden = !companySelected;
        getPanelElement('snow-delete-company-button').hidden = !companySelected;
        getPanelElement('snow-save-activity-button').hidden = activitySelected;
        getPanelElement('snow-rename-activity-button').hidden = !activitySelected;
        getPanelElement('snow-edit-activity-button').hidden = !activitySelected;
        getPanelElement('snow-delete-activity-button').hidden = !activitySelected;
        refreshClearButtons();
    }
    function refreshClearButtons() {
        const company = getPanelElement('snow-company-select');
        const activity = getPanelElement('snow-activity-select');
        const clearCompany = getPanelElement('snow-clear-company');
        const clearActivity = getPanelElement('snow-clear-activity');
        const companyHasValue = Boolean(company && company.selectedIndex >= 0 && company.value);
        const activityHasValue = Boolean(activity && activity.selectedIndex >= 0 && activity.value);
        if (clearCompany) {
            clearCompany.hidden = !companyHasValue || company.disabled;
            clearCompany.style.display = companyHasValue && !company.disabled ? 'grid' : 'none';
        }
        if (clearActivity) {
            clearActivity.hidden = !activityHasValue || activity.disabled;
            clearActivity.style.display = activityHasValue && !activity.disabled ? 'grid' : 'none';
        }
        company?.closest('.snow-select-wrap')?.classList.toggle('has-value', companyHasValue);
        activity?.closest('.snow-select-wrap')?.classList.toggle('has-value', activityHasValue);
    }
    function refreshPanel(selectedCompanyId, selectedActivityId) {
        const data = loadData();
        const companySelect = getPanelElement('snow-company-select');
        const activitySelect = getPanelElement('snow-activity-select');
        const profile = getPanelElement('snow-profile-label');
        if (!companySelect || !activitySelect || !profile) return;
        const oldCompany = selectedCompanyId || companySelect.value;
        const oldActivity = selectedActivityId || activitySelect.value;
        companySelect.innerHTML = '';
        activitySelect.innerHTML = '';
        Object.entries(data.companies).sort(([, a], [, b]) => a.name.localeCompare(b.name, 'it')).forEach(([id, company]) => {
            const option = document.createElement('option');
            option.value = id;
            option.textContent = company.name;
            companySelect.appendChild(option);
        });
        Object.entries(data.activities).sort(([, a], [, b]) => a.name.localeCompare(b.name, 'it')).forEach(([id, activity]) => {
            const option = document.createElement('option');
            option.value = id;
            option.textContent = activity.name;
            activitySelect.appendChild(option);
        });
        if (data.companies[oldCompany]) companySelect.value = oldCompany;
        else companySelect.selectedIndex = -1;
        if (data.activities[oldActivity]) activitySelect.value = oldActivity;
        else activitySelect.selectedIndex = -1;
        refreshCustomSelect('snow-company-select', true);
        refreshCustomSelect('snow-activity-select', true);
        refreshCustomSelect('snow-closure-status', true);
        refreshCustomSelect('snow-closure-code', true);
        profile.textContent = data.profile.assignedTo.displayValue ? TEXTS.panel.profile(data.profile.assignedTo.displayValue) : TEXTS.panel.profileMissing;
        refreshSectionActions();
        refreshFavoriteButton();
        refreshFavoritesGroupingControls();
        refreshFavorites();
        refreshUsefulResponses();
        refreshOpenTickets();
        updateTicketMode();
    }
    function showStatus(message, type = 'success') {
        const element = getPanelElement('snow-status');
        if (!element) return;
        element.textContent = message;
        element.className = `snow-status ${type}`;
    }

    // ======================================================================
    // 11. INTERFACCIA: CONFIGURAZIONI SALVATE
    // Finestra di riepilogo con le schede Agenzia e Attività.
    // ======================================================================

    function closeActivitySummary() {
        const overlay = getPanelElement('snow-activity-summary-overlay');
        if (overlay) overlay.hidden = true;
        getPanelElement('snow-activity-summary-button')?.classList.remove('active');
    }
    function selectSavedConfigurationTab(tab) {
        const isCompany = tab === 'company';
        const companyTab = getPanelElement('snow-config-tab-company');
        const activityTab = getPanelElement('snow-config-tab-activity');
        const list = getPanelElement('snow-activity-summary-list');
        if (!companyTab || !activityTab || !list) return;
        companyTab.classList.toggle('active', isCompany);
        activityTab.classList.toggle('active', !isCompany);
        companyTab.setAttribute('aria-selected', String(isCompany));
        activityTab.setAttribute('aria-selected', String(!isCompany));
        list.setAttribute('aria-labelledby', isCompany ? 'snow-config-tab-company' : 'snow-config-tab-activity');
        renderSavedConfigurationTab(tab);
    }
    function renderSavedConfigurationTab(tab) {
        const data = loadData();
        const list = getPanelElement('snow-activity-summary-list');
        const count = getPanelElement('snow-activity-summary-count');
        if (!list || !count) return;
        list.innerHTML = '';
        const isCompany = tab === 'company';
        const entries = Object.values(isCompany ? (data.companies || {}) : (data.activities || {})).sort((a, b) =>
            String(a.name || '').localeCompare(String(b.name || ''), 'it')
        );
        const countLabel = isCompany ? TEXTS.summary.companyCount(entries.length) : TEXTS.summary.activityCount(entries.length);
        count.textContent = countLabel;
        if (!entries.length) {
            const empty = document.createElement('div');
            empty.className = 'snow-activity-summary-empty';
            empty.textContent = isCompany ? TEXTS.summary.emptyCompany : TEXTS.summary.emptyActivity;
            list.appendChild(empty);
            return;
        }
        entries.forEach(entry => {
            const card = document.createElement('section');
            card.className = 'snow-activity-summary-card';
            const heading = document.createElement('h3');
            heading.textContent = String(entry.name || (isCompany ? TEXTS.summary.unnamedCompany : TEXTS.summary.unnamedActivity));
            card.appendChild(heading);
            const details = document.createElement('dl');
            const fields = isCompany ? [
                [TEXTS.summary.fieldCompany, entry.company?.displayValue || entry.company || '—'],
                [TEXTS.summary.fieldAssignmentGroup, entry.assignmentGroup?.displayValue || entry.assignmentGroup || '—']
            ] : [
                [TEXTS.summary.fieldCategory, entry.category?.displayValue || entry.category || '—'],
                [TEXTS.summary.fieldImpact, entry.impact ?? '—'],
                [TEXTS.summary.fieldUrgency, entry.urgency ?? '—'],
                [TEXTS.summary.fieldTimeWorked, entry.timeWorked === 0 || entry.timeWorked ? TEXTS.summary.timeWorkedValue(entry.timeWorked) : '—'],
                [TEXTS.summary.fieldPrefix, entry.shortDescriptionPrefix || TEXTS.summary.none]
            ];
            fields.forEach(([fieldLabel, value]) => {
                const row = document.createElement('div');
                row.className = 'snow-activity-summary-row';
                const term = document.createElement('dt');
                term.textContent = fieldLabel;
                const description = document.createElement('dd');
                description.textContent = String(value);
                row.append(term, description);
                details.appendChild(row);
            });
            card.appendChild(details);
            list.appendChild(card);
        });
    }
    function openSavedConfigurations() {
        const overlay = getPanelElement('snow-activity-summary-overlay');
        if (!overlay) return;
        closeSettingsMenu();
        selectSavedConfigurationTab('company');
        overlay.hidden = false;
        getPanelElement('snow-activity-summary-button')?.classList.add('active');
        getPanelElement('snow-activity-summary-close')?.focus();
    }

    // ======================================================================
    // 12. INTERFACCIA: ELENCHI A DISCESA CON RICERCA
    // Select personalizzati costruiti sopra i <select> nascosti.
    // ======================================================================

    function closeCustomSelects(exceptId = '') {
        ['snow-company-select', 'snow-activity-select', 'snow-hold-reason', 'snow-closure-status', 'snow-closure-code'].forEach(selectId => {
            if (selectId === exceptId) return;
            const select = getPanelElement(selectId);
            const trigger = getPanelElement(`${selectId}-trigger`);
            const dropdown = getPanelElement(`${selectId}-dropdown`);
            if (select && trigger) {
                const selectedOption = select.selectedIndex >= 0 ? select.options[select.selectedIndex] : null;
                trigger.value = selectedOption?.textContent || '';
                trigger.dataset.searchActive = 'false';
            }
            trigger?.setAttribute('aria-expanded', 'false');
            dropdown?.classList.remove('open');
        });
    }
    function normalizeSelectSearchText(value) {
        return String(value || '')
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .toLocaleLowerCase('it')
            .trim();
    }
    function refreshCustomSelect(selectId, rebuildOptions = false) {
        const select = getPanelElement(selectId);
        const trigger = getPanelElement(`${selectId}-trigger`);
        const dropdown = getPanelElement(`${selectId}-dropdown`);
        if (!select || !trigger || !dropdown) return;
        const selectedOption = select.selectedIndex >= 0 ? select.options[select.selectedIndex] : null;
        if (!dropdown.classList.contains('open') || trigger.dataset.searchActive !== 'true') {
            trigger.value = selectedOption?.textContent || '';
        }
        trigger.placeholder = select.closest('.snow-select-wrap')?.dataset.placeholder || TEXTS.panel.selectPlaceholder;
        trigger.disabled = select.disabled;
        trigger.setAttribute('aria-disabled', String(select.disabled));
        if (rebuildOptions) {
            dropdown.replaceChildren();
            const optionsContainer = document.createElement('div');
            optionsContainer.className = 'snow-select-options';
            Array.from(select.options).forEach(option => {
                const item = document.createElement('button');
                item.type = 'button';
                item.className = 'snow-select-option snow-style-list-item snow-style-button';
                item.textContent = option.textContent;
                item.title = option.textContent;
                item.setAttribute('role', 'option');
                item.setAttribute('aria-selected', String(option.value === select.value && select.selectedIndex >= 0));
                item.disabled = Boolean(option.disabled);
                item.setAttribute('aria-disabled', String(option.disabled));
                item.dataset.value = option.value;
                item.dataset.searchText = normalizeSelectSearchText(option.textContent);
                item.addEventListener('click', () => {
                    if (select.disabled || option.disabled) return;
                    select.value = option.value;
                    if (select.value !== option.value) select.selectedIndex = -1;
                    trigger.value = option.textContent;
                    trigger.dataset.searchActive = 'false';
                    closeCustomSelects();
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    trigger.dataset.preventOpenOnce = 'true';
                    trigger.focus();
                });
                optionsContainer.appendChild(item);
            });
            const noResults = document.createElement('div');
            noResults.className = 'snow-select-no-results';
            noResults.textContent = TEXTS.panel.noMatches;
            noResults.hidden = true;
            dropdown.append(optionsContainer, noResults);
            filterCustomSelectOptions(selectId);
        } else {
            dropdown.querySelectorAll('.snow-select-option').forEach(item => {
                item.setAttribute('aria-selected', String(item.dataset.value === select.value && select.selectedIndex >= 0));
            });
        }
        trigger.setAttribute('aria-expanded', String(dropdown.classList.contains('open')));
        const wrap = select.closest('.snow-select-wrap');
        wrap?.classList.toggle('has-value', Boolean(selectedOption && select.value));
    }
    function filterCustomSelectOptions(selectId) {
        const trigger = getPanelElement(`${selectId}-trigger`);
        const dropdown = getPanelElement(`${selectId}-dropdown`);
        if (!trigger || !dropdown) return;
        const query = trigger.dataset.searchActive === 'true' ? normalizeSelectSearchText(trigger.value) : '';
        let visibleCount = 0;
        dropdown.querySelectorAll('.snow-select-option').forEach(item => {
            const matches = !query || item.dataset.searchText.includes(query);
            item.hidden = !matches;
            if (matches) visibleCount += 1;
        });
        const noResults = dropdown.querySelector('.snow-select-no-results');
        if (noResults) noResults.hidden = visibleCount > 0;
    }
    function openCustomSelect(selectId) {
        const select = getPanelElement(selectId);
        const trigger = getPanelElement(`${selectId}-trigger`);
        const dropdown = getPanelElement(`${selectId}-dropdown`);
        if (!select || !trigger || !dropdown || select.disabled) return;
        if (dropdown.classList.contains('open')) return;
        closeCustomSelects(selectId);
        refreshCustomSelect(selectId, true);
        trigger.dataset.searchActive = 'false';
        dropdown.classList.add('open');
        trigger.setAttribute('aria-expanded', 'true');
        filterCustomSelectOptions(selectId);
        trigger.select();
    }
    function dropdownFocusFirst(selectId) {
        getPanelElement(`${selectId}-dropdown`)?.querySelector('.snow-select-option:not([hidden]):not(:disabled)')?.focus();
    }
    function clearSelect(selectId) {
        const select = getPanelElement(selectId);
        if (!select || select.disabled) return;
        select.selectedIndex = -1;
        refreshCustomSelect(selectId);
        select.dispatchEvent(new Event('change', { bubbles: true }));
    }

    // ======================================================================
    // 13. TOUR GUIDATO
    // Passi del tour, contenuti dimostrativi e posizionamento del fumetto.
    // ======================================================================

    // Passi del tour, in ordine: elemento indicato e stato del pannello durante il passo.
    // Titoli e testi sono in TEXTS.tour.steps (stesso id).
    const TOUR_STEPS = [
        { id: 'profile', targets: ['snow-profile-button'] },
        { id: 'company', targets: ['snow-company-select-trigger', 'snow-save-company-button', 'snow-edit-company-button', 'snow-delete-company-button'], companyActions: 'all' },
        { id: 'activity', targets: ['snow-activity-select-trigger', 'snow-save-activity-button', 'snow-rename-activity-button', 'snow-edit-activity-button', 'snow-delete-activity-button'], activityActions: 'all' },
        { id: 'apply', targets: ['snow-apply-button'] },
        { id: 'favorites', targets: ['snow-favorite-button', 'snow-favorites-toggle', 'snow-tour-favorite-item'], section: 'favorites', close: 'favorites' },
        { id: 'settings', targets: ['snow-favorites-group-controls', 'snow-activity-summary-button', 'snow-import-button', 'snow-export-button'], targetContainer: 'snow-advanced-popover', settings: true, highlightSettingsButton: true },
        { id: 'preview', targets: ['snow-response-preview', 'snow-save-combination-response', 'snow-save-useful-response', 'snow-refresh-response-button'] },
        { id: 'useful', targets: ['snow-useful-toggle', 'snow-tour-response-copy', 'snow-tour-response-delete', 'snow-tour-response-rename'], section: 'useful', close: 'useful' },
        { id: 'copy', targets: ['snow-copy-response-button', 'snow-copy-incident-button'] },
        { id: 'closure', targets: ['snow-prepare-resolution-button', 'snow-closure-status-trigger', 'snow-closure-code-trigger'], closureSettings: true },
        { id: 'closureConfirm', targets: ['snow-prepare-resolution-button'] },
        { id: 'hold', targets: ['snow-prepare-on-hold-button', 'snow-hold-reason-trigger', 'snow-hold-duration'], holdSettings: true, holdShowDuration: true },
        { id: 'holdConfirm', targets: ['snow-prepare-resolution-button'] },
        { id: 'tickets', targets: ['snow-tour-ticket-record'], section: 'tickets' }
    ];
    let tourIndex = -1;
    let tourSnapshot = null;
    const TOUR_SECTION_ELEMENTS = {
        favorites: ['snow-favorites-section', 'snow-favorites-arrow'],
        useful: ['snow-useful-section', 'snow-useful-arrow']
    };
    const TOUR_ACTION_BUTTON_IDS = [
        'snow-save-company-button', 'snow-edit-company-button', 'snow-delete-company-button',
        'snow-save-activity-button', 'snow-rename-activity-button', 'snow-edit-activity-button', 'snow-delete-activity-button'
    ];
    function setTourSection(name, open) {
        const pair = TOUR_SECTION_ELEMENTS[name];
        if (!pair) return;
        getPanelElement(pair[0])?.classList.toggle('open', open);
        const arrow = getPanelElement(pair[1]);
        if (arrow) arrow.innerHTML = open ? ICONS.up : ICONS.down;
    }
    function setTourSettings(open) {
        getPanelElement('snow-advanced-popover')?.classList.toggle('open', open);
    }
    // Sostituisce temporaneamente gli elenchi con voci dimostrative e salva
    // lo stato da ripristinare a fine tour.
    function createTourTemplates() {
        const favorites = getPanelElement('snow-favorites-list');
        const useful = getPanelElement('snow-useful-list');
        const tickets = getPanelElement('snow-open-tickets-list');
        tourSnapshot = {
            favoritesHtml: favorites?.innerHTML || '',
            usefulHtml: useful?.innerHTML || '',
            ticketsHtml: tickets?.innerHTML || '',
            favoritesOpen: getPanelElement('snow-favorites-section')?.classList.contains('open') || false,
            usefulOpen: getPanelElement('snow-useful-section')?.classList.contains('open') || false,
            resolutionLabel: getPanelElement('snow-resolution-action-label')?.textContent || '',
            hiddenButtons: Object.fromEntries(TOUR_ACTION_BUTTON_IDS.map(id => [id, getPanelElement(id)?.hidden]))
        };
        setTourSection('favorites', false);
        setTourSection('useful', false);
        setTourSettings(false);
        if (favorites) {
            favorites.innerHTML = `<button id="snow-tour-favorite-item" class="snow-favorite-entry snow-style-list-item snow-style-button"><span><strong>${TEXTS.tour.demoCompany}</strong><small>${TEXTS.tour.demoActivity}</small></span></button>`;
        }
        if (useful) {
            const row = document.createElement('div');
            row.className = 'snow-useful-row';
            row.innerHTML = `<button id="snow-tour-response-copy" class="snow-useful-entry snow-style-list-item snow-style-button">${ICONS.copy}<span>${TEXTS.tour.demoResponse}</span></button>`
                + `<div class="snow-useful-actions" style="display:flex">`
                + `<button id="snow-tour-response-delete" class="snow-useful-delete snow-style-button snow-style-icon-button snow-style-danger">${ICONS.trash}</button>`
                + `<button id="snow-tour-response-rename" class="snow-useful-rename snow-style-button snow-style-icon-button">${ICONS.rename}</button>`
                + `</div>`;
            useful.innerHTML = '';
            useful.appendChild(row);
        }
        if (tickets) {
            const row = document.createElement('div');
            row.id = 'snow-tour-ticket-record';
            row.className = 'snow-open-ticket-row snow-style-record selected';
            row.innerHTML = `<button class="snow-open-ticket-info snow-style-button">`
                + `<span class="snow-ticket-number-line"><span class="snow-ticket-number">INC0000000</span><span class="snow-ticket-status-indicator">${ICONS.pause}</span></span>`
                + `<span class="snow-ticket-company">${TEXTS.tour.demoCompany}</span><span class="snow-ticket-activity">${TEXTS.tour.demoActivity}</span>`
                + `</button>`
                + `<div class="snow-open-ticket-actions">`
                + `<button id="snow-tour-ticket-copy" class="snow-ticket-action snow-style-button snow-style-icon-button">${ICONS.copy}</button>`
                + `<button id="snow-tour-ticket-open" class="snow-ticket-action snow-style-button snow-style-icon-button">${ICONS.external}</button>`
                + `<button id="snow-tour-ticket-delete" class="snow-ticket-action delete snow-style-button snow-style-icon-button snow-style-danger">${ICONS.trash}</button>`
                + `</div>`;
            tickets.innerHTML = '';
            tickets.appendChild(row);
        }
    }
    function restoreTourTemplates() {
        if (!tourSnapshot) return;
        const favorites = getPanelElement('snow-favorites-list');
        const useful = getPanelElement('snow-useful-list');
        const tickets = getPanelElement('snow-open-tickets-list');
        if (favorites) favorites.innerHTML = tourSnapshot.favoritesHtml;
        if (useful) useful.innerHTML = tourSnapshot.usefulHtml;
        if (tickets) tickets.innerHTML = tourSnapshot.ticketsHtml;
        const resolutionLabel = getPanelElement('snow-resolution-action-label');
        if (resolutionLabel) resolutionLabel.textContent = tourSnapshot.resolutionLabel;
        Object.entries(tourSnapshot.hiddenButtons).forEach(([id, hidden]) => {
            const button = getPanelElement(id);
            if (button) button.hidden = hidden;
        });
        setTourSection('favorites', tourSnapshot.favoritesOpen);
        setTourSection('useful', tourSnapshot.usefulOpen);
        setTourSettings(false);
        tourSnapshot = null;
        refreshFavorites();
        refreshUsefulResponses();
        refreshOpenTickets();
        refreshSectionActions();
    }
    function stopTour() {
        const currentStep = tourIndex >= 0 ? TOUR_STEPS[tourIndex] : null;
        if (currentStep?.holdSettings || currentStep?.closureSettings) setHoldConfiguration(false, { fromTour: true });
        document
            .querySelectorAll(`#${APP.panelId} .snow-tour-target`)
            .forEach(element => element.classList.remove('snow-tour-target'));
        document
            .querySelectorAll(`#${APP.panelId} .snow-tour-layer`)
            .forEach(element => element.classList.remove('snow-tour-layer'));
        getPanelElement('snow-tour-overlay')?.remove();
        getPanelElement('snow-tour-balloon')?.remove();
        restoreTourTemplates();
        tourIndex = -1;
    }
    // Il fumetto va a sinistra degli elementi evidenziati; se non c'è spazio, a destra.
    function positionTourBalloon(targets, balloon) {
        // Area che comprende tutti gli elementi evidenziati nel passo.
        const rects = targets.map(target => target.getBoundingClientRect()).filter(r => r.width || r.height);
        const rect = rects.length ? {
            left: Math.min(...rects.map(r => r.left)),
            right: Math.max(...rects.map(r => r.right)),
            top: Math.min(...rects.map(r => r.top)),
            bottom: Math.max(...rects.map(r => r.bottom))
        } : targets[0].getBoundingClientRect();
        rect.height = rect.bottom - rect.top;
        // Il fumetto si affianca al pannello (non all'elemento), così non copre i campi vicini.
        const panelLeft = getPanelElement(APP.panelId)?.getBoundingClientRect().left ?? rect.left;
        const width = balloon.offsetWidth;
        const gap = 10;
        let left = Math.min(rect.left, panelLeft) - width - gap;
        let top = rect.top + rect.height / 2 - balloon.offsetHeight / 2;
        if (left < 8) {
            // Nel riquadro dell'add-in non c'è spazio a lato: il fumetto va sotto (o sopra) gli elementi.
            left = Math.max(8, (window.innerWidth - width) / 2);
            const below = rect.bottom + gap;
            top = below + balloon.offsetHeight <= window.innerHeight - 8 ? below : rect.top - gap - balloon.offsetHeight;
        }
        top = Math.max(8, Math.min(window.innerHeight - balloon.offsetHeight - 8, top));
        balloon.style.left = `${left}px`;
        balloon.style.top = `${top}px`;
    }
    function applyTourState(step, prev) {
        if (prev?.close) setTourSection(prev.close, false);
        setTourSection('favorites', step.section === 'favorites');
        setTourSection('useful', step.section === 'useful');
        setTourSettings(Boolean(step.settings));
        setHoldConfiguration(Boolean(step.holdSettings), { fromTour: true });
        if (step.closureSettings) setClosureConfiguration(true);
        if (step.holdShowDuration) {
            const field = getPanelElement('snow-hold-duration-field');
            if (field) field.hidden = false;
        }
    }
    // Pulsanti accanto ad Agenzia e Attività durante il tour: 'new' mostra il "+",
    // 'saved' mostra modifica/elimina (come se una voce fosse selezionata), 'all' li mostra tutti.
    function setTourActionState(step) {
        refreshSectionActions();
        const setHidden = (id, hidden) => {
            const button = getPanelElement(id);
            if (button) button.hidden = hidden;
        };
        if (step.companyActions) {
            const all = step.companyActions === 'all';
            const saved = step.companyActions === 'saved';
            setHidden('snow-save-company-button', saved && !all);
            setHidden('snow-edit-company-button', !saved && !all);
            setHidden('snow-delete-company-button', !saved && !all);
        }
        if (step.activityActions) {
            const all = step.activityActions === 'all';
            const saved = step.activityActions === 'saved';
            setHidden('snow-save-activity-button', saved && !all);
            setHidden('snow-rename-activity-button', !saved && !all);
            setHidden('snow-edit-activity-button', !saved && !all);
            setHidden('snow-delete-activity-button', !saved && !all);
        }
    }
    function showTourStep(index) {
        const prev = tourIndex >= 0 ? TOUR_STEPS[tourIndex] : null;
        document
            .querySelectorAll(`#${APP.panelId} .snow-tour-target`)
            .forEach(element => element.classList.remove('snow-tour-target'));
        document
            .querySelectorAll(`#${APP.panelId} .snow-tour-layer`)
            .forEach(element => element.classList.remove('snow-tour-layer'));

        const step = TOUR_STEPS[index];
        setTourActionState(step);
        applyTourState(step, prev);
        const buttonLabel = TEXTS.tour.steps[step.id].buttonLabel;
        if (buttonLabel) {
            const resolutionLabel = getPanelElement('snow-resolution-action-label');
            if (resolutionLabel) resolutionLabel.textContent = buttonLabel;
        }

        if (step.settings) {
            setTourSettings(true);
        }

        setTimeout(() => {
            const targets = step.targets.map(getPanelElement).filter(Boolean);
            if (!targets.length) {
                return index < TOUR_STEPS.length - 1
                    ? showTourStep(index + 1)
                    : stopTour();
            }

            const layer = step.targetContainer
                ? getPanelElement(step.targetContainer)
                : null;

            tourIndex = index;
            layer?.classList.add('snow-tour-layer');
            targets.forEach(target => target.classList.add('snow-tour-target'));
            // I pulsanti del raggruppamento Preferiti si evidenziano uno per uno.
            getPanelElement('snow-favorites-group-controls')?.classList.contains('snow-tour-target')
                && getPanelElement('snow-favorites-group-controls').querySelectorAll('button')
                    .forEach(button => button.classList.add('snow-tour-target'));
            if (step.highlightSettingsButton) {
                getPanelElement('snow-advanced-button')?.classList.add('snow-tour-target');
            }

            targets[0].scrollIntoView({ block: 'nearest', behavior: 'smooth' });

            const balloon = getPanelElement('snow-tour-balloon');
            // Markup statico: titolo e testo del tour vengono assegnati come testo, non interpretati come HTML.
            balloon.innerHTML = `
                <div class="snow-tour-progress"></div>
                <div class="snow-tour-title"></div>
                <div class="snow-tour-text"></div>
                <div class="snow-tour-actions">
                    <button id="snow-tour-close">${TEXTS.tour.close}</button>
                    <span></span>
                    <button id="snow-tour-prev">${TEXTS.tour.prev}</button>
                    <button id="snow-tour-next">${TEXTS.tour.next}</button>
                </div>
            `;
            balloon.querySelector('.snow-tour-progress').textContent = `${index + 1} / ${TOUR_STEPS.length}`;
            balloon.querySelector('.snow-tour-title').textContent = `${index + 1}. ${TEXTS.tour.steps[step.id].title}`;
            balloon.querySelector('.snow-tour-text').textContent = TEXTS.tour.steps[step.id].text;
            getPanelElement('snow-tour-prev').disabled = index === 0;
            getPanelElement('snow-tour-next').textContent = index === TOUR_STEPS.length - 1 ? TEXTS.tour.finish : TEXTS.tour.next;

            getPanelElement('snow-tour-close').onclick = stopTour;
            getPanelElement('snow-tour-prev').onclick = () => showTourStep(index - 1);
            getPanelElement('snow-tour-next').onclick = () =>
                index === TOUR_STEPS.length - 1
                    ? stopTour()
                    : showTourStep(index + 1);

            requestAnimationFrame(() => {
                requestAnimationFrame(() => positionTourBalloon(targets, balloon));
            });
        }, step.settings ? 60 : 0);
    }
    function startTour() {
        closeSettingsMenu();
        const panel = getPanelElement(APP.panelId);
        stopTour();
        createTourTemplates();
        const overlay = document.createElement('div');
        overlay.id = 'snow-tour-overlay';
        overlay.className = 'snow-tour-overlay';
        const balloon = document.createElement('div');
        balloon.id = 'snow-tour-balloon';
        balloon.className = 'snow-tour-balloon';
        document.body.append(overlay, balloon);
        showTourStep(0);
    }

    // ======================================================================
    // 14. STILI
    // Foglio di stile del pannello, del tour e del tooltip.
    // ======================================================================

    // Trasforma THEME in variabili CSS: THEME.colors.accentDark -> --snow-accent-dark,
    // THEME.fontSize.sm -> --snow-font-sm, THEME.radius.md -> --snow-radius-md, ecc.
    function themeCssVariables() {
        const prefixes = { colors: '', font: 'font-', fontSize: 'font-', fontWeight: 'weight-', radius: 'radius-', shadow: 'shadow-', state: '', layout: '' };
        const kebab = name => name.replace(/([a-z0-9])([A-Z])/g, '$1-$2').toLowerCase();
        return Object.entries(THEME)
            .flatMap(([group, values]) => Object.entries(values)
                .map(([key, value]) => `    --snow-${prefixes[group]}${kebab(key)}: ${value};`))
            .join('\n');
    }
    function addStyles() {
        if (document.getElementById(APP.styleId)) return;
        const style = document.createElement('style');
        style.id = APP.styleId;
        style.textContent = `
  /* ===== Variabili del tema =====
     Generate da THEME (blocco 0 dello script): modificare i valori lì. */
  #${APP.panelId},
  .snow-tour-overlay,
  .snow-tour-balloon {
${themeCssVariables()}
  }

  /* ===== Base e contenitore ===== */
  /* Elementi non cliccabili (pulsanti, voci di elenco, campi): un solo aspetto per tutto il pannello.
     Nessun effetto al passaggio del mouse: ogni regola :hover è scritta come
     :hover:where(:not(:disabled, [aria-disabled="true"])) e vale solo per gli elementi attivi. */
  #${APP.panelId} :disabled,
  #${APP.panelId} [aria-disabled="true"] {
    opacity: var(--snow-disabled-opacity);
    cursor: not-allowed;
  }

  /* Elementi nascosti: unica regola che deve prevalere su tutte
     (gli elenchi, i pannelli e i pulsanti si nascondono con l'attributo hidden). */
  #${APP.panelId} [hidden] {
    display: none !important;
  }

  /* Nell'add-in il pannello occupa tutta la larghezza dell'area AMP. */
  #${APP.panelId} {
    position: relative;
    width: 100%;
    background: var(--snow-white);
    color: var(--snow-text);
    border: 1px solid var(--snow-line);
    border-radius: var(--snow-radius-base);
    box-shadow: 0 1px 2px rgba(0,0,0,.05);
    font-family: var(--snow-font-family);
    font-size: var(--snow-font-base);
    display: flex;
    flex-direction: column;
  }

  #${APP.panelId} * {
    box-sizing: border-box;
  }

  #${APP.panelId} button {
    -webkit-tap-highlight-color: transparent;
  }

  #${APP.panelId} button:focus,
  #${APP.panelId} button:focus-visible {
    outline: none;
    box-shadow: none;
  }

  #${APP.panelId} svg {
    width: 16px;
    height: 16px;
    display: block;
    fill: none;
    stroke: currentColor;
    stroke-width: 1.8;
    stroke-linecap: round;
    stroke-linejoin: round;
    pointer-events: none;
  }

  #${APP.panelId} .snow-body {
    flex: 1;
    padding: 12px 14px 18px;
  }

  #${APP.panelId} select,
  #${APP.panelId} select.snow-tour-target {
    position: absolute;
    width: 1px;
    height: 1px;
    min-width: 0;
    min-height: 0;
    padding: 0;
    border: 0;
    opacity: 0;
    pointer-events: none;
    overflow: hidden;
  }

  #${APP.panelId} select:focus,
  #${APP.panelId} textarea:focus {
    border-color: var(--snow-accent);
    outline: none;
    box-shadow: var(--snow-shadow-focus);
  }

  #${APP.panelId} select:invalid,
  #${APP.panelId} select option[value=""] {
    color: var(--snow-text-disabled);
  }

  #${APP.panelId} .snow-file-input {
    display: none;
  }

  /* ===== Header ===== */
  #${APP.panelId} .snow-header-button svg {
    width: 17px;
    height: 17px;
    margin: auto;
  }

  #${APP.panelId} .snow-header {
    height: var(--snow-header-height);
    min-height: var(--snow-header-height);
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
    padding: 0 10px 0 10px;
    border-radius: calc(var(--snow-radius-base) - 1px) calc(var(--snow-radius-base) - 1px) 0 0;
    background: var(--snow-header-background);
    color: var(--snow-white);
  }

  #${APP.panelId} .snow-header-info {
    min-width: 0;
    overflow: hidden;
    flex: 1;
  }

  #${APP.panelId} .snow-title {
    font-size: var(--snow-font-lg);
    font-weight: var(--snow-weight-bold);
  }

  #${APP.panelId} .snow-profile {
    margin-top: 2px;
    font-size: var(--snow-font-sm);
    opacity: .92;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  #${APP.panelId} .snow-header-actions {
    display: flex;
    align-items: center;
    gap: 4px;
  }

  #${APP.panelId} .snow-header-button {
    width: 28px;
    height: 28px;
    padding: 0;
    border: 1px solid rgba(255,255,255,.45);
    border-radius: var(--snow-radius-base);
    background: rgba(255,255,255,.12);
    color: var(--snow-white);
    cursor: pointer;
    font-size: 16px;
  }

  #${APP.panelId} .snow-header-button:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: rgba(255,255,255,.22);
  }

  #${APP.panelId} .snow-header-button {
    display: grid;
    place-items: center;
    line-height: 0;
  }

  #${APP.panelId} .snow-header-button svg {
    margin: 0;
    align-self: center;
    justify-self: center;
  }

  /* ===== Pulsanti comuni ===== */
  #${APP.panelId} .snow-icon-button svg,
  #${APP.panelId} .snow-refresh svg,
  #${APP.panelId} .snow-star svg {
    margin: auto;
  }

  #${APP.panelId} .snow-icon-button {
    width: 25px;
    height: 24px;
    min-height: 24px;
    padding: 0;
    border: 1px solid var(--snow-line-strong);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    color: var(--snow-text-muted);
    cursor: pointer;
    font-size: var(--snow-font-lg);
    line-height: 20px;
  }

  #${APP.panelId} .snow-icon-button:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-bg);
    border-color: var(--snow-accent-border);
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-icon-button.delete {
    color: var(--snow-text-subtle);
    border-color: var(--snow-line);
  }

  #${APP.panelId} .snow-icon-button.delete:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-danger-bg);
    border-color: var(--snow-danger);
    color: var(--snow-danger);
  }

  #${APP.panelId} .snow-button {
    min-height: var(--snow-control-height);
    padding: 6px 9px;
    border: 1px solid var(--snow-line-strong);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    color: var(--snow-text-muted);
    cursor: pointer;
    font-size: var(--snow-font-md);
    transition: background .12s ease,border-color .12s ease,color .12s ease;
  }

  #${APP.panelId} .snow-button:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-bg);
    border-color: var(--snow-accent-border);
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-primary {
    border: 1px solid var(--snow-accent);
    background: var(--snow-accent);
    color: var(--snow-white);
    font-weight: var(--snow-weight-bold);
    font-size: var(--snow-font-base);
  }

  #${APP.panelId} .snow-primary:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-dark);
    border-color: var(--snow-accent-dark);
    color: var(--snow-white);
  }

  #${APP.panelId} .snow-refresh {
    width: 28px;
    height: 26px;
    min-height: 26px;
    padding: 0;
    font-size: 16px;
  }

  #${APP.panelId} .snow-icon-button.delete svg,
  #${APP.panelId} .snow-ticket-action.delete svg {
    width: 17px;
    height: 17px;
  }

  /* ===== Sezioni a tendina (Preferiti, Risposte salvate) ===== */
  #${APP.panelId} .snow-toggle-label {
    display: inline-flex;
    align-items: center;
    gap: 6px;
  }

  #${APP.panelId} .snow-toggle-label svg {
    width: 18px;
    height: 18px;
  }

  #${APP.panelId} .snow-toggle-arrow {
    display: inline-flex;
    align-items: center;
    justify-content: center;
  }

  #${APP.panelId} .snow-toggle-arrow svg {
    width: 14px;
    height: 14px;
  }

  #${APP.panelId} .snow-toggle {
    width: 100%;
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 7px 8px;
    border: 0;
    border-radius: var(--snow-radius-base);
    background: transparent;
    color: var(--snow-text-muted);
    cursor: pointer;
    font-size: var(--snow-font-md);
    font-weight: var(--snow-weight-bold);
  }

  #${APP.panelId} .snow-toggle:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-bg);
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-collapsible {
    display: none;
  }

  #${APP.panelId} .snow-collapsible.open {
    display: block;
  }

  #${APP.panelId} .snow-dropdown-section {
    position: relative;
    z-index: 1;
  }

  #${APP.panelId} .snow-dropdown-section:has(.snow-dropdown-panel.open) {
    z-index: 12;
  }

  #${APP.panelId} .snow-dropdown-panel {
    display: none;
    position: absolute;
    top: calc(100% + 2px);
    left: 0;
    right: 0;
    z-index: 30;
    max-height: 220px;
    overflow-y: auto;
    overscroll-behavior: contain;
    padding: 5px 7px 7px;
    border: 1px solid var(--snow-line-light);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    box-shadow: var(--snow-shadow-dropdown);
  }

  #${APP.panelId} .snow-dropdown-panel.open {
    display: block;
  }

  #${APP.panelId} .snow-empty {
    color: var(--snow-text-subtle);
    font-size: var(--snow-font-sm);
    font-style: italic;
    padding-bottom: 5px;
  }

  #${APP.panelId} .snow-separator {
    height: 1px;
    margin: 13px -14px;
    background: var(--snow-line-light);
  }

  #${APP.panelId} .snow-style-section-toggle {
    border-radius: 0;
    background: transparent;
    box-shadow: none;
  }

  #${APP.panelId} .snow-style-section-toggle:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-bg);
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-style-section-toggle .snow-toggle-label svg {
    color: var(--snow-text-subtle);
    fill: none;
    stroke: currentColor;
  }

  #${APP.panelId} .snow-style-section-toggle .snow-toggle-arrow svg {
    color: inherit;
    stroke: currentColor;
  }

  #${APP.panelId} .snow-style-section-toggle:hover:where(:not(:disabled, [aria-disabled="true"])) .snow-toggle-label svg,
  #${APP.panelId} .snow-style-section-toggle:hover:where(:not(:disabled, [aria-disabled="true"])) .snow-toggle-arrow svg {
    color: inherit;
  }

  /* ===== Campi Agenzia e Attività ===== */
  #${APP.panelId} .snow-field {
    margin-bottom: 13px;
  }

  #${APP.panelId} .snow-label-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 5px;
  }

  #${APP.panelId} .snow-label {
    color: var(--snow-text);
    font-size: var(--snow-font-md);
    font-weight: var(--snow-weight-bold);
  }

  #${APP.panelId} .snow-field-actions {
    display: flex;
    gap: 4px;
  }

  /* ===== Elenchi a discesa con ricerca ===== */
  #${APP.panelId} .snow-select-trigger {
    display: block;
    width: 100%;
    height: var(--snow-control-height);
    min-height: var(--snow-control-height);
    box-sizing: border-box;
    padding: 4px 62px 4px 8px;
    border: 1px solid var(--snow-line);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    color: var(--snow-text);
    caret-color: transparent;
    text-align: left;
    font: inherit;
    font-size: var(--snow-font-base);
    outline: none;
    transition: border-color .12s ease, box-shadow .12s ease, background .12s ease;
  }

  #${APP.panelId} .snow-select-trigger::placeholder {
    color: var(--snow-text-disabled);
    opacity: .75;
  }

  #${APP.panelId} .snow-select-wrap:has(.snow-select-dropdown.open) .snow-select-trigger {
    caret-color: var(--snow-accent);
  }

  #${APP.panelId} .snow-select-trigger:hover:where(:not(:disabled, [aria-disabled="true"])) {
    border-color: var(--snow-accent-border);
    background: var(--snow-accent-bg);
  }

  #${APP.panelId} .snow-select-trigger:focus {
    border-color: var(--snow-accent);
    box-shadow: var(--snow-shadow-focus);
  }

  #${APP.panelId} button.snow-select-trigger-arrow {
    -webkit-appearance: none;
    appearance: none;
    position: absolute;
    right: 9px;
    top: 50%;
    display: grid;
    place-items: center;
    width: 16px;
    height: 16px;
    min-width: 16px;
    margin: 0;
    padding: 0;
    transform: translateY(-50%);
    border: 0;
    border-radius: 0;
    outline: 0;
    background: transparent;
    box-shadow: none;
    color: var(--snow-text-subtle);
    font: inherit;
    line-height: 1;
    cursor: pointer;
    pointer-events: auto;
  }

  #${APP.panelId} button.snow-select-trigger-arrow:hover:where(:not(:disabled, [aria-disabled="true"])),
  #${APP.panelId} button.snow-select-trigger-arrow:active {
    border: 0;
    background: transparent;
    box-shadow: none;
    color: var(--snow-text-subtle);
  }

  #${APP.panelId} button.snow-select-trigger-arrow:focus-visible {
    outline: 0;
  }

  #${APP.panelId} button.snow-select-trigger-arrow svg {
    width: 12px;
    height: 12px;
    pointer-events: none;
  }

  #${APP.panelId} .snow-select-wrap:has(.snow-select-dropdown.open) {
    z-index: 35;
  }

  #${APP.panelId} .snow-select-dropdown {
    display: none;
    position: absolute;
    top: calc(100% + 2px);
    left: 0;
    right: 0;
    width: auto;
    max-width: 100%;
    max-height: 220px;
    box-sizing: border-box;
    overflow-x: hidden;
    overflow-y: auto;
    overscroll-behavior: contain;
    padding: 5px 7px 7px;
    border: 1px solid var(--snow-line-light);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    box-shadow: var(--snow-shadow-dropdown);
    z-index: 36;
  }

  #${APP.panelId} .snow-select-dropdown.open {
    display: block;
  }

  #${APP.panelId} .snow-select-options {
    display: flex;
    flex-direction: column;
    gap: 1px;
  }

  #${APP.panelId} .snow-select-no-results {
    padding: 8px 6px;
    color: var(--snow-text-subtle);
    font-size: var(--snow-font-sm);
    font-style: italic;
  }

  #${APP.panelId} .snow-select-option {
    display: block;
    width: 100%;
    min-width: 0;
    box-sizing: border-box;
    padding: 7px 6px;
    border: 0;
    border-left: 2px solid transparent;
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    color: var(--snow-text);
    text-align: left;
    font-size: var(--snow-font-md);
    line-height: 1.35;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    cursor: pointer;
  }

  #${APP.panelId} .snow-select-option:hover:where(:not(:disabled, [aria-disabled="true"])),
  #${APP.panelId} .snow-select-option:focus-visible {
    border-left-color: var(--snow-accent-border);
    outline: none;
    background: var(--snow-accent-bg);
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-select-wrap {
    position: relative;
  }

  #${APP.panelId} .snow-select-clear {
    position: absolute;
    z-index: 2;
    right: 24px;
    top: 50%;
    transform: translateY(-50%);
    width: 22px;
    height: 22px;
    padding: 0;
    border: 0;
    background: transparent;
    color: var(--snow-text-subtle);
    display: grid;
    place-items: center;
    cursor: pointer;
  }

  #${APP.panelId} .snow-select-clear:hover:where(:not(:disabled, [aria-disabled="true"])) {
    color: var(--snow-danger);
    background: transparent;
    border: 0;
    border-radius: 0;
    box-shadow: none;
  }

  #${APP.panelId} .snow-select-clear:focus,
  #${APP.panelId} .snow-select-clear:focus-visible {
    outline: none;
    border: 0;
    box-shadow: none;
  }

  #${APP.panelId} .snow-select-clear svg {
    width: 12px;
    height: 12px;
    stroke-width: 2.2;
  }

  /* ===== Azioni principali (Compila Ticket, stella) ===== */
  #${APP.panelId} .snow-star svg {
    width: 18px;
    height: 18px;
  }

  #${APP.panelId} .snow-star.active svg {
    fill: currentColor;
  }

  #${APP.panelId} .snow-main-actions {
    display: grid;
    grid-template-columns: minmax(0, 1fr) 44px;
    gap: 6px;
    margin-top: 8px;
    border: 0;
    border-radius: 0;
    overflow: visible;
    background: transparent;
  }

  #${APP.panelId} .snow-star {
    padding: 0;
    font-size: 20px;
    color: var(--snow-text-subtle);
  }

  #${APP.panelId} .snow-star:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-bg);
    border-color: var(--snow-accent-border);
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-star.active {
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-main-actions .snow-primary {
    height: var(--snow-action-height);
    min-height: var(--snow-action-height);
    border: 1px solid var(--snow-accent);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    color: var(--snow-accent);
    box-shadow: var(--snow-shadow-button);
  }

  #${APP.panelId} .snow-main-actions .snow-primary:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-bg);
    border-color: var(--snow-accent-dark);
    color: var(--snow-accent-dark);
    box-shadow: var(--snow-shadow-button-hover);
    transform: translateY(-1px);
  }

  #${APP.panelId} .snow-main-actions .snow-star {
    height: var(--snow-action-height);
    min-height: var(--snow-action-height);
    border: 1px solid var(--snow-accent);
    border-radius: var(--snow-radius-base);
    background: var(--snow-accent);
    color: var(--snow-white);
    box-shadow: var(--snow-shadow-button);
    transition: background .12s ease,border-color .12s ease,color .12s ease,box-shadow .12s ease,transform .12s ease;
  }

  #${APP.panelId} .snow-main-actions .snow-star:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-dark);
    border-color: var(--snow-accent-dark);
    color: var(--snow-white);
    box-shadow: var(--snow-shadow-button-hover);
    transform: translateY(-1px);
  }

  #${APP.panelId} .snow-main-actions .snow-star.active {
    background: var(--snow-accent);
    border-color: var(--snow-accent);
    color: var(--snow-white);
  }

  #${APP.panelId} .snow-main-actions .snow-star.active svg {
    fill: currentColor;
  }

  #${APP.panelId} .snow-main-actions .snow-star.active:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-dark);
    border-color: var(--snow-accent-dark);
    color: var(--snow-white);
  }

  /* ===== Preferiti ===== */
  #${APP.panelId} .snow-favorite-entry {
    display: flex;
    align-items: center;
    gap: 0;
    width: 100%;
    padding: 7px 4px;
    border: 0;
    border-left: 2px solid transparent;
    border-radius: var(--snow-radius-base);
    background: transparent;
    color: var(--snow-text);
    text-align: left;
    cursor: pointer;
    font-size: var(--snow-font-sm);
  }

  #${APP.panelId} .snow-favorite-entry svg {
    width: 14px;
    height: 14px;
    color: var(--snow-text-subtle);
    flex: 0 0 auto;
  }

  #${APP.panelId} .snow-favorites {
    display: flex;
    flex-direction: column;
    gap: 1px;
    padding: 2px 0 10px;
  }

  #${APP.panelId} .snow-favorite-entry span {
    display: flex;
    flex-direction: column;
    gap: 1px;
    min-width: 0;
  }

  #${APP.panelId} .snow-favorite-entry strong {
    font-size: var(--snow-font-md);
    font-weight: var(--snow-weight-medium);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  #${APP.panelId} .snow-favorite-entry small {
    font-size: var(--snow-font-md);
    color: var(--snow-text-subtle);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  #${APP.panelId} .snow-favorite-group {
    padding: 2px 0;
  }

  #${APP.panelId} .snow-favorite-group-toggle {
    width: 100%;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 7px 4px;
    border: 0;
    border-bottom: 1px solid var(--snow-line-light);
    background: transparent;
    color: var(--snow-accent);
    font-size: var(--snow-font-xs);
    font-weight: var(--snow-weight-bold);
    text-transform: uppercase;
    letter-spacing: .35px;
    cursor: pointer;
    text-align: left;
  }

  #${APP.panelId} .snow-favorite-group-toggle:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: transparent;
  }

  #${APP.panelId} .snow-favorite-group-arrow {
    display: grid;
    place-items: center;
    color: var(--snow-text-subtle);
  }

  #${APP.panelId} .snow-favorite-group-arrow svg {
    width: 12px;
    height: 12px;
  }

  #${APP.panelId} .snow-favorite-group-items {
    display: none;
    padding: 2px 0 5px 8px;
  }

  #${APP.panelId} .snow-favorite-group.open .snow-favorite-group-items {
    display: block;
  }

  #${APP.panelId} .snow-favorite-group-items .snow-favorite-entry {
    padding: 6px 4px;
  }

  #${APP.panelId} .snow-favorite-entry.selected strong,
  #${APP.panelId} .snow-favorite-entry.selected small {
    color: inherit;
    font-weight: inherit;
  }

  #${APP.panelId} #snow-favorites-toggle .snow-toggle-label svg {
    fill: currentColor;
  }

  #${APP.panelId} .snow-style-group-toggle {
    background: transparent;
    box-shadow: none;
  }

  #${APP.panelId} .snow-style-group-toggle:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: transparent;
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-style-group-toggle svg {
    stroke: currentColor;
  }

  /* ===== Anteprima risposta e copia ===== */
  #${APP.panelId} .snow-preview-title {
    color: var(--snow-text-muted);
    font-size: var(--snow-font-sm);
    font-weight: var(--snow-weight-bold);
    text-transform: uppercase;
    letter-spacing: .35px;
  }

  #${APP.panelId} .snow-preview-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 6px;
  }

  #${APP.panelId} .snow-preview-wrap {
    position: relative;
    width: 100%;
  }

  #${APP.panelId} .snow-preview {
    display: block;
    width: 100%;
    min-height: 150px;
    resize: vertical;
    padding: 8px;
    border: 1px solid var(--snow-line-strong);
    border-radius: var(--snow-radius-base);
    font-family: var(--snow-font-family);
    font-size: var(--snow-font-md);
    line-height: 1.45;
  }

  #${APP.panelId} .snow-review-notice-slot {
    position: absolute;
    z-index: 2;
    left: 1px;
    right: 1px;
    bottom: 1px;
    min-height: 0;
    box-sizing: border-box;
    padding: 0;
    display: flex;
    align-items: stretch;
    pointer-events: none;
    border-radius: 0 0 calc(var(--snow-radius-base) - 1px) calc(var(--snow-radius-base) - 1px);
    background: transparent;
  }

  #${APP.panelId} .snow-review-notice {
    box-sizing: border-box;
    display: flex;
    align-items: center;
    gap: 5px;
    width: 100%;
    min-height: 16px;
    padding: 5px 8px 6px;
    border: 0;
    border-top: 1px solid transparent;
    border-radius: 0 0 calc(var(--snow-radius-base) - 1px) calc(var(--snow-radius-base) - 1px);
    background: transparent;
    color: var(--snow-status-warning-text);
    font-size: var(--snow-font-md);
    line-height: 16px;
    font-weight: var(--snow-weight-medium);
    visibility: hidden;
    opacity: 0;
    transition: opacity .14s ease;
    pointer-events: none;
  }

  #${APP.panelId} .snow-review-notice.is-visible {
    visibility: visible;
    opacity: 1;
    border-top-color: var(--snow-status-warning-border);
    background: var(--snow-status-warning-bg);
  }

  #${APP.panelId} .snow-review-notice svg {
    flex: 0 0 14px;
    width: 14px;
    height: 14px;
    stroke: currentColor;
    fill: none;
    stroke-width: 2;
  }

  #${APP.panelId} .snow-copy-bar {
    display: grid;
    grid-template-columns: 1fr 1fr;
    margin-top: 8px;
    border: 1px solid var(--snow-line);
    border-radius: var(--snow-radius-base);
    overflow: hidden;
  }

  #${APP.panelId} .snow-copy-action {
    height: var(--snow-action-height);
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    border: 0;
    cursor: pointer;
    font-size: var(--snow-font-base);
    font-weight: var(--snow-weight-medium);
  }

  #${APP.panelId} .snow-copy-action svg {
    width: 18px;
    height: 18px;
  }

  #${APP.panelId} .snow-copy-primary {
    background: var(--snow-white);
    color: var(--snow-accent);
    border-right: 1px solid var(--snow-line);
  }

  #${APP.panelId} .snow-copy-primary:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-bg);
    color: var(--snow-accent-dark);
  }

  #${APP.panelId} .snow-copy-secondary {
    background: var(--snow-white);
    color: var(--snow-text-muted);
    border-left: 1px solid var(--snow-line);
  }

  #${APP.panelId} .snow-copy-secondary:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-bg);
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-preview-actions {
    display: flex;
    gap: 4px;
  }

  /* ===== Risposte salvate ===== */
  #${APP.panelId} #snow-save-useful-response {
    color: var(--snow-text-subtle);
  }

  #${APP.panelId} #snow-save-useful-response:hover:where(:not(:disabled, [aria-disabled="true"])) {
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-useful-list {
    display: flex;
    flex-direction: column;
    gap: 4px;
    padding: 5px 0 9px;
  }

  #${APP.panelId} .snow-useful-row {
    display: flex;
    align-items: center;
    gap: 5px;
  }

  #${APP.panelId} .snow-useful-entry {
    flex: 1;
    display: flex;
    align-items: center;
    gap: 7px;
    padding: 7px 8px;
    border: 1px solid transparent;
    border-radius: var(--snow-radius-base);
    background: var(--snow-bg-soft);
    color: var(--snow-text);
    text-align: left;
    cursor: pointer;
    font-family: inherit;
    font-size: var(--snow-font-base);
    font-weight: var(--snow-weight-normal);
    line-height: 1.35;
  }

  #${APP.panelId} .snow-useful-entry:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-bg);
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-useful-entry svg {
    width: 14px;
    height: 14px;
    flex: 0 0 auto;
  }

  #${APP.panelId} .snow-useful-rename,
  #${APP.panelId} .snow-useful-delete {
    width: 28px;
    height: 28px;
    min-width: 28px;
    flex: 0 0 28px;
    padding: 0;
    border: 1px solid var(--snow-line);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    display: grid;
    place-items: center;
    cursor: pointer;
  }

  #${APP.panelId} .snow-useful-actions {
    display: none;
    align-items: center;
    gap: 4px;
    margin-left: 8px;
  }

  #${APP.panelId} .snow-useful-row:hover:where(:not(:disabled, [aria-disabled="true"])) .snow-useful-actions,
  #${APP.panelId} .snow-useful-row:focus-within .snow-useful-actions {
    display: flex;
  }

  #${APP.panelId} .snow-useful-rename {
    border: 1px solid var(--snow-line-strong);
    color: var(--snow-text-muted);
  }

  #${APP.panelId} .snow-useful-rename:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-bg);
    border-color: var(--snow-accent-border);
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-useful-rename svg,
  #${APP.panelId} .snow-useful-delete svg {
    width: 16px;
    height: 16px;
    stroke-width: 1.8;
  }

  #${APP.panelId} .snow-useful-entry>svg {
    display: block;
    flex: 0 0 auto;
    align-self: center;
    margin: 0;
  }

  /* ===== Chiusura e On Hold ===== */
  #${APP.panelId} .snow-resolution-actions .snow-on-hold {
    display: grid;
    place-items: center;
    height: var(--snow-resolution-height);
    min-height: var(--snow-resolution-height);
    padding: 0;
    border: 1px solid var(--snow-accent);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    color: var(--snow-accent);
    box-shadow: var(--snow-shadow-button);
    transition: background .18s ease,border-color .18s ease,color .18s ease,box-shadow .18s ease,transform .18s ease;
  }

  #${APP.panelId} .snow-resolution-actions .snow-on-hold:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-bg);
    border-color: var(--snow-accent-dark);
    color: var(--snow-accent-dark);
    box-shadow: var(--snow-shadow-button-hover);
    transform: translateY(-1px);
  }

  #${APP.panelId} .snow-resolution-actions .snow-on-hold svg {
    width: 19px;
    height: 19px;
    stroke: currentColor;
  }

  #${APP.panelId} .snow-resolution-actions .snow-on-hold svg path {
    stroke-width: 4.0;
    stroke-linecap: round;
  }

  /* z-index 3: i riquadri di chiusura e On Hold stanno sopra i pulsanti dei Ticket aperti (z-index 2). */
  #${APP.panelId} .snow-resolution-wrapper {
    position: relative;
    z-index: 3;
    margin-top: 9px;
  }

  #${APP.panelId} .snow-resolution-actions {
    display: grid;
    grid-template-columns: minmax(0, 1fr) 44px;
    column-gap: 6px;
    position: relative;
    transition: grid-template-columns .24s ease;
  }

  #${APP.panelId} .snow-resolution-action {
    width: 100%;
    height: var(--snow-resolution-height);
    margin-top: 0;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    border: 1px solid var(--snow-accent);
    border-radius: var(--snow-radius-base);
    background: var(--snow-accent);
    color: var(--snow-white);
    cursor: pointer;
    font-size: var(--snow-font-base);
    font-weight: var(--snow-weight-bold);
    box-shadow: var(--snow-shadow-button);
    transition: background .18s ease,border-color .18s ease,color .18s ease,width .24s ease,box-shadow .18s ease,transform .18s ease;
  }

  #${APP.panelId} .snow-resolution-action svg {
    width: 18px;
    height: 18px;
  }

  #${APP.panelId} .snow-prepared-save-arrow {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    flex: 0 0 auto;
    width: 15px;
    height: 18px;
    vertical-align: middle;
  }

  #${APP.panelId} .snow-prepared-save-arrow svg {
    width: 15px;
    height: 15px;
    fill: none;
    stroke: currentColor;
    stroke-width: 3.5;
    stroke-linecap: round;
    stroke-linejoin: round;
  }

  #${APP.panelId} .snow-resolution-action:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-dark);
    border-color: var(--snow-accent-dark);
    color: var(--snow-white);
    box-shadow: var(--snow-shadow-button-hover);
    transform: translateY(-1px);
  }

  #${APP.panelId} .snow-resolution-action.ready {
    background: var(--snow-success);
    border-color: var(--snow-success);
    color: var(--snow-white);
  }

  #${APP.panelId} .snow-resolution-action.ready:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-success-dark);
    border-color: var(--snow-success-dark);
    color: var(--snow-white);
  }

  #${APP.panelId} .snow-resolution-check {
    margin-left: auto;
  }

  #${APP.panelId} .snow-resolution-check svg {
    width: 20px;
    height: 20px;
    stroke-width: 3;
  }

  #${APP.panelId} .snow-resolution-actions.resolution-ready,
  #${APP.panelId} .snow-resolution-actions.hold-ready,
  #${APP.panelId} .snow-resolution-actions.ticket-closed {
    grid-template-columns: minmax(0, 1fr);
  }

  #${APP.panelId} .snow-resolution-actions.resolution-ready .snow-resolution-action,
  #${APP.panelId} .snow-resolution-actions.hold-ready .snow-resolution-action,
  #${APP.panelId} .snow-resolution-actions.ticket-closed .snow-resolution-action {
    grid-column: 1 / -1;
    border-radius: var(--snow-radius-base);
  }

  #${APP.panelId} .snow-resolution-actions.resolution-ready .snow-on-hold,
  #${APP.panelId} .snow-resolution-actions.hold-ready .snow-on-hold,
  #${APP.panelId} .snow-resolution-actions.ticket-closed .snow-on-hold {
    display: none;
  }

  #${APP.panelId} .snow-resolution-actions.hold-configuring {
    column-gap: 6px;
  }

  #${APP.panelId} .snow-resolution-actions.hold-configuring .snow-resolution-action,
  #${APP.panelId} .snow-resolution-actions.hold-configuring .snow-on-hold {
    border-radius: var(--snow-radius-base);
  }

  #${APP.panelId} .snow-resolution-actions.hold-configuring .snow-on-hold {
    background: var(--snow-white);
    border: 1px solid var(--snow-accent);
    color: var(--snow-accent);
    box-shadow: var(--snow-shadow-button);
  }

  #${APP.panelId} .snow-resolution-actions.hold-configuring .snow-on-hold:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-bg);
    border-color: var(--snow-accent-dark);
    color: var(--snow-accent-dark);
    box-shadow: var(--snow-shadow-button-hover);
    transform: translateY(-1px);
  }

  #${APP.panelId} .snow-resolution-action.hold-ready,
  #${APP.panelId} .snow-resolution-action.hold-ready:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-success);
    border-color: var(--snow-success);
    color: var(--snow-white);
    animation: none;
  }

  #${APP.panelId} .snow-resolution-action.hold-ready:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-success-dark);
    border-color: var(--snow-success-dark);
  }

  #${APP.panelId} .snow-resolution-actions.ticket-closed .snow-resolution-action,
  #${APP.panelId} .snow-resolution-actions.ticket-closed .snow-resolution-action:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent);
    border-color: var(--snow-accent);
    color: var(--snow-white);
    animation: none;
    cursor: default;
    transform: none;
  }

  #${APP.panelId} .snow-resolution-pause svg,
  #${APP.panelId} .snow-on-hold svg {
    width: 18px;
    height: 18px;
    stroke: currentColor;
    stroke-width: 0;
    stroke-linecap: butt;
    stroke-linejoin: miter;
    fill: currentColor;
  }

  #${APP.panelId} .snow-hold-config-panel,
  #${APP.panelId} .snow-closure-config-panel {
    display: grid;
    grid-template-columns: minmax(0, 1fr);
    gap: 11px;
    position: absolute;
    top: calc(100% + 5px);
    right: 0;
    width: 100%;
    box-sizing: border-box;
    margin: 0;
    padding: 13px 12px 11px;
    border: 1px solid var(--snow-accent);
    border-radius: var(--snow-radius-base);
    background: var(--snow-accent);
    color: var(--snow-white);
    box-shadow: 0 8px 22px rgba(0,0,0,.20);
    z-index: 80;
    overflow: visible;
    animation: snow-hold-panel-in .18s cubic-bezier(.2,.75,.25,1) both;
  }

  #${APP.panelId} .snow-hold-config-panel:not(.snow-closure-config-panel) {
    background: var(--snow-white);
    color: var(--snow-text);
  }

  #${APP.panelId} .snow-hold-config-panel:not(.snow-closure-config-panel) .snow-hold-field label {
    color: var(--snow-text);
  }

  #${APP.panelId} .snow-hold-field {
    display: grid;
    gap: 4px;
  }

  #${APP.panelId} .snow-closure-select-wrap .snow-select-trigger {
    width: 100%;
    height: var(--snow-control-height);
    min-height: var(--snow-control-height);
    box-sizing: border-box;
    padding: 4px 34px 4px 8px;
    border: 1px solid var(--snow-line);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    color: var(--snow-text);
    font: inherit;
    font-size: var(--snow-font-md);
  }

  #${APP.panelId} .snow-closure-select-wrap .snow-select-dropdown {
    z-index: 100;
  }

  #${APP.panelId} .snow-hold-field label {
    color: var(--snow-white);
    font-size: var(--snow-font-sm);
    font-weight: var(--snow-weight-bold);
  }

  #${APP.panelId} .snow-hold-reason-wrap .snow-select-trigger {
    width: 100%;
    height: var(--snow-control-height);
    min-height: var(--snow-control-height);
    box-sizing: border-box;
    padding: 4px 34px 4px 8px;
    border: 1px solid var(--snow-line);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    color: var(--snow-text);
    font: inherit;
    font-size: var(--snow-font-md);
  }

  #${APP.panelId} .snow-hold-reason-wrap .snow-select-trigger::placeholder {
    color: var(--snow-text-subtle);
    opacity: 1;
  }

  #${APP.panelId} .snow-hold-reason-wrap .snow-select-dropdown {
    z-index: 100;
  }

  #${APP.panelId} .snow-hold-field input:not([type=range]):not(.snow-select-trigger) {
    position: static;
    display: block;
    width: 100%;
    height: var(--snow-control-height);
    min-height: var(--snow-control-height);
    box-sizing: border-box;
    padding: 5px 8px;
    border: 1px solid var(--snow-line);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    color: var(--snow-text);
    opacity: 1;
    pointer-events: auto;
    font: inherit;
    font-size: var(--snow-font-md);
  }

  #${APP.panelId} .snow-hold-duration-heading {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
  }

  #${APP.panelId} .snow-hold-duration-heading output {
    color: var(--snow-accent);
    font-size: var(--snow-font-sm);
    font-weight: var(--snow-weight-bold);
  }

  #${APP.panelId} .snow-hold-field input[type=range] {
    width: 100%;
    height: 22px;
    margin: 3px 0 0;
    padding: 0;
    appearance: none;
    -webkit-appearance: none;
    background: transparent;
    --snow-hold-progress: 0%;
    --snow-hold-bright-start: 0%;
    cursor: pointer;
  }

  #${APP.panelId} .snow-hold-field input[type=range]::-webkit-slider-runnable-track {
    height: 7px;
    border: 1px solid var(--snow-line);
    border-radius: var(--snow-radius-pill);
    background: linear-gradient(to right, var(--snow-accent-dark) 0%, var(--snow-accent) var(--snow-hold-bright-start), var(--snow-accent) var(--snow-hold-progress), var(--snow-line-light) var(--snow-hold-progress), var(--snow-line-light) 100%);
  }

  #${APP.panelId} .snow-hold-field input[type=range]::-moz-range-track {
    height: 7px;
    border: 1px solid var(--snow-line);
    border-radius: var(--snow-radius-pill);
    background: var(--snow-line-light);
  }

  #${APP.panelId} .snow-hold-field input[type=range]::-moz-range-progress {
    height: 7px;
    border: 0;
    border-radius: var(--snow-radius-pill);
    background: linear-gradient(to right, var(--snow-accent-dark) 0%, var(--snow-accent) 35%, var(--snow-accent) 100%);
  }

  #${APP.panelId} .snow-hold-field input[type=range]::-webkit-slider-thumb {
    -webkit-appearance: none;
    appearance: none;
    width: 17px;
    height: 17px;
    margin-top: -6px;
    border: 2px solid var(--snow-white);
    border-radius: var(--snow-radius-base);
    background: var(--snow-accent);
    box-shadow: 0 0 0 1px var(--snow-accent), 0 1px 3px rgba(0,0,0,.25);
  }

  #${APP.panelId} .snow-hold-field input[type=range]::-moz-range-thumb {
    width: 13px;
    height: 13px;
    border: 2px solid var(--snow-white);
    border-radius: 50%;
    background: var(--snow-accent);
    box-shadow: 0 0 0 1px var(--snow-accent), 0 1px 3px rgba(0,0,0,.25);
  }

  #${APP.panelId} .snow-hold-range-labels {
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 2px;
    color: var(--snow-text);
    font-size: var(--snow-font-sm);
    font-weight: var(--snow-weight-bold);
    text-align: center;
  }

  #${APP.panelId} .snow-hold-config-actions {
    display: flex;
    justify-content: flex-end;
    gap: 6px;
  }

  #${APP.panelId} .snow-hold-cancel {
    min-height: 29px;
    padding: 4px 10px;
    border: 1px solid var(--snow-line);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    color: var(--snow-text);
    cursor: pointer;
    font-size: var(--snow-font-sm);
  }

  @keyframes snow-hold-panel-in {
    from { opacity: 0; transform: translateY(-4px); }
    to { opacity: 1; transform: translateY(0); }
  }

  /* Il pannello On Hold ha priorità d'interazione mentre è aperto. */
  #${APP.panelId}.hold-config-active .snow-body > *:not(#snow-resolution-wrapper) {
    pointer-events: none;
  }

  /* ===== Ticket aperti ===== */
  #${APP.panelId} .snow-open-tickets-title {
    margin-top: 13px;
    padding: 7px 0 5px;
    color: var(--snow-text-muted);
    font-size: var(--snow-font-sm);
    font-weight: var(--snow-weight-bold);
    text-transform: uppercase;
    letter-spacing: .35px;
  }

  #${APP.panelId} .snow-open-tickets-list {
    display: flex;
    flex-direction: column;
    gap: 2px;
    padding: 0 4px 8px 0;
  }

  #${APP.panelId} .snow-open-ticket-row {
    position: relative;
    display: grid;
    grid-template-columns: minmax(0, 1fr);
    align-items: stretch;
    min-height: 62px;
    padding: 7px 2px 7px 8px;
    box-sizing: border-box;
    border-left: 3px solid transparent;
    border-radius: 0;
    transition: background .12s ease,border-color .12s ease;
  }

  /* Stessa area di attivazione per restringere il testo e mostrare i pulsanti. */
  #${APP.panelId} .snow-open-ticket-row:hover:where(:not(:disabled, [aria-disabled="true"])) .snow-open-ticket-info,
  #${APP.panelId} .snow-open-ticket-row:focus-within .snow-open-ticket-info,
  #${APP.panelId} .snow-open-ticket-row.selected .snow-open-ticket-info {
    width: calc(100% - 94px);
    max-width: calc(100% - 94px);
  }

  #${APP.panelId} .snow-open-ticket-row:hover:where(:not(:disabled, [aria-disabled="true"])),
  #${APP.panelId} .snow-open-ticket-row:focus-within,
  #${APP.panelId} .snow-open-ticket-row.selected {
    background: transparent;
    border-left-color: var(--snow-accent);
  }

  #${APP.panelId} .snow-open-ticket-info {
    display: flex;
    flex-direction: column;
    align-items: flex-start;
    gap: 2px;
    min-width: 0;
    padding: 0;
    border: 0;
    background: transparent;
    box-shadow: none;
    text-align: left;
    cursor: pointer;
  }

  #${APP.panelId} .snow-ticket-number-line {
    display: flex;
    align-items: center;
    justify-content: flex-start;
    gap: 6px;
    width: 100%;
    min-width: 0;
    line-height: 16px;
  }

  #${APP.panelId} .snow-ticket-status-indicator {
    position: relative;
    flex: 0 0 16px;
    display: grid;
    place-items: center;
    width: 16px;
    height: 16px;
    box-sizing: border-box;
    margin: 0;
    border-radius: var(--snow-radius-base);
    background: var(--snow-accent);
    color: var(--snow-white);
    cursor: help;
  }

  #${APP.panelId} .snow-ticket-status-indicator svg {
    width: 10px;
    height: 10px;
    stroke-width: 0;
    stroke-linecap: butt;
    stroke-linejoin: miter;
    fill: currentColor;
  }

  #${APP.panelId} .snow-ticket-status-indicator.expired {
    background: var(--snow-danger);
    color: var(--snow-white);
    font-size: var(--snow-font-sm);
    font-weight: var(--snow-weight-heavy);
    line-height: 1;
    text-shadow: none;
    animation: snow-ticket-expired-blink 1s ease-in-out infinite;
  }

  #${APP.panelId} .snow-ticket-status-indicator.expired.acknowledged {
    animation: none;
    opacity: 1;
  }

  @keyframes snow-ticket-expired-blink {
    0%, 100% { opacity: 1; }
    50% { opacity: .3; }
  }

  #${APP.panelId} .snow-ticket-number {
    color: var(--snow-accent);
    font-size: var(--snow-font-md);
    font-weight: var(--snow-weight-bold);
    line-height: 1.25;
  }

  #${APP.panelId} .snow-open-ticket-actions {
    position: absolute;
    z-index: 2;
    right: 2px;
    top: 50%;
    transform: translateY(-50%);
    display: flex;
    align-items: center;
    justify-content: flex-end;
    gap: 3px;
    width: 90px;
    margin-left: 0;
    visibility: hidden;
    opacity: 0;
    pointer-events: none;
  }

  #${APP.panelId} .snow-open-ticket-row:hover:where(:not(:disabled, [aria-disabled="true"])) .snow-open-ticket-actions,
  #${APP.panelId} .snow-open-ticket-row:focus-within .snow-open-ticket-actions,
  #${APP.panelId} .snow-open-ticket-row.selected .snow-open-ticket-actions {
    visibility: visible;
    opacity: 1;
    pointer-events: auto;
  }

  #${APP.panelId} .snow-ticket-action {
    width: 27px;
    height: 27px;
    padding: 0;
    border: 1px solid var(--snow-line-strong);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    color: var(--snow-text-muted);
    display: grid;
    place-items: center;
    cursor: pointer;
  }

  #${APP.panelId} .snow-ticket-action:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-bg);
    border-color: var(--snow-accent-border);
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-style-record {
    background: transparent;
    box-shadow: none;
  }

  #${APP.panelId} .snow-ticket-company {
    max-width: 100%;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    color: var(--snow-text-muted);
    font-size: var(--snow-font-sm);
    font-weight: var(--snow-weight-bold);
    line-height: 1.3;
  }

  #${APP.panelId} .snow-ticket-activity {
    max-width: 100%;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    color: var(--snow-text-subtle);
    font-size: var(--snow-font-sm);
    font-weight: var(--snow-weight-normal);
    line-height: 1.3;
  }

  /* ===== Menu impostazioni ===== */
  #${APP.panelId} .snow-settings-wrapper {
    position: relative;
    z-index: 40;
  }

  #${APP.panelId} .snow-settings-menu {
    display: none;
    position: absolute;
    top: 35px;
    right: -2px;
    width: 246px;
    padding: 0;
    overflow: hidden;
    border: 1px solid var(--snow-line);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    box-shadow: 0 6px 18px rgba(0,0,0,.18);
    z-index: 10;
  }

  #${APP.panelId} .snow-settings-menu.open {
    display: flex;
    flex-direction: column;
    gap: 2px;
  }

  #${APP.panelId} .snow-settings-menu button {
    display: flex;
    align-items: center;
    gap: 8px;
    width: 100%;
    padding: 8px;
    text-align: left;
    cursor: pointer;
    font-size: var(--snow-font-md);
  }

  #${APP.panelId} .snow-settings-menu button:where(:not(.snow-style-list-item)) {
    border: 0;
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    color: var(--snow-text);
  }

  #${APP.panelId} .snow-settings-menu button:where(:not(.snow-style-list-item)):hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-bg);
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-settings-section {
    padding: 10px;
  }

  #${APP.panelId} .snow-settings-section+.snow-settings-section {
    border-top: 1px solid var(--snow-line-light);
  }

  #${APP.panelId} .snow-settings-title {
    margin-bottom: 7px;
    color: var(--snow-text);
    font-size: var(--snow-font-sm);
    font-weight: var(--snow-weight-bold);
    text-transform: uppercase;
    letter-spacing: .3px;
  }

  #${APP.panelId} .snow-settings-label {
    margin-bottom: 5px;
    color: var(--snow-text-subtle);
    font-size: var(--snow-font-xs);
  }

  #${APP.panelId} .snow-settings-segment {
    display: grid;
    grid-template-columns: 1fr 1fr 1fr;
    border: 1px solid var(--snow-line);
    border-radius: var(--snow-radius-base);
    overflow: hidden;
  }

  #${APP.panelId} .snow-settings-segment button {
    justify-content: center;
    padding: 7px 4px;
    border-radius: 0;
    border-right: 1px solid var(--snow-line-light);
    font-size: var(--snow-font-xs);
  }

  #${APP.panelId} .snow-settings-segment button:last-child {
    border-right: 0;
  }

  #${APP.panelId} .snow-settings-segment button.active {
    background: var(--snow-accent-bg);
    color: var(--snow-accent);
    font-weight: var(--snow-weight-bold);
  }

  #${APP.panelId} .snow-settings-actions {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 5px;
  }

  #${APP.panelId} .snow-settings-actions button {
    justify-content: center;
    padding: 7px 5px;
    border: 1px solid var(--snow-line);
    font-size: var(--snow-font-xs);
  }

  /* ===== Finestra configurazioni salvate ===== */

  #${APP.panelId} .snow-activity-summary-overlay {
    position: fixed;
    inset: 0;
    z-index: 2147483647;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 18px;
    background: rgba(25,25,25,.38);
  }

  #${APP.panelId} .snow-activity-summary-panel {
    display: flex;
    flex-direction: column;
    width: min(560px, calc(100vw - 36px));
    max-height: min(78vh, 760px);
    overflow: hidden;
    border: 1px solid var(--snow-line);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    box-shadow: 0 12px 36px rgba(0,0,0,.24);
  }

  #${APP.panelId} .snow-activity-summary-header {
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    gap: 12px;
    padding: 15px 16px 12px;
    border-bottom: 1px solid var(--snow-line-light);
  }

  #${APP.panelId} .snow-activity-summary-title {
    color: var(--snow-text);
    font-size: var(--snow-font-xl);
    font-weight: var(--snow-weight-bold);
  }

  #${APP.panelId} .snow-activity-summary-subtitle {
    margin-top: 4px;
    color: var(--snow-text-subtle);
    font-size: var(--snow-font-sm);
  }

  #${APP.panelId} .snow-activity-summary-close {
    flex: 0 0 auto;
    width: 28px;
    height: 28px;
    justify-content: center;
    padding: 0;
    border: 1px solid var(--snow-line);
    border-radius: var(--snow-radius-base);
    color: var(--snow-text-muted);
    font-size: 21px;
    line-height: 1;
  }

  #${APP.panelId} .snow-activity-summary-close:hover:where(:not(:disabled, [aria-disabled="true"])) {
    border-color: var(--snow-accent-border);
    background: var(--snow-accent-bg);
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-activity-summary-tabs {
    display: flex;
    gap: 6px;
    padding: 6px 12px 1px;
  }

  #${APP.panelId} .snow-activity-summary-tab {
    flex: 1;
    justify-content: center;
    padding: 8px 10px 6px;
    border: 0;
    border-bottom: 3px solid transparent;
    border-radius: 0;
    background: transparent;
    color: var(--snow-text-muted);
    font-size: var(--snow-font-md);
    font-weight: var(--snow-weight-bold);
    box-shadow: none;
  }

  #${APP.panelId} .snow-activity-summary-tab.active {
    border-bottom-color: var(--snow-accent);
    background: transparent;
    color: var(--snow-accent);
    font-weight: var(--snow-weight-bold);
    box-shadow: none;
  }

  #${APP.panelId} .snow-activity-summary-tab:hover:where(:not(:disabled, [aria-disabled="true"])):not(.active) {
    background: transparent;
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-activity-summary-list {
    overflow-y: auto;
    padding: 10px 12px 14px;
  }

  #${APP.panelId} .snow-activity-summary-card {
    padding: 11px 12px;
    border: 1px solid var(--snow-line-light);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
  }

  #${APP.panelId} .snow-activity-summary-card + .snow-activity-summary-card {
    margin-top: 8px;
  }

  #${APP.panelId} .snow-activity-summary-card h3 {
    margin: 0 0 8px;
    color: var(--snow-accent);
    font-size: var(--snow-font-md);
    font-weight: var(--snow-weight-bold);
  }

  #${APP.panelId} .snow-activity-summary-card dl {
    display: grid;
    gap: 5px;
    margin: 0;
  }

  #${APP.panelId} .snow-activity-summary-row {
    display: grid;
    grid-template-columns: minmax(120px, 38%) minmax(0, 1fr);
    gap: 10px;
    font-size: var(--snow-font-sm);
    line-height: 1.35;
  }

  #${APP.panelId} .snow-activity-summary-row dt {
    color: var(--snow-text-subtle);
  }

  #${APP.panelId} .snow-activity-summary-row dd {
    min-width: 0;
    margin: 0;
    color: var(--snow-text);
    overflow-wrap: anywhere;
  }

  #${APP.panelId} .snow-activity-summary-empty {
    padding: 22px 10px;
    color: var(--snow-text-subtle);
    text-align: center;
    font-size: var(--snow-font-md);
  }

  /* ===== Messaggi di stato ===== */
  #${APP.panelId} .snow-status-area {
    flex: 0 0 auto;
    padding: 0 14px 12px;
    background: var(--snow-white);
    border-top: 1px solid transparent;
  }

  #${APP.panelId} .snow-status-area:has(.snow-status.success),
  #${APP.panelId} .snow-status-area:has(.snow-status.error),
  #${APP.panelId} .snow-status-area:has(.snow-status.warning) {
    padding-top: 10px;
    border-top-color: var(--snow-line-light);
  }

  #${APP.panelId} .snow-status {
    display: none;
    margin-top: 0;
    padding: 8px 9px;
    border-radius: var(--snow-radius-base);
    font-size: var(--snow-font-sm);
    line-height: 1.4;
  }

  #${APP.panelId} .snow-status.success {
    display: block;
    background: var(--snow-status-success-bg);
    border: 1px solid var(--snow-status-success-border);
    color: var(--snow-status-success-text);
  }

  #${APP.panelId} .snow-status.error {
    display: block;
    background: var(--snow-status-error-bg);
    border: 1px solid var(--snow-status-error-border);
    color: var(--snow-status-error-text);
  }

  #${APP.panelId} .snow-status.warning {
    display: block;
    background: var(--snow-status-warning-bg);
    border: 1px solid var(--snow-status-warning-border);
    color: var(--snow-status-warning-text);
  }

  /* ===== Stili condivisi (classi snow-style-*) ===== */
  #${APP.panelId} .snow-style-button {
    -webkit-tap-highlight-color: transparent;
  }

  #${APP.panelId} .snow-style-button:focus,
  #${APP.panelId} .snow-style-button:focus-visible {
    outline: none;
    box-shadow: none;
  }

  #${APP.panelId} .snow-style-button svg {
    stroke: currentColor;
  }

  #${APP.panelId} .snow-style-list-item {
    background: transparent;
    border: 0;
    border-left: 3px solid transparent;
    border-radius: 0;
    box-shadow: none;
    color: var(--snow-text);
  }

  #${APP.panelId} .snow-style-list-item:hover:where(:not(:disabled, [aria-disabled="true"])),
  #${APP.panelId} .snow-style-list-item:focus-visible {
    background: var(--snow-accent-bg);
    border-left-color: transparent;
    border-radius: var(--snow-radius-base);
    color: var(--snow-accent);
  }

  #${APP.panelId} .snow-style-list-item:hover:where(:not(:disabled, [aria-disabled="true"])) *,
  #${APP.panelId} .snow-style-list-item:focus-visible * {
    color: inherit;
  }

  /* Voce selezionata: stesso aspetto in tutti gli elenchi (Preferiti, tendine, ...).
     .selected è usata dai Preferiti, aria-selected dalle tendine. */
  #${APP.panelId} .snow-style-list-item.selected,
  #${APP.panelId} .snow-style-list-item[aria-selected="true"] {
    color: var(--snow-accent);
    font-weight: var(--snow-weight-bold);
  }

  #${APP.panelId} .snow-style-list-item svg {
    color: inherit;
    stroke: currentColor;
  }

  #${APP.panelId} .snow-style-icon-button {
    display: grid;
    place-items: center;
  }

  #${APP.panelId} .snow-style-icon-button svg {
    margin: auto;
    stroke: currentColor;
  }

  #${APP.panelId} .snow-style-danger {
    color: var(--snow-text-muted);
    border-color: var(--snow-line-strong);
  }

  #${APP.panelId} .snow-style-danger:hover:where(:not(:disabled, [aria-disabled="true"])) {
    color: var(--snow-danger);
    border-color: var(--snow-danger);
    background: var(--snow-danger-bg);
  }

  /* ===== Tour guidato ===== */
  .snow-tour-overlay {
    position: fixed;
    inset: 0;
    z-index: 2147483645;
    background: rgba(0,0,0,.18);
  }

  .snow-tour-balloon {
    position: fixed;
    z-index: 2147483647;
    width: min(260px, calc(100vw - 46px));
    padding: 14px;
    border: 1px solid var(--snow-line);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    color: var(--snow-text);
    box-shadow: 0 8px 24px rgba(0,0,0,.25);
    font-family: var(--snow-font-family);
  }

  .snow-tour-progress {
    margin-bottom: 5px;
    color: var(--snow-accent);
    font-size: var(--snow-font-xs);
    font-weight: var(--snow-weight-bold);
    text-transform: uppercase;
  }

  .snow-tour-title {
    margin-bottom: 6px;
    font-size: var(--snow-font-lg);
    font-weight: var(--snow-weight-bold);
  }

  .snow-tour-text {
    font-size: var(--snow-font-md);
    line-height: 1.45;
    color: var(--snow-text-muted);
  }

  .snow-tour-actions {
    display: grid;
    grid-template-columns: auto 1fr auto auto;
    gap: 5px;
    align-items: center;
    margin-top: 12px;
  }

  .snow-tour-actions button {
    padding: 6px 8px;
    border: 1px solid var(--snow-line);
    border-radius: var(--snow-radius-base);
    background: var(--snow-white);
    color: var(--snow-text-muted);
    cursor: pointer;
    font-size: var(--snow-font-sm);
  }

  .snow-tour-actions #snow-tour-next {
    border-color: var(--snow-accent);
    background: var(--snow-accent);
    color: var(--snow-white);
    font-weight: var(--snow-weight-bold);
  }

  .snow-tour-actions #snow-tour-next:hover:where(:not(:disabled, [aria-disabled="true"])) {
    background: var(--snow-accent-dark);
    border-color: var(--snow-accent-dark);
    color: var(--snow-white);
  }

  #${APP.panelId} .snow-tour-layer {
    z-index: 2147483646;
  }

  /* Evidenziazione del tour: unico aspetto per qualunque elemento indicato.
     Gli !important qui sono voluti: l'evidenziazione deve prevalere
     sullo stile proprio dell'elemento, qualunque esso sia. */
  #${APP.panelId} .snow-tour-target {
    position: relative;
    z-index: 2147483646 !important;
    color: var(--snow-accent) !important;
    background: var(--snow-accent-bg) !important;
    border-color: var(--snow-accent-border) !important;
    box-shadow: none !important;
    outline: none !important;
  }

  #${APP.panelId} .snow-tour-target svg {
    color: inherit !important;
    stroke: currentColor !important;
  }

  /* Voci degli elenchi (Preferiti, Risposte salvate, voci del menu): nel tour hanno lo stesso
     aspetto che al passaggio del mouse, senza bordo colorato (niente barretta laterale). */
  #${APP.panelId} .snow-tour-target.snow-style-list-item {
    border-color: transparent !important;
    border-radius: var(--snow-radius-base) !important;
  }
  `;
        document.head.appendChild(style);
    }

    // ======================================================================
    // 15. AVVIO E COLLEGAMENTO CON L'ADD-IN
    // L'add-in chiama AmpModule.show() dal pulsante "Apri area AMP".
    // ======================================================================

    let initialized = false;
    function initialize() {
        if (initialized) return;
        const container = document.getElementById('amp-view');
        if (!container) return;
        // Il modulo Incident è nella stessa pagina (anteprima) oppure non c'è (add-in senza connettori):
        // in quel caso le letture restituiscono campi vuoti e il pannello lo segnala.
        const context = findServiceNowContext();
        snWindow = context?.win || window;
        snDocument = context?.doc || document;
        buildPanel(container);
        syncTrackedTicketState();
        document.getElementById('open-amp-button')?.addEventListener('click', show);
        initialized = true;
        console.log(`[${APP.name}] pronto.`);
    }
    function show() {
        initialize();
        document.getElementById('main-view')?.setAttribute('hidden', '');
        document.getElementById('header-view')?.setAttribute('hidden', '');
        const view = document.getElementById('amp-view');
        if (view) view.hidden = false;
        regenerateResponsePreview();
    }
    window.AmpModule = Object.freeze({ initialize, show });
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initialize);
    else initialize();
})();
