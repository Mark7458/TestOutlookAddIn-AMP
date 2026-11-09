# Area AMP: servizi da collegare

L'area AMP (`amp.js`) gestisce le richieste della casella condivisa e i ticket ServiceNow
direttamente da Outlook. Tutto ciò che sta fuori dal pannello passa da un unico punto
(blocco 3 di `amp.js`): oggi lo fornisce `amp-preview-backend.js` (simulato, solo per
l'anteprima); nell'add-in va fornito dalla Function App, con l'accesso aziendale già usato
per `GetTicketByConversationId`.

Nessun file dell'add-in è stato modificato: l'area usa solo `#amp-view`, `#open-amp-button`
e `TestPluginAPI` (`getCurrentMessageContext`, `onMessageChanged`, `showMainView`).

## Regole

1. **ServiceNow resta il riferimento, SharePoint è la copia che il pannello legge.** Elenchi,
   conteggi, scadenze e assegnatario si leggono da SharePoint, senza interrogare ServiceNow.
   Lo storico completo di ServiceNow (commenti di tutti) si legge solo su richiesta.
2. **Ogni movimento: prima ServiceNow, poi SharePoint.** Apertura, tempo, assegnazione, risposta,
   chiusura scrivono su ServiceNow; solo se riesce si aggiornano la copia del ticket, il thread e
   il registro dei movimenti su SharePoint.
3. **Allineamento periodico.** Un job della Function App (ogni 10-15 minuti) legge da ServiceNow,
   in una sola chiamata, i ticket del gruppo modificati dall'ultimo allineamento
   (`sys_updated_on` > ultimo allineamento) e aggiorna SharePoint. Così entrano anche le modifiche
   fatte direttamente su AMP fuori dal pannello (colleghi, chiusure automatiche, riassegnazioni).
4. **Chi ha in carico un ticket = *Assigned to*.** Non si usano i tag di Outlook. Aprire il ticket
   equivale a prendere in carico il thread (assegnato a chi lo apre).
5. **In coda = un thread per conversationId, coda condivisa.** Ci sono i thread non in gestione.
   Quando tutti i ticket del thread si chiudono il thread torna **libero subito**, ma ricompare
   **solo se arriva un messaggio nuovo**. I thread **ignorati** vanno nella sezione *Ignorate* di In coda
   (accanto ad *Attive*), da cui si possono prendere in carico; tornano da soli in coda con un messaggio nuovo.
   Un thread in gestione resta a chi l'ha preso in carico e si segue dalla scheda **Ticket**:
   i **messaggi nuovi** (arrivati dopo l'ultima risposta o dopo l'ultimo clic sul ticket di chi
   l'ha in carico) compaiono come **+N** con l'inizio dell'ultimo messaggio, anche se qualcuno li ha
   già letti nella casella condivisa. Solo il clic sul ticket nel pannello li segna come visti.
6. **Un thread può avere più ticket**, nel tempo (richieste successive sulla stessa scia) o
   **aperti insieme** (es. creazione della DL e aggiunta dei membri). Il thread resta in gestione
   finché almeno uno è aperto.
7. **Principale e collegati.** Il **principale** è il primo ticket (aperto) collegato al thread: nel
   pannello ha l'icona della mail. **Con un solo ticket aperto** sul thread è legato alla **risposta via
   mail**, che lo aggiorna (con l'azione salvata con "Salva e compila", se c'è). **Con più ticket aperti** il
   pannello non può lavorare mentre si scrive la mail in Outlook: ogni ticket, principale compreso, si
   aggiorna e si chiude dal pannello (`updateTicket`: testo e azione subito, in qualunque ordine) e la mail
   la scrive l'operatore (il pannello prepara il testo da copiare); la risposta non tocca ServiceNow.
8. **Da In coda, con più messaggi nel thread,** l'operatore sceglie da quale messaggio nasce il ticket
   (frecce e "Usa questo"; mai in automatico): da quel messaggio vengono *Description* e *Reported*.

## Triangolazione mail ↔ ticket

Per una conversazione, vale l'indizio **più recente** tra:

- **numero INC citato nel thread** (oggetto o corpo: presa in carico, conferma, "Rif. INC…"
  che il pannello propone di inserire in ogni risposta);
- **collegamento registrato su SharePoint** (ticket aperto dall'add-in o confermato dall'operatore).

Se il ticket trovato è chiuso ed è arrivata una mail dopo la chiusura, è una **richiesta nuova**:
il ticket chiuso viene mostrato come "ticket precedente" e si propone di aprirne uno nuovo.

Se non c'è nessun indizio, si cercano **ticket simili** nella copia su SharePoint:
*Short description* = oggetto del thread (senza RE:/I:/prefissi tra parentesi) e *Reported* =
arrivo della mail (± 15 minuti). Sono solo proposte: il collegamento lo conferma l'operatore
(`linkTicket`).

## Servizi richiesti

Ogni servizio restituisce una Promise.

| Servizio | Cosa fa | Dove agisce |
|---|---|---|
| `listQueue()` | Thread delle caselle condivise (uno per conversationId) da prendere in carico, secondo la regola 5, con `mailbox` = casella da cui arriva (AMPCoE, M365 Admin Center), mostrata come etichetta | Graph + SharePoint |
| `listIgnored()` | Archivio dei thread ignorati in attesa del prossimo messaggio | Graph + SharePoint |
| `markThreadSeen(conversationId, utente)` | Chi ha il thread ha cliccato il ticket: i messaggi fin qui non sono più nuovi | SharePoint |
| `listOpenTickets()` | Ticket In Progress e On Hold del gruppo, con la mail collegata, `newMessages`, `lastNewMessage` (chi, quando, prime righe), `lastMessageAt` (ora dell'ultimo messaggio arrivato nel thread, anche già visto) e l'ora dell'ultimo allineamento (`syncedAt`) | Graph + SharePoint |
| `findTicketForMail(mail)` | Triangolazione: `{ tickets (aperti, anche più di uno, ognuno con via e primary = principale del thread), previous, candidates, discarded }` | SharePoint |
| `getTicket(numero)` | Un ticket | SharePoint |
| `linkTicket(conversationId, numero, utente)` | L'operatore conferma il collegamento: il thread passa in gestione | SharePoint |
| `createTicket(mail, dati)` | Crea l'Incident assegnato a chi lo apre; poi copia, thread in gestione, movimento. Con `mail = null` è un ticket aperto da Ticket senza mail: nessun thread da collegare, Caller = chi lo apre, `reportedOn` = data e ora scritte dall'operatore | ServiceNow → SharePoint |
| `assignTicket(numero, utente)` | Prende in carico il ticket di un collega (*Assigned to*) | ServiceNow → SharePoint |
| `getTicketActivity(numero, { offset, limit })` | Storico completo, dal più recente, a pagine (solo su richiesta) | ServiceNow (journal) |
| `setPendingAction(numero, azione)` | Cosa fare all'invio della risposta (`null` = solo commento). Azione = `{ type: comment \| resolve \| hold \| cancel, by, resolutionCode (resolve), holdReason e holdDays (hold, 1–7 giorni), addMinutes }`; `addMinutes`: tempo lavorato da aggiungere (o togliere) insieme | SharePoint |
| `updateTicket(numero, azione, testo, utente)` | Ticket senza mail o di un thread con più ticket aperti, con l'eventuale `addMinutes` dell'azione: testo e azione subito, scritti come all'invio (punto 2 di "All'invio della risposta") | ServiceNow → SharePoint |
| `searchTickets(testo)` | "Cerca su AMP": Incident con il testo nel numero o nella short description, anche chiusi e di altri operatori (i più recenti, max 20). Ogni ticket come quello restituito, più `resolutionNotes`, `lastWorkNote` = `{ author, at, text }` (l'ultima *Work note* dal journal) o `null` e `readOnly: true`. Il pannello li mostra in sola lettura, con "Apri su AMP" (`url`) | ServiceNow |
| `discardMail(conversationId, utente)` / `restoreMail(conversationId)` | Thread che non richiede un ticket (Ignora); torna In coda con un messaggio nuovo | SharePoint |
| `forwardMail(conversationId, gruppo, utente)` | Inoltro della richiesta a un reparto interno (PO, ES, PEL); il thread lascia In coda. Sequenza in "Inoltro (FWD)" | Graph (+ ServiceNow) |
| `subscribe(listener)` | Facoltativo: avvisa il pannello quando i dati cambiano | — |
| `listThreadMessages(conversationId)` | Messaggi ricevuti del thread, dal più vecchio: `[{ messageId, sender, receivedAt, preview }]`. In In coda, con più messaggi, l'operatore sceglie da quale aprire il ticket | Graph |
| `getMailBody(messageId)` | Testo (semplice, senza HTML) di una mail della casella condivisa: in In coda si legge la richiesta e la si prende in carico o la si ignora senza aprirla in Outlook | Graph |
| `sync()` | Job periodico della Function App (non lo chiama il pannello): allinea SharePoint a ServiceNow | ServiceNow → SharePoint |

- `mail` = `{ conversationId, itemId, subject, sender: { name, email }, receivedAt, body, incNumbers }`
  (`incNumbers` = numeri INC letti dalla mail aperta, usati come ultima riserva).
- `utente` = `{ displayValue, value }` (nome e sys_id del profilo nel pannello).
- `createTicket` riceve `fields` con: `company`, `assignmentGroup`, `assignedTo`, `category`
  (riferimenti `{ value: sys_id, displayValue }`), `impact`, `urgency`, `timeWorkedMinutes`,
  `shortDescription` (oggetto + prefisso dell'Attività), `description` (corpo della mail, max 4000
  caratteri), `reportedOn` (arrivo della mail, `AAAA-MM-GG hh:mm:ss`).
- Ticket restituito: `{ number, sysId, url, state, assignedTo: { name, sysId }, companyName,
  activityName, shortDescription, description, reportedOn, openedAt, closedAt, timeWorkedMinutes, holdReason,
  holdExpiration, conversationId, messageId, pendingAction }`.

## SharePoint: le liste

| Lista | Una riga per | Campi |
|---|---|---|
| **Thread AMP** | conversazione | `conversationId`, `internetMessageId` e immutable ID dell'ultima mail, `status` (libero / in gestione / scartato), `owner` (chi l'ha in carico), `ownerSeenAt` (ultimo clic sul ticket di chi l'ha in carico), `tickets` (JSON: `number`, `linkedAt`, `linkedBy`, `confirmedBy` per ogni ticket del thread), `freedAt`, `discardedBy`, `discardedAt`, `updatedAt` |
| **Ticket AMP** | ticket del gruppo | copia dei valori di ServiceNow: `number`, `sysId`, `state`, `assignedTo`, `company`, `activity`, `shortDescription`, `description`, `reportedOn`, `openedAt`, `closedAt`, `timeWorkedMinutes`, `holdReason`, `holdExpiration`, `resolutionCode`; più `pendingAction` (JSON dell'azione all'invio) e `updatedAt` |
| **Movimenti AMP** | azione | `at`, `type` (apertura, collegamento, tempo, assegnazione, risposta, azione all'invio, ignorata, allineamento…), `ticketNumber`, `by`, `details`, `conversationId`, esito su ServiceNow |

Pulizia: thread liberi, ticket chiusi e movimenti più vecchi di N mesi si cancellano in automatico.

## All'invio della risposta (lato server)

Quando dalla casella parte una risposta in una conversazione con ticket aperti:

0. solo se il thread ha **un solo ticket aperto** (regola 7): si aggiorna quello; con più ticket aperti la risposta non tocca ServiceNow;
1. il testo della mail va su ServiceNow come *Additional comments* del ticket (come il pannello AMP su Tampermonkey);
2. si applica l'azione salvata (confermata con "Salva e compila"):
   - se l'azione ha `addMinutes`, il tempo lavorato si aggiorna di quei minuti (impostati nel timer del pannello);
   - `comment` (Aggiorna) → solo *Additional comments*; su un ticket On Hold → torna In Progress;
   - `resolve` / `cancel` → stato Resolved / Canceled, *Resolution code* (resolve), *Resolution notes* = testo della mail
     seguito dal nome dell'operatore;
   - `hold` → stato On Hold, *On hold reason*, *On hold expiration* = momento in cui l'azione parte + `holdDays`
     giorni (da 1 a 7) + 5 minuti, *Work notes* = testo della mail;
3. solo se ServiceNow risponde OK si aggiornano la copia del ticket e il thread su SharePoint
   (ticket chiuso → thread libero), si registra il movimento e l'azione si cancella.

Proposta: farlo nella Function App osservando la Posta inviata della casella con una
sottoscrizione Microsoft Graph, così funziona anche se l'add-in è chiuso. Da verificare con chi
gestisce Exchange: rispondendo dalla casella condivisa, di base la copia inviata può finire nella
Posta inviata personale di chi scrive; va attivato il salvataggio della copia anche nella casella
condivisa (`MessageCopyForSentAsEnabled` / `MessageCopyForSendOnBehalfEnabled`). In alternativa l'evento
`OnMessageSend` dell'add-in (va dichiarato nel manifest e verificato sulle caselle condivise).

## Outlook (Office.js) usato dal pannello

- Un add-in non può selezionare una mail nell'elenco di Outlook: `displayMessageForm(messageId)` la apre
  in una finestra (o scheda) a parte. Per questo In coda mostra il testo della mail (`getMailBody`) e la
  lavorazione del ticket segue la mail selezionata dall'operatore (evento di cambio mail dell'add-in).
- Rispondere con un testo pronto: `Office.context.mailbox.item.displayReplyForm(...)` (solo sulla mail aperta).
- Elenco delle mail della casella condivisa: non disponibile da Office.js → Graph nella Function App.

## Inoltro (FWD)

Quando l'operatore sceglie Inoltra e un gruppo (PO, ES o PEL), in quest'ordine:

1. **apre il ticket** (assegnato a chi inoltra);
2. **risponde alla mail** con il testo di presa in carico e il riferimento al ticket;
3. **controlla che la casella M365** sia in A o in Cc e, se manca, la aggiunge in Cc;
4. **antepone all'oggetto** `[PO]`, `[ES]` o `[PEL]` secondo la scelta;
5. **invia** la risposta;
6. **chiude il ticket** con il testo "Inoltrato a PO" / "Inoltrato a ES" / "Inoltrato a PEL".

Se un passaggio fallisce ci si ferma e il pannello lo mostra (per esempio: ticket aperto ma risposta non inviata),
così l'operatore può riprovare l'invio o annullare il ticket.

**Lato Outlook (passaggi 2–5), fattibile dal pannello con Microsoft Graph.** La risposta si crea come bozza
(createReply sulla mail della casella condivisa), si modificano testo, Cc e oggetto della bozza e poi si invia.
Le bozze restano modificabili anche dopo il 31/12/2026, quando Microsoft bloccherà le modifiche a oggetto e
destinatari delle mail già esistenti senza permessi speciali. Accesso con nested app authentication (MSAL).
Permessi necessari:
- registrazione dell'app (consenso amministratore): `Mail.ReadWrite.Shared` e `Mail.Send.Shared`;
- casella condivisa, per ogni operatore: accesso completo (Full Access) e "Invia come" (Send As) oppure
  "Invia per conto di" (Send on Behalf).

**Lato ServiceNow (passaggi 1 e 6).** Dal pannello servono OAuth per l'app e regole CORS sull'istanza;
altrimenti restano alla Function App. La prima versione probabilmente agirà solo lato Outlook: apertura e
chiusura del ticket le farà l'operatore, con un'interfaccia temporanea per i ticket da inoltrare (da definire).

## Riferimenti a ServiceNow (sys_id)

Nelle impostazioni Company, Assignment group, Category e Assigned to si scrivono per nome.
Il sys_id (l'identificativo interno di ServiceNow) non si mostra né si modifica: arriva dal file
importato da Tampermonkey oppure va trovato dal servizio cercando il nome su ServiceNow.
Le voci senza sys_id compaiono come "da verificare". Servizio da aggiungere:
`resolveReference(tabella, nome)` → `{ value: sys_id, displayValue }` (tabelle `core_company`,
`sys_user_group`, `sys_user` e quella delle categorie).

## Da definire

- Flusso FWD: testo di presa in carico, indirizzo della casella M365 (nelle impostazioni), interfaccia
  temporanea per i ticket da inoltrare nella prima versione (solo Outlook).
