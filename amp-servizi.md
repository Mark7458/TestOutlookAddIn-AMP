# Area AMP: servizi da collegare

L'area AMP (`amp.js`) gestisce le richieste della casella condivisa e i ticket ServiceNow
direttamente da Outlook. Tutto ciò che sta fuori dal pannello passa da un unico punto
(blocco 3 di `amp.js`): oggi lo fornisce `amp-preview-backend.js` (simulato, solo per
l'anteprima); nell'add-in va fornito dalla Function App, con l'accesso aziendale già usato
per `GetTicketByConversationId`.

Nessun file dell'add-in è stato modificato: l'area usa solo `#amp-view`, `#open-amp-button`
e `TestPluginAPI` (`getCurrentMessageContext`, `onMessageChanged`, `showMainView`).

## Regole

1. **ServiceNow è l'unica fonte per il ticket.** Stato, assegnatario, tempo lavorato, scadenza
   On Hold e storico si leggono sempre da ServiceNow. Il pannello non li conserva e SharePoint
   non li duplica.
2. **Chi ha in carico un ticket = *Assigned to*.** Non si usano i tag di Outlook. Una mail senza
   ticket è "in coda"; aprire il ticket equivale a prenderla in carico (assegnato a chi lo apre).
3. **Prima ServiceNow, poi SharePoint.** Ogni operazione scrive su ServiceNow; solo se riesce
   si registra su SharePoint. Se SharePoint fallisce, la triangolazione ricostruisce il collegamento.
4. **Un thread può avere più ticket nel tempo** (lo stesso thread si riusa per richieste nuove).

## Triangolazione mail ↔ ticket

Per una conversazione, vale l'indizio **più recente** tra:

- **numero INC citato nel thread** (oggetto o corpo: presa in carico, conferma, "Rif. INC…"
  che il pannello propone di inserire in ogni risposta);
- **collegamento registrato su SharePoint** (ticket aperto dall'add-in o confermato dall'operatore).

Se il ticket trovato è chiuso ed è arrivata una mail dopo la chiusura, è una **richiesta nuova**:
il ticket chiuso viene mostrato come "ticket precedente" e si propone di aprirne uno nuovo.

Se non c'è nessun indizio, si cercano **ticket simili** su ServiceNow:
*Short description* = oggetto del thread (senza RE:/I:/prefissi tra parentesi) e *Reported* =
arrivo della mail (± 15 minuti). Sono solo proposte: il collegamento lo conferma l'operatore
(`linkTicket`).

## Servizi richiesti

Ogni servizio restituisce una Promise.

| Servizio | Cosa fa | Dove agisce |
|---|---|---|
| `listQueue()` | Mail della casella senza ticket aperto e non scartate (con ticket precedente e possibili ticket) | Graph + triangolazione |
| `listOpenTickets()` | Ticket In Progress e On Hold della coda, con la mail collegata se c'è | ServiceNow (+ SharePoint per la mail) |
| `findTicketForMail(mail)` | Triangolazione: `{ ticket, via, previous, candidates, discarded }` | ServiceNow + SharePoint |
| `getTicket(numero)` | Un ticket | ServiceNow |
| `linkTicket(conversationId, numero, utente)` | L'operatore conferma il collegamento | SharePoint |
| `createTicket(mail, dati)` | Crea l'Incident assegnato a chi lo apre, poi registra il collegamento | ServiceNow → SharePoint |
| `assignTicket(numero, utente)` | Prende in carico il ticket di un collega (*Assigned to*) | ServiceNow |
| `addTime(numero, minuti, nome)` | Aggiunge (o toglie, se negativo) tempo lavorato | ServiceNow |
| `getTicketActivity(numero, { offset, limit })` | Storico, dal più recente, a pagine | ServiceNow (journal) |
| `setPendingAction(numero, azione)` | Cosa fare all'invio della risposta (`null` = solo commento) | SharePoint |
| `discardMail(conversationId, utente)` / `restoreMail(conversationId)` | Mail che non richiede un ticket; torna in coda se nel thread arriva una mail nuova | SharePoint |
| `subscribe(listener)` | Facoltativo: avvisa il pannello quando i dati cambiano | — |

- `mail` = `{ conversationId, itemId, subject, sender: { name, email }, receivedAt, body, incNumbers }`
  (`incNumbers` = numeri INC letti dalla mail aperta, usati come ultima riserva).
- `utente` = `{ displayValue, value }` (nome e sys_id del profilo nel pannello).
- `createTicket` riceve `fields` con: `company`, `assignmentGroup`, `assignedTo`, `category`
  (riferimenti `{ value: sys_id, displayValue }`), `impact`, `urgency`, `timeWorkedMinutes`,
  `shortDescription` (oggetto + prefisso dell'Attività), `description` (corpo della mail, max 4000
  caratteri), `reportedOn` (arrivo della mail, `AAAA-MM-GG hh:mm:ss`).
- Ticket restituito: `{ number, sysId, url, state, assignedTo: { name, sysId }, companyName,
  activityName, shortDescription, reportedOn, openedAt, closedAt, timeWorkedMinutes, holdReason,
  holdExpiration, conversationId, messageId, pendingAction }`.

## SharePoint: cosa si registra

Solo collegamenti e decisioni dell'operatore, mai i dati del ticket.

| Lista | Campi |
|---|---|
| Collegamenti | `conversationId`, `ticketNumber`, `linkedAt`, `linkedBy`, `confirmedBy` (se confermato a mano) |
| Mail scartate | `conversationId`, `discardedBy`, `discardedAt` |
| Azioni all'invio | `ticketNumber`, `pendingAction` (JSON: `type` comment / resolve / cancel / hold, `resolutionCode`, `holdReason`, `holdExpiration`, `workNotes`, `by`) |

## All'invio della risposta (lato server)

Quando dalla casella parte una risposta in una conversazione con ticket aperto:

1. il testo della mail va su ServiceNow come *Additional comment*;
2. si applica l'azione salvata:
   - `resolve` / `cancel` → stato Resolved / Canceled, *Resolution code*, *Resolution notes* = testo della mail;
   - `hold` → stato On Hold, *On hold reason*, *On hold expiration*;
   - nessuna azione su un ticket On Hold → torna In Progress;
   - `workNotes`, se presente → *Work notes*;
3. solo se ServiceNow risponde OK, l'azione si cancella da SharePoint.

Proposta: farlo nella Function App osservando la Posta inviata della casella con una
sottoscrizione Microsoft Graph, così funziona anche se l'add-in è chiuso. In alternativa l'evento
`OnMessageSend` dell'add-in (va dichiarato nel manifest e verificato sulle caselle condivise).

## Outlook (Office.js) usato dal pannello

- Aprire una mail: `Office.context.mailbox.displayMessageForm(messageId)`.
- Rispondere con un testo pronto: `Office.context.mailbox.item.displayReplyForm(...)` (solo sulla mail aperta).
- Elenco delle mail della casella condivisa: non disponibile da Office.js → Graph nella Function App.

## Riferimenti a ServiceNow (sys_id)

Nelle impostazioni Company, Assignment group, Category e Assigned to si scrivono per nome.
Il sys_id (l'identificativo interno di ServiceNow) non si mostra né si modifica: arriva dal file
importato da Tampermonkey oppure va trovato dal servizio cercando il nome su ServiceNow.
Le voci senza sys_id compaiono come "da verificare". Servizio da aggiungere:
`resolveReference(tabella, nome)` → `{ value: sys_id, displayValue }` (tabelle `core_company`,
`sys_user_group`, `sys_user` e quella delle categorie).

## Da definire

- Flusso FWD (inoltro al reparto interno competente): priorità per il primo rilascio.
