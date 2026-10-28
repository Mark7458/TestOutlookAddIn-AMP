# Area AMP: servizi da collegare

L'area AMP (`amp.js`) gestisce le richieste della casella condivisa e i ticket ServiceNow
direttamente da Outlook. Tutto ciò che sta fuori dal pannello passa da un unico punto
(blocco 3 di `amp.js`): oggi lo fornisce `amp-preview-backend.js` (simulato, solo per
l'anteprima); nell'add-in va fornito dalla Function App, con l'accesso aziendale già usato
per `GetTicketByConversationId`.

Nessun file dell'add-in è stato modificato: l'area usa solo `#amp-view`, `#open-amp-button`
e `TestPluginAPI` (`getCurrentMessageContext`, `onMessageChanged`, `showMainView`).

## Servizi richiesti

Ogni servizio restituisce una Promise. `id` è l'id della richiesta (riga SharePoint).

| Servizio | Cosa fa | Dove agisce |
|---|---|---|
| `listRequests()` | Tutte le richieste aperte o con ticket aperto | SharePoint |
| `getRequestByConversation(conversationId)` | Richiesta della mail aperta, o `null` | SharePoint |
| `registerRequest(mail)` | Aggiunge alla coda una mail non ancora registrata | SharePoint |
| `takeCharge(id, tag)` | Prende in carico (`tag` = es. "Marco") o rilascia (`''`) | SharePoint + categoria Outlook sulla conversazione (Graph) |
| `discardRequest(id, tag)` / `restoreRequest(id)` | Mail che non richiede un ticket / rimessa in coda | SharePoint |
| `createTicket(id, dati)` | Crea l'Incident con i campi indicati e lo assegna a chi lo apre | ServiceNow + SharePoint |
| `addTime(id, minuti, tag)` | Aggiunge (o toglie, se negativo) tempo lavorato | ServiceNow |
| `getTicketActivity(id, { offset, limit })` | Storico del ticket, dal più recente, a pagine | ServiceNow (journal) |
| `setPendingAction(id, azione)` | Salva cosa fare all'invio della risposta (`null` = solo commento) | SharePoint |
| `subscribe(listener)` | Facoltativo: avvisa il pannello quando i dati cambiano | — |

`createTicket` riceve `fields` con: `company`, `assignmentGroup`, `assignedTo`, `category`
(riferimenti `{ value: sys_id, displayValue }`), `impact`, `urgency`, `timeWorkedMinutes`,
`shortDescription` (oggetto della mail + prefisso dell'Attività), `description` (corpo della mail,
max 4000 caratteri), `reportedOn` (data di ricezione, formato `AAAA-MM-GG hh:mm:ss`).

## Lista SharePoint "Richieste AMP" (una riga per conversazione)

| Campo | Contenuto |
|---|---|
| `conversationId` | Chiave: ID conversazione normalizzato (come `GetTicketByConversationId`) |
| `messageId` | Ultima mail della conversazione (per aprirla dal pannello) |
| `subject`, `senderName`, `senderEmail`, `receivedAt`, `preview` | Dati della prima mail |
| `owner` | Tag di chi l'ha in carico (= chi apre il ticket) |
| `status` | `queue` · `discarded` · `ticket` |
| `discardedBy`, `discardedAt` | Se scartata |
| `ticketNumber`, `ticketSysId`, `companyId`, `activityId` | Ticket collegato |
| `pendingAction` | JSON: `{ type: comment / resolve / cancel / hold, resolutionCode, holdReason, holdExpiration, workNotes, by }` |

Stato, tempo lavorato e scadenza On Hold si leggono da ServiceNow (non vanno duplicati).

## All'invio della risposta (lato server)

Quando dalla casella parte una risposta in una conversazione con ticket aperto:

1. il testo della mail va su ServiceNow come *Additional comment*;
2. si applica `pendingAction`:
   - `resolve` / `cancel` → stato Resolved / Canceled, *Resolution code*, *Resolution notes* = testo della mail;
   - `hold` → stato On Hold, *On hold reason*, *On hold expiration*;
   - nessuna azione su un ticket On Hold → torna In Progress;
   - `workNotes`, se presente → *Work notes*;
3. `pendingAction` si azzera.

Proposta: farlo nella Function App osservando la cartella Posta inviata della casella con una
sottoscrizione Microsoft Graph, così funziona anche se l'add-in è chiuso. In alternativa l'evento
`OnMessageSend` dell'add-in (va dichiarato nel manifest e verificato sulle caselle condivise).

## Outlook (Office.js) usato dal pannello

- Aprire una mail della coda: `Office.context.mailbox.displayMessageForm(messageId)`.
- Rispondere con un testo pronto: `Office.context.mailbox.item.displayReplyForm(...)` (solo sulla mail aperta).
- Elenco delle mail e categorie della casella condivisa: non disponibile da Office.js per tutta la
  casella → serve Graph nella Function App (`listRequests` / `takeCharge`).
