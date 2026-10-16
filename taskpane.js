(() => {
  "use strict";

  let loadSequence = 0;
  let healthSequence = 0;

  const elementIds = {
    loading: "loading",
    error: "error",
    data: "message-data",
    id: "message-id",
    conversationId: "conversationId",
    subject: "message-subject",
    sender: "message-sender",
    received: "message-received",
    body: "message-body"
  };

  function setHealthButtonLoading(isLoading) {
  const button =
    document.getElementById("test-function-button");

  if (!button) {
    return;
  }

  button.disabled = isLoading;

  button.textContent = isLoading
    ? "Connessione in corso..."
    : "Chiama Health";
}

function showHealthResult(result) {
  const resultContainer =
    document.getElementById("function-result");

  const statusElement =
    document.getElementById("function-status");

  const authModeElement =
    document.getElementById("function-auth-mode");

  const responseElement =
    document.getElementById("function-response");

  if (
    !resultContainer ||
    !statusElement ||
    !authModeElement ||
    !responseElement
  ) {
    return;
  }

  statusElement.textContent =
    `${result.status} ${result.statusText || ""}`.trim();

  statusElement.className =
    result.ok
      ? "status-success"
      : "status-error";

  authModeElement.textContent =
    result.authenticationMode === "silent"
      ? "Token ottenuto silenziosamente"
      : "Login interattivo";

  let formattedBody = result.body;

  try {
    formattedBody = JSON.stringify(
      JSON.parse(result.body),
      null,
      2
    );

    resultContainer.hidden = false;
    resultContainer.style.display = "block";
  } catch {
    // Risposta non JSON: viene mostrata così com'è.
  }

  responseElement.textContent = formattedBody;
  resultContainer.hidden = false;
}

function showHealthError(error) {
  const resultContainer =
    document.getElementById("function-result");

  const statusElement =
    document.getElementById("function-status");

  const authModeElement =
    document.getElementById("function-auth-mode");

  const responseElement =
    document.getElementById("function-response");

  if (
    !resultContainer ||
    !statusElement ||
    !authModeElement ||
    !responseElement
  ) {
    return;
  }

  statusElement.textContent = "Errore";
  statusElement.className = "status-error";

  authModeElement.textContent =
    "Autenticazione o connessione non completata";

  const errorDetails = {
    authenticationPhase:
      error?.authenticationPhase ?? null,

    name:
      error?.name ?? null,

    message:
      error?.message ?? String(error),

    errorCode:
      error?.errorCode ?? null,

    subError:
      error?.subError ?? null,

    correlationId:
      error?.correlationId ?? null,

    timestamp:
      error?.timestamp ?? null,

    silentError:
      error?.silentError ?? null
  };

  responseElement.textContent = JSON.stringify(
    errorDetails,
    null,
    2
  );

  resultContainer.hidden = false;
}

function resetHealthResult() {
  /*
   * Invalida eventuali chiamate Health ancora in corso.
   * Se una vecchia chiamata termina successivamente,
   * non potrà più aggiornare la schermata.
   */
  healthSequence++;

  const resultContainer =
    document.getElementById("function-result");

  const statusElement =
    document.getElementById("function-status");

  const authModeElement =
    document.getElementById("function-auth-mode");

  const responseElement =
    document.getElementById("function-response");

  if (statusElement) {
    statusElement.textContent = "";
    statusElement.className = "";
  }

  if (authModeElement) {
    authModeElement.textContent = "";
  }

  if (responseElement) {
    responseElement.textContent = "";
  }

  if (resultContainer) {
    resultContainer.hidden = true;

    /*
     * Ridondante ma utile contro eventuali regole CSS
     * che non rispettino correttamente hidden.
     */
    resultContainer.style.display = "none";
  }

  setHealthButtonLoading(false);
}

async function testHealthFunction() {
  resetHealthResult();

  const sequence = ++healthSequence;

  setHealthButtonLoading(true);

  try {
    if (!window.WorkflowHealth) {
      throw new Error(
        "Il modulo di autenticazione non è stato caricato."
      );
    }

    const result =
      await window.WorkflowHealth.callHealth();

    /*
     * Nel frattempo è cambiato il messaggio
     * oppure è iniziato un altro test Health.
     */
    if (sequence !== healthSequence) {
      return;
    }

    showHealthResult(result);
  } catch (error) {
    if (sequence !== healthSequence) {
      return;
    }

    console.error(
      "Errore durante il test Health:",
      error
    );

    showHealthError(error);
  } finally {
    if (sequence === healthSequence) {
      setHealthButtonLoading(false);
    }
  }
}

function registerHealthButton() {
  const button =
    document.getElementById("test-function-button");

  if (!button) {
    return;
  }

  button.addEventListener(
    "click",
    testHealthFunction
  );
}

  function byId(id) {
    return document.getElementById(id);
  }

  function setText(id, value, fallback = "Non disponibile") {
    const element = byId(id);
    const text = value === undefined || value === null ? "" : String(value).trim();
    element.textContent = text || fallback;
  }

  function showLoading() {
    byId(elementIds.loading).hidden = false;
    byId(elementIds.error).hidden = true;
    byId(elementIds.data).hidden = true;
  }

  function showError(message) {
    byId(elementIds.loading).hidden = true;
    byId(elementIds.data).hidden = true;
    setText(elementIds.error, message, "Errore imprevisto.");
    byId(elementIds.error).hidden = false;
  }

  function showData() {
    byId(elementIds.loading).hidden = true;
    byId(elementIds.error).hidden = true;
    byId(elementIds.data).hidden = false;
  }

  function formatAddress(address) {
    if (!address) {
      return "";
    }

    const displayName = (address.displayName || "").trim();
    const emailAddress = (address.emailAddress || "").trim();

    if (displayName && emailAddress) {
      return `${displayName} <${emailAddress}>`;
    }

    return displayName || emailAddress;
  }

  function formatReceivedDate(value) {
    if (!value) {
      return "";
    }

    const date = value instanceof Date ? value : new Date(value);
    if (Number.isNaN(date.getTime())) {
      return String(value);
    }

    return new Intl.DateTimeFormat("it-IT", {
      dateStyle: "full",
      timeStyle: "medium"
    }).format(date);
  }

  function getBodyAsText(item) {
    return new Promise((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || "");
          return;
        }

        const details = result.error && result.error.message
          ? result.error.message
          : "Office.js non ha restituito il corpo del messaggio.";
        reject(new Error(details));
      });
    });
  }

  async function loadSelectedMessage() {
    const sequence = ++loadSequence;
    
    resetHealthResult();
    showLoading();

    try {
      const mailbox = Office.context && Office.context.mailbox;
      const item = mailbox && mailbox.item;

      if (!item) {
        throw new Error("Seleziona un messaggio nella visualizzazione di lettura e riapri il task pane.");
      }

      // In Message Read, Office.js espone itemId, subject, from/sender e
      // dateTimeCreated direttamente. Per i messaggi ricevuti, la data è
      // presentata all'utente come data/ora di ricezione.
      const sender = item.from || item.sender;
      const bodyText = await getBodyAsText(item);

      // Ignora risultati asincroni appartenenti al messaggio precedente.
      if (sequence !== loadSequence) {
        return;
      }

      setText(elementIds.id, item.itemId);
      setText(elementIds.conversationId, item.conversationId);
      setText(elementIds.subject, item.subject, "(Nessun oggetto)");
      setText(elementIds.sender, formatAddress(sender));
      setText(elementIds.received, formatReceivedDate(item.dateTimeCreated));
      setText(elementIds.body, bodyText, "(Corpo vuoto)");
      showData();
    } catch (error) {
      if (sequence !== loadSequence) {
        return;
      }

      const message = error instanceof Error ? error.message : String(error);
      showError(message);
    }
  }

  function registerItemChangedHandler() {
    const mailbox = Office.context && Office.context.mailbox;
    if (!mailbox || typeof mailbox.addHandlerAsync !== "function") {
      return;
    }

    mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadSelectedMessage, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.warn("Impossibile registrare ItemChanged:", result.error);
      }
    });
  }

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Outlook) {
      showError("Questa pagina deve essere aperta come task pane di Outlook.");
      return;
    }

    registerItemChangedHandler();
    registerHealthButton();
    loadSelectedMessage();
  });
})();
