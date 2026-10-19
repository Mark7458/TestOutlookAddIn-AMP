(() => {
  "use strict";

  let loadSequence = 0;

  const elementIds = {
    loading: "loading",
    error: "error",
    data: "message-data",
    id: "message-id",
    conversationId: "conversationId",
    categories: "message-categories",
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
  } catch {
    // La risposta non è JSON: viene mostrata così com'è.
  }

  responseElement.textContent = formattedBody;

  resultContainer.hidden = false;
  resultContainer.style.removeProperty("display");
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
  resultContainer.style.removeProperty("display");
}

function registerHealthResultCloseButton() {
  const button =
    document.getElementById("close-function-result");

  if (!button) {
    return;
  }

  button.addEventListener(
    "click",
    resetHealthResult
  );
}

function resetHealthResult() {
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
    resultContainer.style.display = "none";
  }

  setHealthButtonLoading(false);
}

async function testHealthFunction() {
  /*
   * Ripulisce un risultato precedente prima di iniziare
   * una nuova operazione.
   */
  resetHealthResult();
  setHealthButtonLoading(true);

  try {
    if (!window.WorkflowHealth) {
      throw new Error(
        "Il modulo di autenticazione non è stato caricato."
      );
    }

    const result =
      await window.WorkflowHealth.callHealth();

    showHealthResult(result);
  } catch (error) {
    console.error(
      "Errore durante il test Health:",
      error
    );

    /*
     * Mostra sempre l'errore restituito dopo la chiusura
     * del popup MFA.
     */
    showHealthError(error);
  } finally {
    setHealthButtonLoading(false);
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

  function normalizeConversationId(conversationId) {
    if (!conversationId) {
        return "";
    }

    return String(conversationId)
        .replace(/\//g, "-")
        .replace(/\+/g, "_");
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

  function getCategories(item) {
    return new Promise((resolve, reject) => {
        if (
            !item.categories ||
            typeof item.categories.getAsync !== "function"
        ) {
            resolve([]);
            return;
        }

        item.categories.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value || []);
                return;
            }

            const details =
                result.error && result.error.message
                    ? result.error.message
                    : "Office.js non ha restituito le categorie.";

            reject(new Error(details));
        });
    });
  }

  function formatCategories(categories) {
    if (!Array.isArray(categories) || categories.length === 0) {
        return "Nessuna";
    }

    return categories
        .map(category => category.displayName || String(category))
        .filter(Boolean)
        .join(", ");
  }

  function loadDiagnostics() {
    const mailbox = Office.context && Office.context.mailbox;
    const diagnostics = mailbox && mailbox.diagnostics;
    const userProfile = mailbox && mailbox.userProfile;
    const requirements = Office.context && Office.context.requirements;

    setText(
        "diagnostic-host",
        diagnostics && diagnostics.hostName
    );

    setText(
        "diagnostic-host-version",
        diagnostics && diagnostics.hostVersion
    );

    setText(
        "diagnostic-platform",
        Office.context && Office.context.platform
    );

    setText(
        "diagnostic-user",
        userProfile && userProfile.displayName
    );

    setText(
        "diagnostic-email",
        userProfile && userProfile.emailAddress
    );

    setText(
        "diagnostic-account-type",
        userProfile && userProfile.accountType
    );

    setText(
        "diagnostic-timezone",
        userProfile && userProfile.timeZone
    );

    if (requirements) {
        setText(
            "diagnostic-mailbox-18",
            requirements.isSetSupported("Mailbox", "1.8")
        );

        setText(
            "diagnostic-mailbox-113",
            requirements.isSetSupported("Mailbox", "1.13")
        );

        setText(
            "diagnostic-mailbox-114",
            requirements.isSetSupported("Mailbox", "1.14")
        );
    }
  }

  function loadSharedDiagnostics(item) {
    return new Promise((resolve) => {
        if (
            !item ||
            typeof item.getSharedPropertiesAsync !== "function"
        ) {
            setText("diagnostic-shared", "API non disponibile");
            setText("diagnostic-target-mailbox", "");
            setText("diagnostic-owner", "");
            setText("diagnostic-delegate-permissions", "");
            resolve();
            return;
        }

        item.getSharedPropertiesAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const shared = result.value;

                setText("diagnostic-shared", "Sì");
                setText(
                    "diagnostic-target-mailbox",
                    shared.targetMailbox
                );
                setText(
                    "diagnostic-owner",
                    shared.owner
                );
                setText(
                    "diagnostic-delegate-permissions",
                    shared.delegatePermissions
                );

                resolve();
                return;
            }

            setText("diagnostic-shared", "No");
            setText("diagnostic-target-mailbox", "");
            setText("diagnostic-owner", "");
            setText("diagnostic-delegate-permissions", "");

            resolve();
        });
    });
  }

  function loadRoamingOperator() {
    const value =
        Office.context.roamingSettings.get("operatorTest");

    const valueElement =
        document.getElementById("roaming-operator-value");

    const inputElement =
        document.getElementById("roaming-operator-input");

    if (valueElement) {
        valueElement.textContent =
            value || "Non impostato";
    }

    if (inputElement && value) {
        inputElement.value = value;
    }

    console.log(
        "Roaming setting operatorTest:",
        value
    );
  }

  function saveRoamingOperator() {
    const inputElement =
        document.getElementById("roaming-operator-input");

    const statusElement =
        document.getElementById("roaming-operator-status");

    if (!inputElement || !statusElement) {
        return;
    }

    const value = inputElement.value.trim();

    if (!value) {
        statusElement.textContent =
            "Inserisci un valore prima di salvare.";
        return;
    }

    Office.context.roamingSettings.set(
        "operatorTest",
        value
    );

    Office.context.roamingSettings.saveAsync((result) => {
        if (
            result.status ===
            Office.AsyncResultStatus.Succeeded
        ) {
            statusElement.textContent =
                "Valore salvato.";

            loadRoamingOperator();

            console.log(
                "Roaming setting salvato:",
                value
            );
        } else {
            statusElement.textContent =
                "Errore durante il salvataggio: " +
                result.error.message;

            console.error(
                "Errore saveAsync:",
                result.error
            );
        }
    });
  }

  function registerRoamingSettingsControls() {
    const saveButton =
        document.getElementById("roaming-operator-save");

    if (saveButton) {
        saveButton.addEventListener(
            "click",
            saveRoamingOperator
        );
    }

    loadRoamingOperator();
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

      await loadSharedDiagnostics(item);

      const sender = item.from || item.sender;
      const [bodyText, categories] = await Promise.all([
        getBodyAsText(item),
        getCategories(item)
      ]);

      // Ignora risultati asincroni appartenenti al messaggio precedente.
      if (sequence !== loadSequence) {
        return;
      }

      setText(elementIds.id, item.itemId);
      setText(
        elementIds.conversationId,
        normalizeConversationId(item.conversationId)
      );
      setText(elementIds.subject, item.subject, "(Nessun oggetto)");
      setText(elementIds.sender, formatAddress(sender));
      setText(elementIds.received, formatReceivedDate(item.dateTimeCreated));
      setText(elementIds.body, bodyText, "(Corpo vuoto)");
      setText(elementIds.categories, formatCategories(categories),"Nessuna");
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
    registerHealthResultCloseButton();
    loadDiagnostics();
    registerRoamingSettingsControls();
    loadSelectedMessage();
  });
})();
