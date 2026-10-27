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
    to: "message-to",
    cc: "message-cc",
    received: "message-received",
    body: "message-body",
    attachments: "message-attachments"
  };

  let currentHeaderText = "";
  let currentMessageContext = null;
  const messageContextListeners = new Set();
  const VIEW_PERSISTENCE_KEY = "persistSectionVisibility";
  const SECTION_STATE_KEY = "persistedSectionState";
  const sectionIds = ["email-details", "settings", "diagnostics"];
  function isViewPersistenceEnabled() { return Boolean(Office.context.roamingSettings.get(VIEW_PERSISTENCE_KEY)); }
  function readPersistedSectionState() {
    const value = Office.context.roamingSettings.get(SECTION_STATE_KEY);
    return value && typeof value === "object" ? value : {};
  }
  function collectSectionState() {
    return Object.fromEntries(sectionIds.map((name) => [
      name,
      document.getElementById(`${name}-toggle`)?.getAttribute("aria-expanded") === "true"
    ]));
  }
  function saveRoamingSettings(callback) {
    Office.context.roamingSettings.saveAsync((result) => {
      refreshRoamingSettingsDiagnostics();
      if (typeof callback === "function") callback(result);
    });
  }
  function saveSectionStateIfEnabled() {
    if (!isViewPersistenceEnabled()) return;
    Office.context.roamingSettings.set(SECTION_STATE_KEY, collectSectionState());
    saveRoamingSettings();
  }
  function setSectionOpen(name, open, options = {}) {
    const button = document.getElementById(`${name}-toggle`); const panel = document.getElementById(`${name}-panel`);
    if (!button || !panel) return;
    button.setAttribute("aria-expanded", String(Boolean(open)));
    panel.hidden = !open;
    if (options.persist !== false) saveSectionStateIfEnabled();
  }
  function restorePersistedSectionState() {
    if (!isViewPersistenceEnabled()) {
      sectionIds.forEach((name) => setSectionOpen(name, false, { persist: false }));
      return;
    }
    const state = readPersistedSectionState();
    sectionIds.forEach((name) => setSectionOpen(name, state[name] === true, { persist: false }));
  }
  function resetSectionsIfNeeded() { if (!isViewPersistenceEnabled()) sectionIds.forEach((name) => setSectionOpen(name, false, { persist: false })); }
  function registerSectionToggles() { sectionIds.forEach((name) => document.getElementById(`${name}-toggle`)?.addEventListener("click", () => setSectionOpen(name, document.getElementById(`${name}-toggle`)?.getAttribute("aria-expanded") !== "true"))); }
  function setHealthButtonState(state) {
    const button = document.getElementById("test-function-button");
    const label = document.getElementById("connection-status-label");
    const dot = document.getElementById("connection-status-dot");

    if (!button || !label || !dot) {
      return;
    }

    button.disabled = state === "loading";
    dot.className = `connection-dot connection-dot-${
      state === "connected" ? "success" : state === "error" ? "error" : "unknown"
    }`;
    label.textContent =
      state === "loading"
        ? "Connessione..."
        : state === "connected"
          ? "Connesso"
          : state === "error"
            ? "Errore connessione"
            : "Test Connessione";
  }

  function setConnectionStatus(state, title) {
    const button = document.getElementById("test-function-button");
    if (button && title) {
      button.title = title;
    }
    setHealthButtonState(
      state === "success" ? "connected" : state === "error" ? "error" : "idle"
    );
  }

  function showHealthResult(result) {
    const ok = Boolean(result?.ok);
    setConnectionStatus(
      ok ? "success" : "error",
      ok
        ? `Connessione OK - HTTP ${result.status}`
        : `Connessione non riuscita - HTTP ${result?.status ?? "?"}`
    );
  }

  function showHealthError(error) {
    setConnectionStatus(
      "error",
      `Connessione non riuscita: ${error?.message ?? String(error)}`
    );
  }

  function resetHealthResult() {
    setConnectionStatus("unknown", "Connessione non ancora verificata");
  }

  async function testHealthFunction() {
    resetHealthResult();
    setHealthButtonState("loading");
    refreshAuthDiagnostics();

    try {
      if (!window.WorkflowHealth) {
        throw new Error("Il modulo di autenticazione non è stato caricato.");
      }
      const result = await window.WorkflowHealth.callHealthOnDemand();
      showHealthResult(result);
      refreshAuthDiagnostics();
    } catch (error) {
      console.error("Errore durante il test Health:", error);
      showHealthError(error);
      refreshAuthDiagnostics();
    }
  }

  function formatAuthDiagnosticEntry(entry) {
    const details = entry?.details
      ? ` | ${Object.entries(entry.details)
          .map(([key, value]) => `${key}=${String(value)}`)
          .join(" | ")}`
      : "";
    return `${entry?.time || "--:--:--"} | ${entry?.source || "AUTH"} | ${entry?.event || "event"}${details}`;
  }

  function refreshAuthDiagnostics() {
    const output = document.getElementById("auth-diagnostics-output");
    if (!output) return;
    const workflowHealthAvailable = Boolean(window.WorkflowHealth);
    const diagnosticsApiAvailable = Boolean(
      window.WorkflowHealth?.getDiagnostics
    );
    const events = window.WorkflowHealth?.getDiagnostics?.() || [];
    const lastEvent = events.at(-1)?.event || "nessuno";

    const summary = [
      `WorkflowHealth disponibile: ${workflowHealthAvailable}`,
      `API diagnostica disponibile: ${diagnosticsApiAvailable}`,
      `Eventi locali: ${events.length}`,
      `Ultimo evento: ${lastEvent}`,
      ""
    ].join("\n");

    output.textContent = events.length
      ? summary + events.map(formatAuthDiagnosticEntry).join("\n")
      : summary + "Nessun evento registrato.";
    output.scrollTop = output.scrollHeight;
  }

  async function copyAuthDiagnostics() {
    const output = document.getElementById("auth-diagnostics-output");
    const status = document.getElementById("auth-diagnostics-status");
    if (!output || !status) return;
    const text = output.textContent || "";
    try {
      if (navigator.clipboard?.writeText) await navigator.clipboard.writeText(text);
      else if (!fallbackCopyText(text)) throw new Error("Copia non disponibile.");
      status.textContent = "Diagnostica copiata.";
    } catch {
      status.textContent = fallbackCopyText(text)
        ? "Diagnostica copiata (modalità compatibile)."
        : "Copia non riuscita.";
    }
  }

  function clearAuthDiagnostics() {
    window.WorkflowHealth?.clearDiagnostics?.();
    refreshAuthDiagnostics();
    const status = document.getElementById("auth-diagnostics-status");
    if (status) status.textContent = "Diagnostica pulita.";
  }

  function registerAuthDiagnosticsControls() {
    document.getElementById("auth-diagnostics-refresh")?.addEventListener("click", refreshAuthDiagnostics);
    document.getElementById("auth-diagnostics-copy")?.addEventListener("click", copyAuthDiagnostics);
    document.getElementById("auth-diagnostics-clear")?.addEventListener("click", clearAuthDiagnostics);
    window.addEventListener("workflow-auth-diagnostics-changed", refreshAuthDiagnostics);
    refreshAuthDiagnostics();
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

  function formatAddressList(addresses) { return (Array.isArray(addresses) ? addresses : []).map(formatAddress).filter(Boolean).join("; ") || "Nessuno"; }
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


  function formatAttachmentSize(size) {
    const bytes = Number(size);

    if (!Number.isFinite(bytes) || bytes < 0) {
      return "Dimensione non disponibile";
    }

    if (bytes < 1024) {
      return `${bytes} B`;
    }

    if (bytes < 1024 * 1024) {
      return `${(bytes / 1024).toFixed(1)} KB`;
    }

    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  function isMessageAttachment(attachment) {
    const type = String(
      attachment?.attachmentType || ""
    ).toLowerCase();

    const itemType = String(
      Office.MailboxEnums?.AttachmentType?.Item || "item"
    ).toLowerCase();

    return type === itemType || type === "item";
  }

  function isEmlFileAttachment(attachment) {
    const name = String(attachment?.name || "").trim().toLowerCase();
    return name.endsWith(".eml");
  }

  function decodeBase64Text(base64Content) {
    const binary = atob(String(base64Content || ""));
    const bytes = Uint8Array.from(binary, char => char.charCodeAt(0));
    try {
      return new TextDecoder("utf-8", { fatal: false }).decode(bytes);
    } catch {
      return binary;
    }
  }

  function getAttachmentContent(item, attachmentId) {
    return new Promise((resolve, reject) => {
      if (
        !item ||
        typeof item.getAttachmentContentAsync !== "function"
      ) {
        reject(
          new Error(
            "Il client Outlook non espone il contenuto degli allegati."
          )
        );
        return;
      }

      item.getAttachmentContentAsync(
        attachmentId,
        (result) => {
          if (
            result.status ===
            Office.AsyncResultStatus.Succeeded
          ) {
            resolve(result.value);
            return;
          }

          const details =
            result.error && result.error.message
              ? result.error.message
              : "Office.js non ha restituito il contenuto dell'allegato.";

          reject(new Error(details));
        }
      );
    });
  }

  function extractMessageHeader(emlContent) {
    const content = String(emlContent || "");
    const separator = /\r?\n\r?\n/.exec(content);

    if (!separator) {
      return content.trim();
    }

    return content
      .slice(0, separator.index)
      .trim();
  }

  function setHeaderStatus(message, isError = false) {
    const status = byId("header-status");

    if (!status) {
      return;
    }

    status.textContent = message || "";
    status.className = isError
      ? "header-status status-error"
      : "header-status";
  }

  function cloneContext(value) {
    return value == null ? null : JSON.parse(JSON.stringify(value));
  }
  function publishMessageContext() {
    const snapshot = cloneContext(currentMessageContext);
    messageContextListeners.forEach((listener) => {
      try { listener(snapshot); } catch (error) { console.error("Listener messaggio non riuscito:", error); }
    });
  }
  function getCurrentMessageContext() { return cloneContext(currentMessageContext); }
  function onMessageChanged(listener) {
    if (typeof listener !== "function") throw new TypeError("Il listener deve essere una funzione.");
    messageContextListeners.add(listener);
    return () => messageContextListeners.delete(listener);
  }

  function showMainView() {
    const mainView = byId("main-view");
    const headerView = byId("header-view");
    const ampView = byId("amp-view");

    if (mainView) {
      mainView.hidden = false;
    }

    if (headerView) {
      headerView.hidden = true;
    }
    if (ampView) {
      ampView.hidden = true;
    }
  }

  function showHeaderView() {
    const mainView = byId("main-view");
    const headerView = byId("header-view");

    if (mainView) {
      mainView.hidden = true;
    }

    if (headerView) {
      headerView.hidden = false;
    }
  }

  async function openAttachmentHeader(attachment) {
    const item =
      Office.context &&
      Office.context.mailbox &&
      Office.context.mailbox.item;

    const contentElement =
      byId("attachment-header-content");

    const nameElement =
      byId("header-attachment-name");

    currentHeaderText = "";

    if (nameElement) {
      nameElement.textContent =
        attachment.name || "Messaggio allegato";
    }

    if (contentElement) {
      contentElement.textContent = "";
    }

    setHeaderStatus("Recupero dell'header in corso...");
    showHeaderView();

    try {
      const attachmentContent =
        await getAttachmentContent(
          item,
          attachment.id
        );

      const format = String(
        attachmentContent?.format || ""
      ).toLowerCase();

      const emlFormat = String(
        Office.MailboxEnums
          ?.AttachmentContentFormat
          ?.Eml || "eml"
      ).toLowerCase();

      const base64Format = String(
        Office.MailboxEnums?.AttachmentContentFormat?.Base64 || "base64"
      ).toLowerCase();
      let emlContent;

      if (format === emlFormat || format === "eml") {
        emlContent = attachmentContent.content;
      } else if (
        isEmlFileAttachment(attachment) &&
        (format === base64Format || format === "base64")
      ) {
        emlContent = decodeBase64Text(attachmentContent.content);
      } else {
        throw new Error(
          "L'allegato non è stato restituito come messaggio EML leggibile."
        );
      }

      currentHeaderText = extractMessageHeader(emlContent);

      if (!currentHeaderText) {
        throw new Error(
          "L'EML non contiene un header leggibile."
        );
      }

      contentElement.textContent = currentHeaderText;
      setHeaderStatus("Header recuperato.");
    } catch (error) {
      const message =
        error instanceof Error
          ? error.message
          : String(error);

      if (contentElement) {
        contentElement.textContent = "";
      }

      setHeaderStatus(
        `Impossibile recuperare l'header: ${message}`,
        true
      );
    }
  }

  function getAttachmentExtension(name) {
    const match = String(name || "").trim().match(/\.([^.]+)$/);
    return match ? match[1].toLowerCase() : "";
  }

  function getAttachmentMimeType(attachment) {
    const knownTypes = {
      pdf: "application/pdf", png: "image/png", jpg: "image/jpeg",
      jpeg: "image/jpeg", gif: "image/gif", webp: "image/webp",
      txt: "text/plain", csv: "text/csv", html: "text/html",
      htm: "text/html", eml: "message/rfc822"
    };
    return knownTypes[getAttachmentExtension(attachment?.name)] || "application/octet-stream";
  }

  function base64ToBlob(base64Content, mimeType) {
    const binary = atob(String(base64Content || ""));
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return new Blob([bytes], { type: mimeType });
  }

  function triggerAttachmentDownload(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName || "allegato";
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  async function downloadAttachment(attachment) {
    try {
      const cloudType = String(Office.MailboxEnums?.AttachmentType?.Cloud || "cloud").toLowerCase();
      if (String(attachment?.attachmentType || "").toLowerCase() === cloudType) {
        window.open(attachment.id, "_blank", "noopener,noreferrer");
        return;
      }
      const content = await getAttachmentContent(Office.context?.mailbox?.item, attachment.id);
      const format = String(content?.format || "").toLowerCase();
      let blob;
      if (format === "base64") {
        blob = base64ToBlob(content.content, getAttachmentMimeType(attachment));
      } else if (format === "eml") {
        blob = new Blob([content.content], { type: "message/rfc822" });
      } else {
        throw new Error(`Formato allegato non gestito: ${content?.format || "sconosciuto"}.`);
      }
      triggerAttachmentDownload(blob, attachment.name || "allegato");
    } catch (error) {
      console.error("Download allegato non riuscito:", error);
      window.alert(`Impossibile scaricare l'allegato: ${error?.message || String(error)}`);
    }
  }

  function renderAttachments(attachments) {
    const container = byId(elementIds.attachments);

    if (!container) {
      return;
    }

    container.replaceChildren();

    if (
      !Array.isArray(attachments) ||
      attachments.length === 0
    ) {
      const empty = document.createElement("div");
      empty.className = "attachments-empty";
      empty.textContent = "Nessuno";
      container.appendChild(empty);
      return;
    }

    attachments.forEach((attachment) => {
      const card = document.createElement("div");
      card.className = "attachment-card";

      const summary = document.createElement("div");
      summary.className = "attachment-summary";

      const name = document.createElement("div");
      name.className = "attachment-name";
      name.textContent =
        attachment.name || "Allegato senza nome";

      const meta = document.createElement("div");
      meta.className = "attachment-meta";

      const messageAttachment = isMessageAttachment(attachment);
      const emlFileAttachment = isEmlFileAttachment(attachment);
      const canShowHeader = messageAttachment || emlFileAttachment;

      let attachmentLabel = "File";
      if (messageAttachment) {
        attachmentLabel = "Elemento Outlook";
      } else if (emlFileAttachment) {
        attachmentLabel = "Messaggio EML";
      }

      meta.textContent = [
        attachmentLabel,
        formatAttachmentSize(attachment.size)
      ].join(" · ");

      summary.append(name, meta);
      card.appendChild(summary);

      const actions = document.createElement("div");
      actions.className = "attachment-actions";
      if (canShowHeader) {
        const headerAction = document.createElement("button");
        headerAction.type = "button";
        headerAction.className = "attachment-action-button";
        headerAction.textContent = "Mostra header";
        headerAction.addEventListener("click", () => openAttachmentHeader(attachment));
        actions.appendChild(headerAction);
      }
      const downloadAction = document.createElement("button");
      downloadAction.type = "button";
      downloadAction.className = "attachment-action-button";
      downloadAction.textContent = "Scarica";
      downloadAction.addEventListener("click", () => downloadAttachment(attachment));
      actions.appendChild(downloadAction);
      card.appendChild(actions);
      container.appendChild(card);
    });
  }

  function fallbackCopyText(text) {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.setAttribute("readonly", "");
    textarea.style.position = "fixed";
    textarea.style.left = "-9999px";
    textarea.style.top = "0";
    document.body.appendChild(textarea);
    textarea.focus();
    textarea.select();
    textarea.setSelectionRange(0, textarea.value.length);
    let copied = false;
    try { copied = document.execCommand("copy"); }
    finally { textarea.remove(); }
    return copied;
  }

  async function copyCurrentHeader() {
    if (!currentHeaderText) {
      setHeaderStatus("Nessun header disponibile da copiare.", true);
      return;
    }
    if (navigator.clipboard && typeof navigator.clipboard.writeText === "function") {
      try {
        await navigator.clipboard.writeText(currentHeaderText);
        setHeaderStatus("Header copiato negli appunti.");
        return;
      } catch (error) {
        console.warn("Clipboard API bloccata; provo il fallback execCommand.", error);
      }
    }
    try {
      if (!fallbackCopyText(currentHeaderText)) {
        throw new Error("Il client non consente la copia automatica.");
      }
      setHeaderStatus("Header copiato negli appunti (modalità compatibile).");
    } catch (error) {
      setHeaderStatus(`Copia non riuscita: ${error?.message || String(error)}`, true);
    }
  }

  function registerHeaderViewControls() {
    const backButton =
      byId("header-back-button");

    const copyButton =
      byId("copy-header-button");

    if (backButton) {
      backButton.addEventListener(
        "click",
        showMainView
      );
    }

    if (copyButton) {
      copyButton.addEventListener(
        "click",
        copyCurrentHeader
      );
    }
  }

  function refreshRoamingSettingsDiagnostics() {
    const values = {
      [VIEW_PERSISTENCE_KEY]: Office.context.roamingSettings.get(VIEW_PERSISTENCE_KEY),
      [SECTION_STATE_KEY]: Office.context.roamingSettings.get(SECTION_STATE_KEY)
    };
    setText("diagnostic-roaming-settings", JSON.stringify(values, null, 2));
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
      const versions = ["1.14", "1.13", "1.12", "1.11", "1.10", "1.9", "1.8", "1.7", "1.6", "1.5", "1.4", "1.3", "1.2", "1.1"];
      setText("diagnostic-mailbox-version", versions.find((version) => requirements.isSetSupported("Mailbox", version)) || "Non rilevata");
    }
  }

  function loadSharedDiagnostics(item) {
    return new Promise((resolve) => {
      if (!item || typeof item.getSharedPropertiesAsync !== "function") {
        setText("diagnostic-shared", "API non disponibile");
        setText("diagnostic-target-mailbox", "");
        setText("diagnostic-owner", "");
        setText("diagnostic-delegate-permissions", "");
        resolve(null);
        return;
      }
      item.getSharedPropertiesAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const shared = result.value;
          setText("diagnostic-shared", "Sì");
          setText("diagnostic-target-mailbox", shared.targetMailbox);
          setText("diagnostic-owner", shared.owner);
          setText("diagnostic-delegate-permissions", shared.delegatePermissions);
          resolve({
            targetMailbox: shared.targetMailbox || null,
            owner: shared.owner || null,
            delegatePermissions: shared.delegatePermissions || null
          });
          return;
        }
        setText("diagnostic-shared", "No");
        setText("diagnostic-target-mailbox", "");
        setText("diagnostic-owner", "");
        setText("diagnostic-delegate-permissions", "");
        resolve(null);
      });
    });
  }
  function loadPluginSettings() {
    const checkbox = document.getElementById("persist-view-setting");
    if (checkbox) checkbox.checked = isViewPersistenceEnabled();
    restorePersistedSectionState();
    refreshRoamingSettingsDiagnostics();
  }
  function savePluginSettings() {
    const checkbox = document.getElementById("persist-view-setting"); const status = document.getElementById("settings-status");
    const enabled = Boolean(checkbox?.checked);
    Office.context.roamingSettings.set(VIEW_PERSISTENCE_KEY, enabled);
    if (enabled) Office.context.roamingSettings.set(SECTION_STATE_KEY, collectSectionState());
    else Office.context.roamingSettings.remove(SECTION_STATE_KEY);
    saveRoamingSettings((result) => {
      if (status) status.textContent = result.status === Office.AsyncResultStatus.Succeeded ? "Impostazione salvata." : `Errore: ${result.error?.message || "salvataggio non riuscito"}`;
    });
  }
  function registerPluginSettingsControls() { document.getElementById("persist-view-setting")?.addEventListener("change", savePluginSettings); loadPluginSettings(); }
  const GET_TICKET_BY_CONVERSATION_ID_URL =
    "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/GetTicketByConversationId";

  function resetTicketResult() {
    const panel = byId("ticket-result");
    if (panel) {
      panel.hidden = true;
    }
    setText("ticket-status", "", "");
    setText("ticket-auth-mode", "", "");
    setText("ticket-response", "", "");
  }

  function setTicketButtonLoading(isLoading) {
    const button = byId("test-ticket-button");
    if (!button) {
      return;
    }
    button.disabled = isLoading;
    button.textContent = isLoading ? "Lookup in corso..." : "Test Lookup Ticket";
  }

  function showTicketResult(result) {
    const panel = byId("ticket-result");
    const status = byId("ticket-status");
    const authMode = byId("ticket-auth-mode");
    const response = byId("ticket-response");
    if (!panel || !status || !authMode || !response) {
      return;
    }

    panel.hidden = false;
    status.textContent = `${result.status} ${result.statusText || ""}`.trim();
    status.className = result.ok ? "status-success" : "status-error";
    authMode.textContent = result.authenticationMode || "Non disponibile";

    try {
      response.textContent = JSON.stringify(JSON.parse(result.body || "{}"), null, 2);
    } catch {
      response.textContent = result.body || "(Risposta vuota)";
    }
  }

  function showTicketError(error) {
    const panel = byId("ticket-result");
    const status = byId("ticket-status");
    const authMode = byId("ticket-auth-mode");
    const response = byId("ticket-response");
    if (!panel || !status || !authMode || !response) {
      return;
    }

    panel.hidden = false;
    status.textContent = "Errore";
    status.className = "status-error";
    authMode.textContent = "Non disponibile";
    response.textContent = error?.message || String(error);
  }

  async function testTicketLookup() {
    resetTicketResult();
    setTicketButtonLoading(true);

    try {
      const conversationId = currentMessageContext?.conversationId;
      if (!conversationId) {
        throw new Error("Conversation ID non disponibile per il messaggio selezionato.");
      }
      if (!window.WorkflowHealth?.getAccessTokenSilently) {
        throw new Error("API autenticazione silent non disponibile.");
      }

      const authentication = await window.WorkflowHealth.getAccessTokenSilently();
      if (!authentication?.accessToken) {
        throw new Error("Token di accesso non disponibile.");
      }

      const url = `${GET_TICKET_BY_CONVERSATION_ID_URL}?conversationId=${encodeURIComponent(conversationId)}`;
      const startedAt = Date.now();
      const fetchResponse = await fetch(url, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${authentication.accessToken}`,
          Accept: "application/json"
        }
      });
      const body = await fetchResponse.text();
      const result = {
        ok: fetchResponse.ok,
        status: fetchResponse.status,
        statusText: fetchResponse.statusText,
        authenticationMode: authentication.acquisitionMode,
        body
      };

      console.info("GetTicketByConversationId:", {
        status: result.status,
        ok: result.ok,
        authenticationMode: result.authenticationMode,
        elapsedMs: Date.now() - startedAt,
        body: result.body
      });
      showTicketResult(result);
    } catch (error) {
      console.error("Test GetTicketByConversationId non riuscito:", error);
      showTicketError(error);
    } finally {
      setTicketButtonLoading(false);
    }
  }

  function registerTicketTestControls() {
    byId("test-ticket-button")?.addEventListener("click", testTicketLookup);
    byId("close-ticket-result")?.addEventListener("click", resetTicketResult);
  }

  async function loadSelectedMessage() {
    const sequence = ++loadSequence;

    currentHeaderText = "";
    resetSectionsIfNeeded();
    showMainView();
    resetTicketResult();
    showLoading();

    try {
      const mailbox = Office.context && Office.context.mailbox;
      const item = mailbox && mailbox.item;

      if (!item) {
        throw new Error("Seleziona un messaggio nella visualizzazione di lettura e riapri il task pane.");
      }

      const sharedProperties = await loadSharedDiagnostics(item);

      const sender = item.from || item.sender;
      const attachments = item.attachments
        ? Array.from(item.attachments)
        : [];
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
      setText(elementIds.to, formatAddressList(item.to));
      setText(elementIds.cc, formatAddressList(item.cc));
      setText(elementIds.received, formatReceivedDate(item.dateTimeCreated));
      setText(elementIds.body, bodyText, "(Corpo vuoto)");
      setText(elementIds.categories, formatCategories(categories),"Nessuna");
      currentMessageContext = {
        itemId: item.itemId || null,
        conversationId: item.conversationId || null,
        normalizedConversationId: normalizeConversationId(item.conversationId),
        subject: item.subject || "",
        sender: sender ? {
          displayName: sender.displayName || "",
          emailAddress: sender.emailAddress || ""
        } : null,
        to: (Array.isArray(item.to) ? item.to : []).map((address) => ({
          displayName: address.displayName || "",
          emailAddress: address.emailAddress || ""
        })),
        cc: (Array.isArray(item.cc) ? item.cc : []).map((address) => ({
          displayName: address.displayName || "",
          emailAddress: address.emailAddress || ""
        })),
        dateTimeCreated: item.dateTimeCreated instanceof Date
          ? item.dateTimeCreated.toISOString()
          : (item.dateTimeCreated || null),
        bodyText,
        categories: (Array.isArray(categories) ? categories : []).map((category) => ({
          displayName: category.displayName || String(category),
          color: category.color || null
        })),
        attachments: attachments.map((attachment) => ({
          id: attachment.id || null,
          name: attachment.name || "",
          size: Number.isFinite(Number(attachment.size)) ? Number(attachment.size) : null,
          attachmentType: attachment.attachmentType || null,
          isInline: Boolean(attachment.isInline),
          contentType: attachment.contentType || null
        })),
        sharedProperties,
        mailbox: {
          userDisplayName: mailbox.userProfile?.displayName || null,
          userEmailAddress: mailbox.userProfile?.emailAddress || null,
          accountType: mailbox.userProfile?.accountType || null,
          timeZone: mailbox.userProfile?.timeZone || null
        }
      };
      renderAttachments(attachments);
      showData();
      publishMessageContext();
    } catch (error) {
      if (sequence !== loadSequence) {
        return;
      }

      currentMessageContext = null;
      publishMessageContext();
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

  window.TestPluginAPI = Object.freeze({
    version: "1.0.0",
    getCurrentMessageContext,
    onMessageChanged,
    showMainView
  });

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Outlook) {
      showError("Questa pagina deve essere aperta come task pane di Outlook.");
      return;
    }

    registerItemChangedHandler();
    registerSectionToggles();
    registerHealthButton();
    registerTicketTestControls();
    registerAuthDiagnosticsControls();
    loadDiagnostics();
    registerPluginSettingsControls();
    registerHeaderViewControls();
    window.AmpModule?.initialize?.().catch((error) => console.error("AMP init error:", error));
    // UI subito disponibile; autenticazione solo alla prima chiamata che ne ha bisogno.
    loadSelectedMessage();
    refreshAuthDiagnostics();
  });
})();
