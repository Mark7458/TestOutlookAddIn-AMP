// Test-Plugin file version: 0.2-taskpane.js-devtools-10
const TEST_PLUGIN_BUILD = Object.freeze({
  application: '0.2',
  release: '0.2-devtools',
  files: Object.freeze({
    'taskpane.html': '0.2-taskpane.html-devtools-10',
    'taskpane.js': '0.2-taskpane.js-devtools-10',
    'taskpane.css': '0.2-taskpane.css-devtools-10',
    'auth-src.js': '0.2-auth-src.js-3',
    'auth-dialog.html': '0.2-auth-dialog.html-2',
    'auth-dialog-src.js': '0.2-auth-dialog-src.js-3',
    'auth-popup-redirect.html': '0.2-auth-popup-redirect.html-1',
    'auth-popup-redirect-src.js': '0.2-auth-popup-redirect-src.js-1',
  }),
  database: Object.freeze({ schemaVersion: 4, requiredDataVersion: 5 }),
  generated: Object.freeze(["auth.bundle.js", "auth-dialog.bundle.js", "auth-popup-redirect.bundle.js"])
});
window.TestPluginBuild = TEST_PLUGIN_BUILD;
(() => {
  "use strict";

  let loadSequence = 0;

  const elementIds = {
    loading: "loading",
    error: "error",
    data: "message-data",
    id: "message-id",
    idToggle: "message-id-toggle",
    conversationId: "conversationId",
    conversationIdToggle: "conversation-id-toggle",
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
  let currentMessageIdentifiers = null;
  const identifierViewState = {
    itemId: "normalized",
    conversationId: "normalized"
  };
  let currentSharePointData = null;
  let developerToolsAllowed = false;
  let developerLastItemId = "";
  let developerResultView = "fields";
  let developerConfig = null;
  let developerCurrentOperatorId = "";
  let developerLastCreateBody = null;
  const messageContextListeners = new Set();
  const sharePointDataListeners = new Set();
  const UI_PREFERENCES_KEY = "test-plugin-ui-preferences";
  const sectionIds = ["email-details", "settings", "diagnostics"];
  function readUiPreferences() {
    try {
      const parsed = JSON.parse(localStorage.getItem(UI_PREFERENCES_KEY) || "{}");
      return parsed && typeof parsed === "object" ? parsed : {};
    } catch { return {}; }
  }
  function writeUiPreferences(preferences) {
    localStorage.setItem(UI_PREFERENCES_KEY, JSON.stringify(preferences));
    refreshRoamingSettingsDiagnostics();
  }
  function isDiagnosticsEnabled() { return Boolean(readUiPreferences().diagnosticsEnabled); }
  function applyDiagnosticsVisibility() {
    const enabled = isDiagnosticsEnabled();
    document.querySelectorAll("[data-diagnostics-log-ui]").forEach((element) => { element.hidden = !enabled; });
  }
  function isViewPersistenceEnabled() { return Boolean(readUiPreferences().persistSectionVisibility); }
  function readPersistedSectionState() {
    const value = readUiPreferences().persistedSectionState;
    return value && typeof value === "object" ? value : {};
  }
  function collectSectionState() {
    return Object.fromEntries(sectionIds.map((name) => [
      name,
      document.getElementById(`${name}-toggle`)?.getAttribute("aria-expanded") === "true"
    ]));
  }
  function saveSectionStateIfEnabled() {
    if (!isViewPersistenceEnabled()) return;
    const preferences = readUiPreferences();
    preferences.persistedSectionState = collectSectionState();
    writeUiPreferences(preferences);
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

    const build = window.TestPluginBuild;
    const componentLines = build?.files ? Object.entries(build.files).map(([name, version]) => `- ${name}: ${version}`) : ["- Non disponibili"];
    const summary = [`Test-Plugin: ${build?.application || "Non disponibile"}`, `Release: ${build?.release || "Non disponibile"}`, "Componenti:", ...componentLines, "", `WorkflowHealth disponibile: ${workflowHealthAvailable}`, `API diagnostica disponibile: ${diagnosticsApiAvailable}`, `Eventi locali: ${events.length}`, `Ultimo evento: ${lastEvent}`, ""].join("\n");
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

  async function resetAuthenticationCompletelyFromDiagnostics() {
  const status = document.getElementById("auth-reset-status");
  const button = document.getElementById("auth-full-reset");
  if (!window.WorkflowHealth?.resetAuthenticationCompletely) {
    if (status) status.textContent = "API reset completo non disponibile.";
    return;
  }
  if (button) button.disabled = true;
  if (status) status.textContent = "Reset completo auth in corso...";
  try {
    const result = await window.WorkflowHealth.resetAuthenticationCompletely();
    if (status) status.textContent = `Reset completo auth eseguito (${result?.removedCount || 0} chiavi dialog rimosse).`;
    resetHealthResult();
    refreshAuthDiagnostics();
  } catch (error) {
    if (status) status.textContent = `Reset completo auth non riuscito: ${error?.message || String(error)}`;
  } finally {
    if (button) button.disabled = false;
  }
}
function registerAuthDiagnosticsControls() {
    document.getElementById("auth-diagnostics-refresh")?.addEventListener("click", refreshAuthDiagnostics);
    document.getElementById("auth-diagnostics-copy")?.addEventListener("click", copyAuthDiagnostics);
    document.getElementById("auth-diagnostics-clear")?.addEventListener("click", clearAuthDiagnostics);
      document.getElementById("auth-full-reset")?.addEventListener("click", resetAuthenticationCompletelyFromDiagnostics);
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

  function convertItemIdToRestV2(originalItemId) {
    if (!originalItemId) return "";
    if (typeof Office.context?.mailbox?.convertToRestId !== "function") {
      throw new Error("Office.js non rende disponibile la conversione dell'ID messaggio in formato REST v2.0.");
    }
    return Office.context.mailbox.convertToRestId(originalItemId, Office.MailboxEnums.RestVersion.v2_0);
  }

  function updateIdentifierField(kind) {
    const isItemId = kind === "itemId";
    const fieldId = isItemId ? elementIds.id : elementIds.conversationId;
    const toggleId = isItemId ? elementIds.idToggle : elementIds.conversationIdToggle;
    const originalKey = isItemId ? "originalItemId" : "originalConversationId";
    const normalizedKey = isItemId ? "normalizedItemId" : "normalizedConversationId";
    const showOriginal = identifierViewState[kind] === "original";
    const value = currentMessageIdentifiers?.[showOriginal ? originalKey : normalizedKey] || "";
    const button = byId(toggleId);

    setText(fieldId, value);
    if (!button) return;

    button.textContent = showOriginal ? "Mostra normalizzato" : "Mostra originale";
    button.setAttribute("aria-pressed", String(showOriginal));
    button.setAttribute(
      "aria-label",
      `${showOriginal ? "Mostra il valore normalizzato" : "Mostra il valore originale"} per ${isItemId ? "ID del messaggio" : "Conversation ID"}`
    );
    button.disabled = !currentMessageIdentifiers?.[originalKey] || !currentMessageIdentifiers?.[normalizedKey];
  }

  function renderIdentifierFields() {
    updateIdentifierField("itemId");
    updateIdentifierField("conversationId");
  }

  function resetIdentifierViews() {
    identifierViewState.itemId = "normalized";
    identifierViewState.conversationId = "normalized";
    renderIdentifierFields();
  }

  function toggleIdentifierView(kind) {
    if (!currentMessageIdentifiers) return;
    identifierViewState[kind] = identifierViewState[kind] === "normalized" ? "original" : "normalized";
    updateIdentifierField(kind);
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

  function categoryDisplayNames(categories) {
    return (Array.isArray(categories) ? categories : [])
      .map((category) => String(category?.displayName || category || "").trim())
      .filter(Boolean);
  }

  function mutateCategories(item, method, categoryNames) {
    return new Promise((resolve, reject) => {
      const categoriesApi = item?.categories;
      if (!categoriesApi || typeof categoriesApi[method] !== "function") {
        reject(new Error(`Il client Outlook non espone categories.${method}.`));
        return;
      }

      const values = Array.isArray(categoryNames) ? categoryNames : [];
      if (!values.length) {
        resolve();
        return;
      }

      categoriesApi[method](values, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }

        const details = result.error?.message || `Office.js non ha completato categories.${method}.`;
        reject(new Error(details));
      });
    });
  }

  function addCategories(item, categoryNames) {
    return mutateCategories(item, "addAsync", categoryNames);
  }

  function removeCategories(item, categoryNames) {
    return mutateCategories(item, "removeAsync", categoryNames);
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
  function updateCurrentMessageCategories(categories) {
    const normalized = (Array.isArray(categories) ? categories : []).map((category) => ({
      displayName: category?.displayName || String(category),
      color: category?.color || null
    }));
    setText(elementIds.categories, formatCategories(normalized), "Nessuna");
    if (currentMessageContext) currentMessageContext.categories = normalized;
    publishMessageContext();
    return normalized;
  }
  function getCurrentMessageContext() { return cloneContext(currentMessageContext); }
  function onMessageChanged(listener) {
    if (typeof listener !== "function") throw new TypeError("Il listener deve essere una funzione.");
    messageContextListeners.add(listener);
    return () => messageContextListeners.delete(listener);
  }
  function publishSharePointData() {
    const snapshot = cloneContext(currentSharePointData);
    sharePointDataListeners.forEach((listener) => {
      try { listener(snapshot); } catch (error) { console.error("Listener SharePoint non riuscito:", error); }
    });
  }
  function getCurrentSharePointData() { return cloneContext(currentSharePointData); }
  function onSharePointDataChanged(listener) {
    if (typeof listener !== "function") throw new TypeError("Il listener deve essere una funzione.");
    sharePointDataListeners.add(listener);
    return () => sharePointDataListeners.delete(listener);
  }

  function showMainView() {
    const mainView = byId("main-view");
    const headerView = byId("header-view");

    if (mainView) {
      mainView.hidden = false;
    }

    if (headerView) {
      headerView.hidden = true;
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

  function syncDeveloperConversationId() {
    if (getDeveloperAction() !== "GetRecordByConversationId") return;
    const input = byId("developer-action-input");
    if (input) input.value = currentMessageContext?.normalizedConversationId || "";
  }
  function showDeveloperToolsView() {
    if (!developerToolsAllowed) return;
    showMainView();
    setSectionOpen("email-details", true);
    syncDeveloperConversationId();
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
    setText("diagnostic-roaming-settings", JSON.stringify({
      storage: "localStorage",
      key: UI_PREFERENCES_KEY,
      ...readUiPreferences()
    }, null, 2));
  }

  function loadBuildVersions() {
    const build = window.TestPluginBuild;
    setText("plugin-version", build?.application ? `v.${build.application}` : "", "Versione non disponibile");
    const lines = build?.files ? [`Applicazione: ${build.application}`, `Release: ${build.release}`, ...Object.entries(build.files).map(([name, version]) => `${name}: ${version}`), "Bundle generati:", ...(build.generated || []).map((name) => `${name}: generato`)] : [];
    setText("diagnostic-component-versions", lines.join("\n"));
    window.WorkflowHealth?.recordDiagnostic?.("component_versions_loaded", { application: build?.application || null, release: build?.release || null, files: build?.files ? JSON.stringify(build.files) : null });
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
    const diagnosticsCheckbox = document.getElementById("diagnostics-enabled-setting");
    if (checkbox) checkbox.checked = isViewPersistenceEnabled();
    if (diagnosticsCheckbox) diagnosticsCheckbox.checked = isDiagnosticsEnabled();
    restorePersistedSectionState();
    applyDiagnosticsVisibility();
    refreshRoamingSettingsDiagnostics();
  }
  function savePluginSettings() {
    const checkbox = document.getElementById("persist-view-setting"); const status = document.getElementById("settings-status");
    const enabled = Boolean(checkbox?.checked);
    const diagnosticsEnabled = Boolean(document.getElementById("diagnostics-enabled-setting")?.checked);
    try {
      const preferences = readUiPreferences();
      preferences.persistSectionVisibility = enabled;
      preferences.diagnosticsEnabled = diagnosticsEnabled;
      if (enabled) preferences.persistedSectionState = collectSectionState();
      else delete preferences.persistedSectionState;
      writeUiPreferences(preferences);
      applyDiagnosticsVisibility();
      if (status) status.textContent = "Impostazione salvata.";
    } catch (error) {
      if (status) status.textContent = `Errore: ${error?.message || "salvataggio non riuscito"}`;
    }
  }
  function registerPluginSettingsControls() { document.getElementById("persist-view-setting")?.addEventListener("change", savePluginSettings); document.getElementById("diagnostics-enabled-setting")?.addEventListener("change", savePluginSettings); loadPluginSettings(); }
  const OPERATIONAL_CONFIG_URL = "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/PluginConfig";
  const OPERATIONAL_CONFIG_KEY = "test-plugin-operational-config";
  function readOperationalConfigCache() {
    try { return JSON.parse(localStorage.getItem(OPERATIONAL_CONFIG_KEY) || "null"); }
    catch (error) { window.WorkflowHealth?.recordDiagnostic?.("config_cache_invalid", { message: error?.message || String(error) }); return null; }
  }
  function collectConfiguredCategories(payload) {
    const candidates = [
      ...Object.values(payload?.operators || {}).map((operator) => operator?.tag),
      ...Object.values(payload?.groups || {}).map((group) => group?.tag),
      ...(Array.isArray(payload?.categories?.generic)
        ? payload.categories.generic.map((category) => category?.tag)
        : [])
    ];
    const unique = new Map();
    candidates.forEach((value) => {
      const tag = String(value || "").trim();
      const key = tag.toLocaleLowerCase("it-IT");
      if (tag && !unique.has(key)) unique.set(key, tag);
    });
    return Array.from(unique.values());
  }
  function refreshOperationalConfigDiagnostics(cache = readOperationalConfigCache(), status = null) {
    const payload = cache?.payload;
    const lines = [
      `Stato: ${status || (payload ? "disponibile" : "non disponibile")}`,
      `Schema richiesto: ${TEST_PLUGIN_BUILD.database.schemaVersion}`,
      `Versione richiesta: ${TEST_PLUGIN_BUILD.database.requiredDataVersion}`,
      `Schema disponibile: ${cache?.schemaVersion ?? "-"}`,
      `Versione disponibile: ${cache?.dataVersion ?? "-"}`,
      `Scaricata: ${cache?.downloadedAt || "-"}`,
      `Agenzie: ${Object.keys(payload?.agencies || {}).length}`,
      `Operatori: ${Object.keys(payload?.operators || {}).length}`,
      `Gruppi: ${Object.keys(payload?.groups || {}).length}`,
      `Categorie configurate: ${collectConfiguredCategories(payload).length}`,
      `Stati AMP: ${Array.isArray(payload?.ticketing?.ampStatuses) ? payload.ticketing.ampStatuses.length : 0}`,
      `Stati SharePoint: ${Array.isArray(payload?.ticketing?.sharePointStatuses) ? payload.ticketing.sharePointStatuses.length : 0}`,
      `Workflow: ${Array.isArray(payload?.ticketing?.workflows) ? payload.ticketing.workflows.length : 0}`
    ];
    setText("diagnostic-database", lines.join("\n"));
  }
  function validateTicketingConfig(payload) {
    const ticketing = payload?.ticketing;
    const requiredCollections = ["ampStatuses", "workflows"];
    const missing = requiredCollections.filter((name) =>
      !Array.isArray(ticketing?.[name]) || !ticketing[name].some((value) => String(value || "").trim())
    );
    if (missing.length) {
      throw new Error(`PluginConfig incompleta: ticketing.${missing.join(", ticketing.")} mancante o vuoto.`);
    }
    if (!collectConfiguredCategories(payload).length) {
      throw new Error("PluginConfig incompleta: nessun tag presente in operators[*].tag, groups[*].tag o categories.generic[*].tag.");
    }
    return payload;
  }

  async function ensureOperationalConfig(accessToken) {
    const required = TEST_PLUGIN_BUILD.database;
    const cached = readOperationalConfigCache();
    window.WorkflowHealth?.recordDiagnostic?.("config_cache_read", { schemaVersion: cached?.schemaVersion ?? null, dataVersion: cached?.dataVersion ?? null });
    if (cached?.schemaVersion === required.schemaVersion && cached?.dataVersion === required.requiredDataVersion && cached?.payload) {
      try {
        validateTicketingConfig(cached.payload);
        window.WorkflowHealth?.recordDiagnostic?.("config_cache_valid", { dataVersion: cached.dataVersion });
        refreshOperationalConfigDiagnostics(cached, "cache valida");
        developerConfig = cached.payload;
        return cached.payload;
      } catch (error) {
        window.WorkflowHealth?.recordDiagnostic?.("config_cache_invalid", { message: error?.message || String(error) });
        refreshOperationalConfigDiagnostics(cached, "cache incompleta: nuovo download necessario");
      }
    }
    window.WorkflowHealth?.recordDiagnostic?.("config_version_mismatch", { requiredSchemaVersion: required.schemaVersion, requiredDataVersion: required.requiredDataVersion, cachedSchemaVersion: cached?.schemaVersion ?? null, cachedDataVersion: cached?.dataVersion ?? null });
    refreshOperationalConfigDiagnostics(cached, "aggiornamento necessario");
    window.WorkflowHealth?.recordDiagnostic?.("config_download_start");
    try {
      const response = await fetch(OPERATIONAL_CONFIG_URL, { method: "GET", headers: { Authorization: `Bearer ${accessToken}`, Accept: "application/json" } });
      if (!response.ok) throw new Error(`PluginConfig HTTP ${response.status} ${response.statusText || ""}`.trim());
      const payload = await response.json();
      if (payload?.schemaVersion !== required.schemaVersion || payload?.dataVersion !== required.requiredDataVersion) {
        throw new Error(`PluginConfig incompatibile: schema ${payload?.schemaVersion}, dati ${payload?.dataVersion}.`);
      }
      validateTicketingConfig(payload);
      const stored = { schemaVersion: payload.schemaVersion, dataVersion: payload.dataVersion, downloadedAt: new Date().toISOString(), payload };
      localStorage.setItem(OPERATIONAL_CONFIG_KEY, JSON.stringify(stored));
      window.WorkflowHealth?.recordDiagnostic?.("config_cache_saved", { schemaVersion: stored.schemaVersion, dataVersion: stored.dataVersion, agencies: Object.keys(payload.agencies || {}).length, operators: Object.keys(payload.operators || {}).length, groups: Object.keys(payload.groups || {}).length });
      refreshOperationalConfigDiagnostics(stored, "scaricata");
      developerConfig = payload;
      return payload;
    } catch (error) {
      window.WorkflowHealth?.recordDiagnostic?.("config_download_error", { message: error?.message || String(error) });
      refreshOperationalConfigDiagnostics(cached, `errore: ${error?.message || String(error)}`);
      throw error;
    }
  }

  function normalizeEmail(value) {
    return String(value || "").trim().toLowerCase();
  }

  function evaluateDeveloperToolsAccess(config) {
    const outlookEmail = normalizeEmail(Office.context?.mailbox?.userProfile?.emailAddress);
    const operators = Object.values(config?.operators || {});
    const roles = Object.values(config?.roles || {});
    const operator = operators.find((candidate) =>
      candidate?.enabled !== false && normalizeEmail(candidate?.email) === outlookEmail
    );
    const role = roles.find((candidate) =>
      String(candidate?.role || "").trim().toLowerCase() === String(operator?.role || "").trim().toLowerCase()
    );
    const priority = Number(role?.priority);

    return {
      allowed: Boolean(operator && Number.isFinite(priority) && priority <= 1),
      operator,
      role,
      outlookEmail
    };
  }

  function refreshDeveloperToolsAccess(config) {
    const access = evaluateDeveloperToolsAccess(config);
    developerConfig = config || null;
    developerToolsAllowed = access.allowed;
    developerCurrentOperatorId = Object.entries(config?.operators || {}).find(([, candidate]) => candidate === access.operator)?.[0] || "";
    const toggle = byId("email-details-toggle");
    const panel = byId("email-details-panel");
    if (toggle) toggle.hidden = !developerToolsAllowed;
    if (!developerToolsAllowed && panel) {
      panel.hidden = true;
      toggle?.setAttribute("aria-expanded", "false");
    }
    populateCreateTicketForm(config, developerCurrentOperatorId);
    populateCategoryForm(config);

    window.WorkflowHealth?.recordDiagnostic?.("developer_tools_access_evaluated", {
      allowed: developerToolsAllowed,
      operatorId: developerCurrentOperatorId || null,
      role: access.operator?.role || null,
      priority: Number.isFinite(Number(access.role?.priority)) ? Number(access.role.priority) : null,
      emailMatched: Boolean(access.operator)
    });
    return developerToolsAllowed;
  }

  async function initializeOperationalConfig() {
    const required = TEST_PLUGIN_BUILD.database;
    const cached = readOperationalConfigCache();
    let cacheIsValid = Boolean(
      cached?.schemaVersion === required.schemaVersion &&
      cached?.dataVersion === required.requiredDataVersion &&
      cached?.payload
    );

    if (cacheIsValid) {
      try { validateTicketingConfig(cached.payload); }
      catch { cacheIsValid = false; }
    }

    if (cacheIsValid) {
      refreshOperationalConfigDiagnostics(cached, "cache valida");
      refreshDeveloperToolsAccess(cached.payload);
      return cached.payload;
    }

    refreshDeveloperToolsAccess(null);
    if (!window.WorkflowHealth?.getAccessTokenSilently) {
      window.WorkflowHealth?.recordDiagnostic?.("config_initialize_skipped", { reason: "silent_auth_api_unavailable" });
      return null;
    }

    try {
      window.WorkflowHealth?.recordDiagnostic?.("config_initialize_start");
      const authentication = await window.WorkflowHealth.getAccessTokenSilently();
      if (!authentication?.accessToken) {
        throw new Error("Token di accesso non disponibile per inizializzare PluginConfig.");
      }
      const config = await ensureOperationalConfig(authentication.accessToken);
      refreshDeveloperToolsAccess(config);
      window.WorkflowHealth?.recordDiagnostic?.("config_initialize_complete", {
        authenticationMode: authentication.acquisitionMode || null
      });
      return config;
    } catch (error) {
      refreshDeveloperToolsAccess(null);
      window.WorkflowHealth?.recordDiagnostic?.("config_initialize_error", {
        name: error?.name || "Error",
        errorCode: error?.errorCode || null,
        message: error?.message || String(error)
      });
      console.error("Inizializzazione PluginConfig non riuscita:", error);
      return null;
    } finally {
      refreshAuthDiagnostics();
    }
  }

  const DEVELOPER_ENDPOINTS = Object.freeze({
    GetRecordByConversationId: "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/GetRecordByConversationId",
    GetRecordById: "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/GetRecordById",
    CreateTicket: "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/ManageTicketMetadata"
  });

  function resetTicketResult() {
    const panel = byId("developer-ticket-result");
    if (panel) panel.hidden = true;
    developerResultView = "fields";
    ["developer-result-action", "developer-ticket-status", "developer-ticket-auth-mode", "developer-ticket-response"].forEach((id) => setText(id, "", ""));
    byId("developer-ticket-fields")?.replaceChildren();
  }

  function getDeveloperAction() {
    return byId("developer-action-select")?.value || "GetRecordByConversationId";
  }

  function setSelectOptions(select, options, preferredValue = "") {
    if (!select) return;
    const previous = String(select.value || "");
    select.replaceChildren();
    options.forEach(({ value, label }) => {
      const option = document.createElement("option");
      option.value = String(value);
      option.textContent = String(label);
      select.appendChild(option);
    });
    const wanted = String(previous || preferredValue || "");
    if (wanted && options.some((entry) => String(entry.value) === wanted)) select.value = wanted;
  }

  function populateCreateTicketForm(config = developerConfig, currentOperatorId = developerCurrentOperatorId) {
    const status = byId("developer-create-config-status");
    const operators = Object.entries(config?.operators || {})
      .filter(([, operator]) => operator?.enabled !== false)
      .map(([id, operator]) => ({ value: id, label: `${operator?.displayName || id} · ${id}` }));
    const groups = Object.entries(config?.groups || {})
      .filter(([, group]) => group?.enabled !== false)
      .map(([id, group]) => ({ value: group?.displayName || id, label: `${group?.displayName || id} · ${id}` }));
    const ampStatuses = (config?.ticketing?.ampStatuses || []).map((value) => ({ value, label: value }));

    setSelectOptions(byId("developer-create-operator-id"), operators, currentOperatorId);
    setSelectOptions(byId("developer-create-requester-id"), operators, currentOperatorId);
    setSelectOptions(byId("developer-create-group-name"), groups);
    setSelectOptions(byId("developer-create-amp-status"), ampStatuses);

    const itemInput = byId("developer-create-item-id");
    if (itemInput && !itemInput.value && developerLastItemId) itemInput.value = developerLastItemId;
    if (status) {
      const workflowsOk = Array.isArray(config?.ticketing?.workflows) && config.ticketing.workflows.includes("CreateTicket");
      status.textContent = config
        ? (workflowsOk ? "Valori caricati da PluginConfig." : "PluginConfig non abilita il workflow CreateTicket.")
        : "Configurazione operativa non disponibile.";
      status.className = `developer-create-help${config && !workflowsOk ? " status-error" : ""}`;
    }
  }

  function populateCategoryForm(config = developerConfig) {
    const categories = collectConfiguredCategories(config);
    const list = byId("developer-category-options");
    if (list) {
      list.replaceChildren();
      categories.forEach((category) => {
        const option = document.createElement("option");
        option.value = category;
        list.appendChild(option);
      });
    }
    const available = byId("developer-category-available");
    if (available) {
      available.textContent = categories.length
        ? `Categorie configurate: ${categories.join(", ")}`
        : "Categorie configurate non disponibili.";
      available.className = `developer-create-help${categories.length ? "" : " status-error"}`;
    }
  }

  function updateDeveloperActionUi() {
    const action = getDeveloperAction();
    const isCreate = action === "CreateTicket";
    const isCategory = action === "AddCategoriesOfficeJs" || action === "RemoveCategoriesOfficeJs";
    const isAddCategory = action === "AddCategoriesOfficeJs";
    const isConversationLookup = action === "GetRecordByConversationId";
    const getForm = byId("developer-get-form");
    const createForm = byId("developer-create-form");
    const categoryForm = byId("developer-category-form");
    if (getForm) getForm.hidden = isCreate || isCategory;
    if (createForm) createForm.hidden = !isCreate;
    if (categoryForm) categoryForm.hidden = !isCategory;
    const replaceRow = byId("developer-category-replace-row");
    const clearRow = byId("developer-category-clear-row");
    if (replaceRow) replaceRow.hidden = !isAddCategory;
    if (clearRow) clearRow.hidden = isAddCategory;
    const conversationModeRow = byId("developer-conversation-mode-row");
    if (conversationModeRow) conversationModeRow.hidden = !isConversationLookup;

    if (isCreate) {
      populateCreateTicketForm();
      const itemInput = byId("developer-create-item-id");
      if (itemInput && developerLastItemId) itemInput.value = developerLastItemId;
      return;
    }

    if (isCategory) {
      populateCategoryForm();
      const label = byId("developer-category-input-label");
      if (label) label.textContent = isAddCategory
        ? "Categorie da aggiungere (separate da virgola)"
        : "Categorie da rimuovere (separate da virgola)";
      return;
    }

    const label = byId("developer-action-input-label");
    const input = byId("developer-action-input");
    if (!label || !input) return;
    if (action === "GetRecordById") {
      label.textContent = "Item ID";
      input.placeholder = "es. 9840";
      input.value = developerLastItemId || "";
    } else {
      label.textContent = "Conversation ID";
      input.placeholder = "Conversation ID del messaggio";
      input.value = currentMessageContext?.normalizedConversationId || "";
    }
  }

  function setTicketButtonLoading(value) {
    const button = byId("developer-execute-button");
    if (button) {
      button.disabled = value;
      button.textContent = value ? "Esecuzione..." : "Esegui";
    }
  }

  function parseAmpMetadata(value) {
    if (!value) return { tickets: [], error: null };
    try {
      const parsed = typeof value === "string" ? JSON.parse(value) : value;
      return { tickets: Array.isArray(parsed?.tickets) ? parsed.tickets : [], error: null };
    } catch (error) {
      return { tickets: [], error: error?.message || String(error) };
    }
  }

  function findOperatorByLookupId(id) {
    const config = developerConfig || readOperationalConfigCache()?.payload;
    return Object.entries(config?.operators || {}).find(([, operator]) => String(operator?.lookupId) === String(id)) || null;
  }

  function addDeveloperField(container, label, value) {
    const card = document.createElement("div");
    card.className = "developer-field-card";
    if (label === "Note ticket") card.classList.add("developer-note-field");
    const name = document.createElement("span");
    name.textContent = label;
    const content = document.createElement("strong");
    content.textContent = value == null || value === "" ? "Non disponibile" : String(value);
    card.append(name, content);
    container.appendChild(card);
  }

  function removeRecordLogs(value) {
    if (Array.isArray(value)) return value.map(removeRecordLogs);
    if (!value || typeof value !== "object") return value;
    return Object.fromEntries(
      Object.entries(value)
        .filter(([key]) => {
          const normalizedKey = key.replace(/[^a-z]/gi, "").toLowerCase();
          return !["log", "logs", "logentry", "logentries", "ticketlog", "ticketlogs"].includes(normalizedKey);
        })
        .map(([key, entry]) => [key, removeRecordLogs(entry)])
    );
  }

  function getRecordResultEntries(parsed) {
    if (Array.isArray(parsed)) return parsed;
    const collections = [
      parsed?.records,
      parsed?.items,
      parsed?.results,
      parsed?.value,
      parsed?.data,
      parsed?.data?.records,
      parsed?.data?.items
    ];
    const collection = collections.find(Array.isArray);
    if (collection) return collection;
    if (parsed?.record && typeof parsed.record === "object") return [parsed.record];
    return parsed?.fields && typeof parsed.fields === "object" ? [parsed] : [];
  }

  function hasRecordResultShape(parsed) {
    return Array.isArray(parsed) || Boolean(
      parsed?.record ||
      parsed?.fields ||
      [parsed?.records, parsed?.items, parsed?.results, parsed?.value, parsed?.data, parsed?.data?.records, parsed?.data?.items].some(Array.isArray)
    );
  }

  function getRecordFields(record) {
    return record?.fields && typeof record.fields === "object" ? record.fields : {};
  }

  function getRecordItemId(record) {
    const fields = getRecordFields(record);
    return record?.itemId || record?.id || fields?.id || fields?.Id || fields?.ID || "";
  }

  function getTicketNotes(record) {
    const fields = getRecordFields(record);
    const candidates = [
      record?.notes,
      record?.note,
      fields?.Notes,
      fields?.Note,
      fields?.notes,
      fields?.note,
      fields?.TicketNotes,
      fields?.TicketNote
    ];
    const direct = candidates.find((value) => value !== undefined && value !== null && value !== "");
    if (direct !== undefined) return direct;
    const dynamicEntry = Object.entries(fields).find(([key, value]) => {
      if (value === undefined || value === null || value === "") return false;
      const normalizedKey = key.replace(/_x[0-9a-f]{4}_/gi, "").replace(/[^a-z]/gi, "").toLowerCase();
      return ["note", "notes", "ticketnote", "ticketnotes", "noteticket", "notesticket"].includes(normalizedKey);
    });
    return dynamicEntry?.[1];
  }

  function formatTicketNotes(notes) {
    if (Array.isArray(notes)) {
      return notes.map(formatTicketNotes).filter(Boolean).join("\n\n");
    }
    if (notes && typeof notes === "object") {
      const text = notes.text ?? notes.body ?? notes.content ?? notes.note ?? notes.value;
      if (text !== undefined && text !== null && text !== "") return String(text);
      return JSON.stringify(removeRecordLogs(notes), null, 2);
    }
    return notes == null ? "" : String(notes);
  }

  function renderSelectedAmpTicket(container, ticket) {
    const details = document.createElement("div");
    details.className = "developer-fields";
    addDeveloperField(details, "Ticket interno", ticket?.ticketId);
    addDeveloperField(details, "Riferimento AMP", ticket?.ampRequestId);
    addDeveloperField(details, "Stato AMP", ticket?.status || ticket?.ampStatus);
    addDeveloperField(details, "Lock operatore", ticket?.lock?.operatorId || "Nessun lock");
    addDeveloperField(details, "Lock richiesto da", ticket?.lock?.requesterId);
    addDeveloperField(details, "Workflow lock", ticket?.lock?.workflow);
    addDeveloperField(details, "Lock acquisito", ticket?.lock?.acquiredAt);
    addDeveloperField(details, "Creato", ticket?.createdAt);
    addDeveloperField(details, "Ultima attività", ticket?.lastActivityAt);
    container.appendChild(details);
  }

  function appendTicketSelector(container, tickets, labelText = "Ticket AMP") {
    if (!tickets.length) return;
    const selectorBox = document.createElement("div");
    selectorBox.className = "developer-ticket-selector";
    const label = document.createElement("label");
    label.textContent = labelText;
    const selector = document.createElement("select");
    tickets.forEach((ticket, index) => {
      const option = document.createElement("option");
      option.value = String(index);
      option.textContent = ticket?.ticketId || `Ticket ${index + 1}`;
      selector.appendChild(option);
    });
    const host = document.createElement("div");
    const redraw = () => {
      host.replaceChildren();
      renderSelectedAmpTicket(host, tickets[Number(selector.value) || 0]);
    };
    selector.addEventListener("change", redraw);
    selectorBox.append(label, selector);
    container.append(selectorBox, host);
    redraw();
  }

  function renderRecordFields(container, record) {
    const fields = getRecordFields(record);
    const itemId = getRecordItemId(record);
    const lookupId = fields?.AssegnataaLookupId;
    const operator = findOperatorByLookupId(lookupId);
    addDeveloperField(container, "Item ID", itemId);
    addDeveloperField(container, "Gruppo", fields?.Gruppo);
    addDeveloperField(container, "Assegnatario", operator ? `${operator[1].displayName} (${operator[0]} · LookupId ${lookupId})` : (lookupId ? `LookupId ${lookupId}` : null));
    addDeveloperField(container, "Stato SharePoint", fields?._Status);
    const notes = getTicketNotes(record);
    if (notes !== undefined) addDeveloperField(container, "Note ticket", formatTicketNotes(notes));
    const metadata = parseAmpMetadata(fields?.AMP_x002d_AVANADE);
    if (metadata.error) {
      const error = document.createElement("div");
      error.className = "developer-ticket-empty status-error";
      error.textContent = `Metadata AMP non valido: ${metadata.error}`;
      container.appendChild(error);
      return;
    }
    if (!metadata.tickets.length) {
      const empty = document.createElement("div");
      empty.className = "developer-ticket-empty";
      empty.textContent = "Nessun ticket AMP presente.";
      container.appendChild(empty);
      return;
    }
    appendTicketSelector(container, metadata.tickets);
  }

  function appendRecordSelector(container, records) {
    addDeveloperField(container, "Record restituiti", records.length);
    if (!records.length) {
      const empty = document.createElement("div");
      empty.className = "developer-ticket-empty";
      empty.textContent = "Nessun record restituito.";
      container.appendChild(empty);
      return;
    }

    if (records.length === 1) {
      renderRecordFields(container, records[0]);
      return;
    }

    const selectorBox = document.createElement("div");
    selectorBox.className = "developer-ticket-selector";
    const label = document.createElement("label");
    label.textContent = "Record SharePoint";
    const selector = document.createElement("select");
    records.forEach((record, index) => {
      const option = document.createElement("option");
      option.value = String(index);
      const itemId = getRecordItemId(record);
      option.textContent = itemId ? `Item ${itemId}` : `Record ${index + 1}`;
      selector.appendChild(option);
    });
    const host = document.createElement("div");
    const redraw = () => {
      host.replaceChildren();
      host.className = "developer-fields";
      renderRecordFields(host, records[Number(selector.value) || 0]);
    };
    selector.addEventListener("change", redraw);
    selectorBox.append(label, selector);
    container.append(selectorBox, host);
    redraw();
  }

  function renderRecordResultFields(parsed) {
    const container = byId("developer-ticket-fields");
    if (!container) return;
    container.replaceChildren();
    container.className = "developer-fields";
    appendRecordSelector(container, getRecordResultEntries(parsed));
  }

  function ticketSummaryValue(ticket) {
    return [
      ticket?.ticketId,
      ticket?.ampRequestId ? `AMP ${ticket.ampRequestId}` : null,
      ticket?.status || ticket?.ampStatus
    ].filter(Boolean).join(" · ");
  }

  function appendExistingTickets(container, tickets) {
    const list = document.createElement("div");
    list.className = "developer-existing-tickets";
    tickets.forEach((ticket, index) => {
      const card = document.createElement("div");
      card.className = "developer-existing-ticket";
      const heading = document.createElement("h4");
      heading.textContent = ticketSummaryValue(ticket) || `Lavorazione ${index + 1}`;
      const details = document.createElement("dl");
      [
        ["Ticket interno", ticket?.ticketId],
        ["Riferimento AMP", ticket?.ampRequestId],
        ["Stato AMP", ticket?.status || ticket?.ampStatus],
        ["Operatore", ticket?.lock?.operatorId],
        ["Workflow", ticket?.lock?.workflow]
      ].forEach(([label, value]) => {
        const term = document.createElement("dt");
        const description = document.createElement("dd");
        term.textContent = label;
        description.textContent = value == null || value === "" ? "Non disponibile" : String(value);
        details.append(term, description);
      });
      card.append(heading, details);
      list.appendChild(card);
    });
    container.appendChild(list);
  }

  function getCreateResponseTickets(parsed) {
    if (Array.isArray(parsed?.metadata?.tickets)) return parsed.metadata.tickets;
    if (Array.isArray(parsed?.tickets)) return parsed.tickets;
    if (Array.isArray(parsed?.existingTickets)) return parsed.existingTickets;
    return [];
  }

  function renderCreateTicketFields(parsed) {
    const container = byId("developer-ticket-fields");
    if (!container) return;
    container.replaceChildren();
    container.className = "developer-fields";
    const body = developerLastCreateBody || {};
    const tickets = getCreateResponseTickets(parsed);
    const createdTicketId = parsed?.ticketId || parsed?.createdTicket?.ticketId;
    const createdTicket = parsed?.createdTicket || tickets.find((ticket) => ticket?.ticketId === createdTicketId) || tickets.at(-1) || null;
    const reason = String(parsed?.reason || parsed?.code || parsed?.status || "");

    addDeveloperField(container, "Operazione", parsed?.action || "CreateTicket");
    addDeveloperField(container, "Item ID", parsed?.itemId || body.itemId);

    if (reason === "existing_tickets_found") {
      const conflict = document.createElement("div");
      conflict.className = "developer-conflict";
      const heading = document.createElement("h4");
      heading.textContent = "Creazione non eseguita";
      const explanation = document.createElement("p");
      explanation.textContent = "Esistono già lavorazioni associate all'elemento SharePoint.";
      conflict.append(heading, explanation);
      const existingTickets = Array.isArray(parsed?.existingTickets)
        ? parsed.existingTickets
        : (Array.isArray(parsed?.existing_tickets) ? parsed.existing_tickets : tickets);
      if (existingTickets.length) appendExistingTickets(conflict, existingTickets);
      addDeveloperField(conflict, "AMP prenotato", parsed?.reservedAmpRequestId || parsed?.reserved_amp_request_id || body.ampRequestId);
      const retry = document.createElement("button");
      retry.type = "button";
      retry.className = "primary-button developer-force-create";
      retry.textContent = "Crea comunque nuova lavorazione";
      retry.addEventListener("click", () => {
        if (!developerLastCreateBody) return;
        void executeDeveloperAction({ ...developerLastCreateBody, createAdditionalTicket: true });
      });
      conflict.appendChild(retry);
      container.appendChild(conflict);
      return;
    }

    addDeveloperField(container, "Esito", parsed?.status || parsed?.message);
    addDeveloperField(container, "Ticket creato", createdTicketId || createdTicket?.ticketId);
    addDeveloperField(container, "Riferimento AMP", parsed?.ampRequestId || createdTicket?.ampRequestId || body.ampRequestId);
    addDeveloperField(container, "Stato AMP", createdTicket?.status || createdTicket?.ampStatus || body.ampStatus);
    addDeveloperField(container, "Stato SharePoint", parsed?.sharePointStatus);
    addDeveloperField(container, "Gruppo", parsed?.groupName || body.groupName);
    addDeveloperField(container, "Lock", createdTicket?.lock ? `${createdTicket.lock.operatorId || body.operatorId} · ${createdTicket.lock.workflow || "CreateTicket"}` : `${body.operatorId || "-"} · CreateTicket`);
    addDeveloperField(container, "Totale lavorazioni", Array.isArray(parsed?.metadata?.tickets) ? parsed.metadata.tickets.length : (parsed?.totalTickets ?? (tickets.length || null)));
    if (parsed?.message && parsed?.status === "error") addDeveloperField(container, "Errore", parsed.message);
    if (tickets.length) appendTicketSelector(container, tickets, "Lavorazioni aggiornate");
  }

  function isCategoryDeveloperAction(action) {
    return action === "AddCategoriesOfficeJs" || action === "RemoveCategoriesOfficeJs";
  }

  function isRecordDeveloperAction(action) {
    return action === "GetRecordByConversationId" || action === "GetRecordById";
  }

  function formatCategoryNameList(value) {
    if (!Array.isArray(value)) return "Non disponibile";
    return value.length ? value.join(", ") : "Nessuna";
  }

  function renderCategoryFields(parsed, action) {
    const container = byId("developer-ticket-fields");
    if (!container) return;
    container.replaceChildren();
    container.className = "developer-fields";
    const isAdd = action === "AddCategoriesOfficeJs";
    const mode = isAdd
      ? (parsed?.replace ? "Sostituzione completa" : "Aggiunta")
      : (parsed?.clear ? "Rimozione completa" : "Rimozione selettiva");
    addDeveloperField(container, "Operazione", action);
    addDeveloperField(container, "Modalità", mode);
    addDeveloperField(container, "Categorie richieste", formatCategoryNameList(parsed?.categories));
    addDeveloperField(container, "Categorie before", formatCategoryNameList(parsed?.categoriesBefore));
    addDeveloperField(container, "Categorie after", formatCategoryNameList(parsed?.categoriesAfter));
    if (parsed?.message && parsed?.status === "error") addDeveloperField(container, "Errore", parsed.message);
  }

  function applyDeveloperResultView() {
    const raw = byId("developer-ticket-response");
    const fields = byId("developer-ticket-fields");
    const button = byId("developer-result-view-toggle");
    const showJson = developerResultView === "json";
    if (raw) raw.hidden = !showJson;
    if (fields) fields.hidden = showJson;
    if (button) button.textContent = showJson ? "Mostra campi" : "Mostra JSON";
  }

  function showTicketResult(result, action) {
    const panel = byId("developer-ticket-result");
    const status = byId("developer-ticket-status");
    const auth = byId("developer-ticket-auth-mode");
    const response = byId("developer-ticket-response");
    if (!panel || !status || !auth || !response) return;
    panel.hidden = false;
    setText("developer-result-action", action, "");
    status.textContent = `${result.status} ${result.statusText || ""}`.trim();
    status.className = result.ok ? "status-success" : "status-error";
    auth.textContent = result.authenticationMode || "Non disponibile";
    let parsed = null;
    try {
      parsed = JSON.parse(result.body || "{}");
      if (isRecordDeveloperAction(action)) parsed = removeRecordLogs(parsed);
      response.textContent = JSON.stringify(parsed, null, 2);
    } catch {
      response.textContent = result.body || "(Risposta vuota)";
    }

    const returnedRecord = isRecordDeveloperAction(action) ? getRecordResultEntries(parsed)[0] : null;
    const returnedItemId = returnedRecord ? getRecordItemId(returnedRecord) : (parsed?.itemId || parsed?.fields?.id);
    if (action !== "CreateTicket" && returnedItemId) {
      developerLastItemId = String(returnedItemId);
      const createItemInput = byId("developer-create-item-id");
      if (createItemInput) createItemInput.value = developerLastItemId;
    }

    const hasFieldView = isRecordDeveloperAction(action)
      ? hasRecordResultShape(parsed)
      : (action === "CreateTicket" || isCategoryDeveloperAction(action)
        ? Boolean(parsed && typeof parsed === "object")
        : Boolean(parsed?.fields));
    developerResultView = hasFieldView ? "fields" : "json";
    if (action === "CreateTicket" && parsed) renderCreateTicketFields(parsed);
    else if (isCategoryDeveloperAction(action) && parsed) renderCategoryFields(parsed, action);
    else if (isRecordDeveloperAction(action) && hasRecordResultShape(parsed)) renderRecordResultFields(parsed);
    else byId("developer-ticket-fields")?.replaceChildren();
    applyDeveloperResultView();
  }

  function requiredFormValue(id, label) {
    const value = String(byId(id)?.value || "").trim();
    if (!value) throw new Error(`${label} è obbligatorio.`);
    return value;
  }

  function parseCategoryInput(value) {
    const unique = new Map();
    String(value || "").split(",").forEach((part) => {
      const category = part.trim();
      const key = category.toLocaleLowerCase("it-IT");
      if (category && !unique.has(key)) unique.set(key, category);
    });
    return Array.from(unique.values());
  }

  function canonicalizeConfiguredCategories(values, config = developerConfig) {
    const configured = collectConfiguredCategories(config);
    const byKey = new Map(configured.map((category) => [category.toLocaleLowerCase("it-IT"), category]));
    const unknown = values.filter((category) => !byKey.has(category.toLocaleLowerCase("it-IT")));
    if (unknown.length) {
      throw new Error(`Categorie non presenti in PluginConfig: ${unknown.join(", ")}.`);
    }
    return values.map((category) => byKey.get(category.toLocaleLowerCase("it-IT")));
  }

  function normalizeCategoryActionRequest(action, categoriesValue, option) {
    const rawCategories = parseCategoryInput(Array.isArray(categoriesValue) ? categoriesValue.join(",") : categoriesValue);
    if (action === "AddCategoriesOfficeJs") {
      const categories = canonicalizeConfiguredCategories(rawCategories);
      if (!categories.length) throw new Error("Inserisci almeno una categoria da aggiungere.");
      return { categories, replace: Boolean(option) };
    }
    const clear = Boolean(option);
    if (clear) return { categories: [], clear: true };
    const categories = canonicalizeConfiguredCategories(rawCategories);
    if (!categories.length) throw new Error("Inserisci almeno una categoria da rimuovere oppure attiva clear.");
    return { categories, clear: false };
  }

  async function executeOfficeCategoryAction(action, request) {
    const item = Office.context?.mailbox?.item;
    if (!item || !currentMessageContext) {
      throw new Error("Seleziona un messaggio Outlook e attendi il completamento della lettura automatica.");
    }

    validateTicketingConfig(developerConfig);
    let categoriesBefore = null;
    let categoriesAfter = null;
    let operationError = null;

    try {
      const before = await getCategories(item);
      categoriesBefore = categoryDisplayNames(before);
      const beforeByKey = new Map(categoriesBefore.map((category) => [category.toLocaleLowerCase("it-IT"), category]));

      if (action === "AddCategoriesOfficeJs") {
        if (request.replace) {
          await removeCategories(item, categoriesBefore);
          await addCategories(item, request.categories);
        } else {
          const missing = request.categories.filter((category) => !beforeByKey.has(category.toLocaleLowerCase("it-IT")));
          await addCategories(item, missing);
        }
      } else {
        const namesToRemove = request.clear
          ? categoriesBefore
          : request.categories
            .map((category) => beforeByKey.get(category.toLocaleLowerCase("it-IT")))
            .filter(Boolean);
        await removeCategories(item, namesToRemove);
      }
    } catch (error) {
      operationError = error;
    }

    try {
      const after = await getCategories(item);
      categoriesAfter = categoryDisplayNames(after);
      updateCurrentMessageCategories(after);
    } catch (error) {
      operationError ||= error;
    }

    const payload = {
      status: operationError ? "error" : "ok",
      action,
      categories: request.categories,
      ...(action === "AddCategoriesOfficeJs" ? { replace: request.replace } : { clear: request.clear }),
      categoriesBefore,
      categoriesAfter,
      ...(operationError ? { message: operationError?.message || String(operationError) } : {})
    };
    return {
      ok: !operationError,
      status: operationError ? "Errore" : 200,
      statusText: operationError ? "" : "OK",
      authenticationMode: "Office.js",
      body: JSON.stringify(payload)
    };
  }

  async function AddCategoriesOfficeJs(categories, replace = false) {
    const request = normalizeCategoryActionRequest("AddCategoriesOfficeJs", categories, replace);
    return executeOfficeCategoryAction("AddCategoriesOfficeJs", request);
  }

  async function RemoveCategoriesOfficeJs(categories, clear = false) {
    const request = normalizeCategoryActionRequest("RemoveCategoriesOfficeJs", categories, clear);
    return executeOfficeCategoryAction("RemoveCategoriesOfficeJs", request);
  }

  function collectCreateTicketBody() {
    validateTicketingConfig(developerConfig);
    if (!developerConfig.ticketing.workflows.includes("CreateTicket")) {
      throw new Error("Il workflow CreateTicket non è disponibile in PluginConfig.");
    }
    return {
      itemId: requiredFormValue("developer-create-item-id", "Item ID"),
      action: "CreateTicket",
      ampRequestId: requiredFormValue("developer-create-amp-request-id", "Riferimento AMP"),
      ampStatus: requiredFormValue("developer-create-amp-status", "Stato AMP"),
      operatorId: requiredFormValue("developer-create-operator-id", "Operatore"),
      requesterId: requiredFormValue("developer-create-requester-id", "Requester"),
      groupName: requiredFormValue("developer-create-group-name", "Gruppo"),
      createAdditionalTicket: Boolean(byId("developer-create-additional")?.checked)
    };
  }

  function getConversationRecordMode() {
    const mode = String(byId("developer-conversation-mode")?.value || "all").trim().toLowerCase();
    if (!["all", "first", "last"].includes(mode)) {
      throw new Error("Mode non valido: usa all, first oppure last.");
    }
    return mode;
  }

  async function executeDeveloperAction(createBodyOverride = null) {
    const action = createBodyOverride ? "CreateTicket" : getDeveloperAction();
    const value = String(byId("developer-action-input")?.value || "").trim();
    if (action === "CreateTicket" && !createBodyOverride) developerLastCreateBody = null;
    resetTicketResult();
    setTicketButtonLoading(true);
    try {
      if (isCategoryDeveloperAction(action)) {
        developerConfig ||= readOperationalConfigCache()?.payload || null;
        validateTicketingConfig(developerConfig);
        const categories = String(byId("developer-category-input")?.value || "");
        const categoryResult = action === "AddCategoriesOfficeJs"
          ? await AddCategoriesOfficeJs(categories, Boolean(byId("developer-category-replace")?.checked))
          : await RemoveCategoriesOfficeJs(categories, Boolean(byId("developer-category-clear")?.checked));
        showTicketResult(categoryResult, action);
        return;
      }
      if (action !== "CreateTicket" && !value) {
        throw new Error(action === "GetRecordById" ? "Inserisci un Item ID." : "Inserisci un Conversation ID.");
      }
      const authentication = await window.WorkflowHealth.getAccessTokenSilently();
      if (!authentication?.accessToken) throw new Error("Token di accesso non disponibile.");
      const config = await ensureOperationalConfig(authentication.accessToken);
      refreshDeveloperToolsAccess(config);

      let url;
      let request;
      let requestContext;
      if (action === "CreateTicket") {
        const body = createBodyOverride ? { ...createBodyOverride } : collectCreateTicketBody();
        body.action = "CreateTicket";
        body.createAdditionalTicket = Boolean(body.createAdditionalTicket);
        developerLastCreateBody = { ...body };
        url = DEVELOPER_ENDPOINTS.CreateTicket;
        request = {
          method: "POST",
          headers: {
            Authorization: `Bearer ${authentication.accessToken}`,
            Accept: "application/json",
            "Content-Type": "application/json"
          },
          body: JSON.stringify(body)
        };
        requestContext = body;
      } else {
        const parameter = action === "GetRecordById" ? "itemId" : "conversationId";
        const query = new URLSearchParams({ [parameter]: value });
        const mode = action === "GetRecordByConversationId" ? getConversationRecordMode() : null;
        if (mode) query.set("mode", mode);
        url = `${DEVELOPER_ENDPOINTS[action]}?${query.toString()}`;
        request = {
          method: "GET",
          headers: { Authorization: `Bearer ${authentication.accessToken}`, Accept: "application/json" }
        };
        requestContext = { [parameter]: value, ...(mode ? { mode } : {}) };
      }

      const fetchResponse = await fetch(url, request);
      const responseBody = await fetchResponse.text();
      const result = {
        ok: fetchResponse.ok,
        status: fetchResponse.status,
        statusText: fetchResponse.statusText,
        authenticationMode: authentication.acquisitionMode,
        body: responseBody
      };
      showTicketResult(result, action);
      let parsed = null;
      try { parsed = JSON.parse(responseBody || "{}"); } catch {}
      if (parsed && isRecordDeveloperAction(action)) parsed = removeRecordLogs(parsed);
      const returnedRecord = isRecordDeveloperAction(action) ? getRecordResultEntries(parsed)[0] : null;
      currentSharePointData = {
        operation: action,
        conversationId: action === "GetRecordByConversationId" ? value : (parsed?.conversationId || returnedRecord?.conversationId || null),
        itemId: action === "CreateTicket" ? requestContext.itemId : (returnedRecord ? getRecordItemId(returnedRecord) : (parsed?.itemId || parsed?.fields?.id || null)),
        ok: result.ok,
        status: result.status,
        statusText: result.statusText,
        authenticationMode: result.authenticationMode,
        retrievedAt: new Date().toISOString(),
        request: action === "CreateTicket" || isRecordDeveloperAction(action) ? requestContext : null,
        data: parsed,
        rawBody: parsed ? null : responseBody
      };
      publishSharePointData();
    } catch (error) {
      showTicketResult({
        ok: false,
        status: "Errore",
        statusText: "",
        authenticationMode: "Non disponibile",
        body: JSON.stringify({ status: "error", message: error?.message || String(error) })
      }, action);
    } finally {
      setTicketButtonLoading(false);
    }
  }

  function registerDeveloperToolsControls() {
    byId("developer-action-select")?.addEventListener("change", updateDeveloperActionUi);
    byId("developer-execute-button")?.addEventListener("click", () => { void executeDeveloperAction(); });
    byId("developer-result-view-toggle")?.addEventListener("click", () => {
      developerResultView = developerResultView === "fields" ? "json" : "fields";
      applyDeveloperResultView();
    });
    byId("close-developer-ticket-result")?.addEventListener("click", resetTicketResult);
    byId(elementIds.idToggle)?.addEventListener("click", () => toggleIdentifierView("itemId"));
    byId(elementIds.conversationIdToggle)?.addEventListener("click", () => toggleIdentifierView("conversationId"));
    updateDeveloperActionUi();
  }

  async function loadSelectedMessage() {
    const sequence = ++loadSequence;

    currentMessageIdentifiers = null;
    resetIdentifierViews();
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

      const originalItemId = item.itemId || "";
      const originalConversationId = item.conversationId || "";
      const normalizedItemId = convertItemIdToRestV2(originalItemId);
      const normalizedConversationId = normalizeConversationId(originalConversationId);
      currentMessageIdentifiers = {
        originalItemId,
        originalConversationId,
        normalizedItemId,
        normalizedConversationId
      };
      renderIdentifierFields();
      setText(elementIds.subject, item.subject, "(Nessun oggetto)");
      setText(elementIds.sender, formatAddress(sender));
      setText(elementIds.to, formatAddressList(item.to));
      setText(elementIds.cc, formatAddressList(item.cc));
      setText(elementIds.received, formatReceivedDate(item.dateTimeCreated));
      setText(elementIds.body, bodyText, "(Corpo vuoto)");
      setText(elementIds.categories, formatCategories(categories),"Nessuna");
      currentMessageContext = {
        itemId: normalizedItemId || null,
        conversationId: normalizedConversationId || null,
        normalizedConversationId: normalizedConversationId || null,
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
      syncDeveloperConversationId();
    } catch (error) {
      if (sequence !== loadSequence) {
        return;
      }

      currentMessageContext = null;
      currentMessageIdentifiers = null;
      resetIdentifierViews();
      publishMessageContext();
      syncDeveloperConversationId();
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
    version: window.TestPluginBuild?.application || "0.2",
    getCurrentMessageContext,
    onMessageChanged,
    getCurrentSharePointData,
    onSharePointDataChanged,
    refreshCurrentMessage: loadSelectedMessage,
    showMainView,
    showDeveloperTools: showDeveloperToolsView
  });

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Outlook) {
      showError("Questa pagina deve essere aperta come task pane di Outlook.");
      return;
    }

    registerItemChangedHandler();
    registerSectionToggles();
    registerHealthButton();
    registerDeveloperToolsControls();
    registerAuthDiagnosticsControls();
    loadBuildVersions();
    refreshOperationalConfigDiagnostics();
    loadDiagnostics();
    registerPluginSettingsControls();
    registerHeaderViewControls();
    // La lettura Outlook e la configurazione operativa sono pipeline indipendenti.
    // Nessuna delle due esegue automaticamente GetRecordByConversationId.
    void loadSelectedMessage();
    void initializeOperationalConfig();
    refreshAuthDiagnostics();
  });
})();
