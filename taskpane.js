// Test-Plugin file version: 0.4-initial
const TEST_PLUGIN_BUILD = Object.freeze({
  application: '0.4-initial',
  release: '0.4-initial',
  files: Object.freeze({
    'taskpane.html': '0.4-initial',
    'taskpane.js': '0.4-initial',
    'taskpane.css': '0.4-initial',
    'auth-src.js': '0.2-auth-src.js-3',
    'auth-dialog.html': '0.2-auth-dialog.html-2',
    'auth-dialog-src.js': '0.2-auth-dialog-src.js-3',
    'auth-popup-redirect.html': '0.2-auth-popup-redirect.html-1',
    'auth-popup-redirect-src.js': '0.2-auth-popup-redirect-src.js-1',
  }),
  database: Object.freeze({ schemaVersion: 5, requiredDataVersion: 8 }),
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
  let officeCategoryLoadSequence = 0;
  let graphCategoryLoadSequence = 0;
  let developerGraphCategoryState = {
    mail: "",
    messageId: "",
    categories: [],
    loaded: false,
    loading: false,
    error: ""
  };
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

  function getMasterCategories(mailbox = Office.context?.mailbox) {
    return new Promise((resolve, reject) => {
      const api = mailbox?.masterCategories;
      if (!api || typeof api.getAsync !== "function") {
        reject(new Error("Il client Outlook non espone mailbox.masterCategories.getAsync."));
        return;
      }
      api.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(Array.isArray(result.value) ? result.value : []);
          return;
        }
        reject(new Error(result.error?.message || "Office.js non ha restituito la Master Category List."));
      });
    });
  }

  function categoryKey(value) {
    return String(value || "").trim().toLocaleLowerCase("it-IT");
  }

  function uniqueCategoryNames(values) {
    const unique = new Map();
    (Array.isArray(values) ? values : []).forEach((value) => {
      const name = String(value?.displayName || value || "").trim();
      const key = categoryKey(name);
      if (key && !unique.has(key)) unique.set(key, name);
    });
    return Array.from(unique.values());
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
      ...Object.values(payload?.forwardingTargets || {}).map((target) => target?.tag),
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
      `Destinazioni forwarding: ${Object.keys(payload?.forwardingTargets || {}).length}`,
      `Stati ticket: ${Array.isArray(payload?.ticketing?.statuses) ? payload.ticketing.statuses.length : 0}`,
      `Categorie configurate: ${collectConfiguredCategories(payload).length}`,
      `Mailbox categorie: ${String(payload?.ticketing?.email || "").trim() || "-"}`
    ];
    setText("diagnostic-database", lines.join("\n"));
  }

  function validateTicketingConfig(payload) {
    if (!payload || typeof payload !== "object") {
      throw new Error("PluginConfig mancante o non valida.");
    }
    const statuses = payload?.ticketing?.statuses;
    if (!Array.isArray(statuses) || !statuses.some((value) => String(value || "").trim())) {
      throw new Error("PluginConfig incompleta: ticketing.statuses mancante o vuoto.");
    }
    for (const name of ["operators", "roles", "forwardingTargets"]) {
      if (!payload[name] || typeof payload[name] !== "object" || !Object.keys(payload[name]).length) {
        throw new Error(`PluginConfig incompleta: ${name} mancante o vuoto.`);
      }
    }
    if (!collectConfiguredCategories(payload).length) {
      throw new Error("PluginConfig incompleta: nessun tag in operators, forwardingTargets o categories.generic.");
    }
    return payload;
  }

  async function ensureOperationalConfig(accessToken) {
    const required = TEST_PLUGIN_BUILD.database;
    const cached = readOperationalConfigCache();
    window.WorkflowHealth?.recordDiagnostic?.("config_cache_read", {
      schemaVersion: cached?.schemaVersion ?? null,
      dataVersion: cached?.dataVersion ?? null
    });
    if (cached?.schemaVersion === required.schemaVersion &&
        cached?.dataVersion === required.requiredDataVersion && cached?.payload) {
      try {
        validateTicketingConfig(cached.payload);
        refreshOperationalConfigDiagnostics(cached, "cache valida");
        developerConfig = cached.payload;
        return cached.payload;
      } catch (error) {
        window.WorkflowHealth?.recordDiagnostic?.("config_cache_invalid", { message: error?.message || String(error) });
      }
    }

    refreshOperationalConfigDiagnostics(cached, "aggiornamento necessario");
    const response = await fetch(OPERATIONAL_CONFIG_URL, {
      method: "GET",
      headers: { Authorization: `Bearer ${accessToken}`, Accept: "application/json" }
    });
    if (!response.ok) throw new Error(`PluginConfig HTTP ${response.status} ${response.statusText || ""}`.trim());
    const payload = await response.json();
    if (payload?.schemaVersion !== required.schemaVersion || payload?.dataVersion !== required.requiredDataVersion) {
      throw new Error(`PluginConfig incompatibile: schema ${payload?.schemaVersion}, dati ${payload?.dataVersion}.`);
    }
    validateTicketingConfig(payload);
    const stored = {
      schemaVersion: payload.schemaVersion,
      dataVersion: payload.dataVersion,
      downloadedAt: new Date().toISOString(),
      payload
    };
    localStorage.setItem(OPERATIONAL_CONFIG_KEY, JSON.stringify(stored));
    refreshOperationalConfigDiagnostics(stored, "scaricata");
    developerConfig = payload;
    return payload;
  }

  function normalizeEmail(value) {
    return String(value || "").trim().toLowerCase();
  }

  function evaluateDeveloperToolsAccess(config) {
    const outlookEmail = normalizeEmail(Office.context?.mailbox?.userProfile?.emailAddress);
    const operatorEntry = Object.entries(config?.operators || {}).find(([, candidate]) =>
      candidate?.enabled !== false && normalizeEmail(candidate?.email) === outlookEmail
    );
    const operator = operatorEntry?.[1] || null;
    const role = Object.values(config?.roles || {}).find((candidate) =>
      String(candidate?.role || "").trim().toLowerCase() === String(operator?.role || "").trim().toLowerCase()
    );
    const priority = Number(role?.priority);
    return {
      allowed: Boolean(operator && Number.isFinite(priority) && priority <= 1),
      operator,
      operatorId: operatorEntry?.[0] || "",
      role,
      outlookEmail
    };
  }

  function refreshDeveloperToolsAccess(config) {
    const access = evaluateDeveloperToolsAccess(config);
    developerConfig = config || null;
    developerToolsAllowed = access.allowed;
    developerCurrentOperatorId = access.operatorId;
    const toggle = byId("email-details-toggle");
    const panel = byId("email-details-panel");
    if (toggle) toggle.hidden = !developerToolsAllowed;
    if (!developerToolsAllowed && panel) {
      panel.hidden = true;
      toggle?.setAttribute("aria-expanded", "false");
    }
    populateDeveloperForms(config, access.operator?.email || "");
    if (getDeveloperAction() === "ManageCategories") void refreshOfficeCategoryForm({ preserveSelection: true });
    populateGraphCategoryForm(config);
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
    try {
      if (cached?.schemaVersion === required.schemaVersion &&
          cached?.dataVersion === required.requiredDataVersion && cached?.payload) {
        validateTicketingConfig(cached.payload);
        refreshOperationalConfigDiagnostics(cached, "cache valida");
        refreshDeveloperToolsAccess(cached.payload);
        return cached.payload;
      }
    } catch (error) {
      window.WorkflowHealth?.recordDiagnostic?.("config_cache_invalid", { message: error?.message || String(error) });
    }

    refreshDeveloperToolsAccess(null);
    if (!window.WorkflowHealth?.getAccessTokenSilently) return null;
    try {
      const authentication = await window.WorkflowHealth.getAccessTokenSilently();
      if (!authentication?.accessToken) throw new Error("Token di accesso non disponibile per inizializzare PluginConfig.");
      const config = await ensureOperationalConfig(authentication.accessToken);
      refreshDeveloperToolsAccess(config);
      return config;
    } catch (error) {
      refreshDeveloperToolsAccess(null);
      refreshOperationalConfigDiagnostics(cached, `errore: ${error?.message || String(error)}`);
      window.WorkflowHealth?.recordDiagnostic?.("config_initialize_error", { message: error?.message || String(error) });
      console.error("Inizializzazione PluginConfig non riuscita:", error);
      return null;
    } finally {
      refreshAuthDiagnostics();
    }
  }

  const DEVELOPER_ENDPOINTS = Object.freeze({
    GetRecordByConversationId: "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/GetRecordByConversationId",
    GetRecordById: "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/GetRecordById",
    SearchTicketRecords: "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/SearchTicketRecords",
    createTicket: "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/ManageTicket",
    editTicket: "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/ManageTicket",
    ForwardConversation: "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/ForwardConversation",
    SetMessageCategories: "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/SetMessageCategories"
  });
  const TICKET_NOTES_MAX_LENGTH = 2000;
  const TICKET_FIELDS = [
    ["id", "ID item ticket"], ["parentItemId", "Parent Item ID"], ["ticketId", "Ticket ID"],
    ["ampReference", "Riferimento AMP"], ["messageId", "Message ID"], ["operator", "Operatore"],
    ["status", "Stato"], ["agency", "Agenzia"], ["shortDescription", "Descrizione breve"],
    ["description", "Descrizione"], ["notes", "Note"], ["resolutionNotes", "Note di risoluzione"],
    ["elapsedTime", "Tempo trascorso"], ["created", "Creato"], ["modified", "Modificato"]
  ];
  const EDIT_FIELDS = [
    { key: "operator", id: "operator", nullable: false },
    { key: "status", id: "status", nullable: false },
    { key: "ampReference", id: "amp-reference", nullable: true },
    { key: "agency", id: "agency", nullable: true },
    { key: "shortDescription", id: "short-description", nullable: true },
    { key: "description", id: "description", nullable: true },
    { key: "notes", id: "notes", nullable: true },
    { key: "resolutionNotes", id: "resolution-notes", nullable: true },
    { key: "elapsedTime", id: "elapsed-time", nullable: true }
  ];

  function getDeveloperAction() {
    return byId("developer-action-select")?.value || "GetRecordByConversationId";
  }

  function resetTicketResult() {
    const panel = byId("developer-ticket-result");
    if (panel) panel.hidden = true;
    developerResultView = "fields";
    ["developer-result-action", "developer-ticket-status", "developer-ticket-auth-mode", "developer-ticket-response"].forEach((id) => setText(id, "", ""));
    byId("developer-ticket-fields")?.replaceChildren();
  }

  function setSelectOptions(select, options, preferredValue = "", preserve = false) {
    if (!select) return;
    const previous = preserve ? String(select.value || "") : "";
    select.replaceChildren();
    options.forEach(({ value, label }) => {
      const option = document.createElement("option");
      option.value = String(value);
      option.textContent = String(label);
      select.appendChild(option);
    });
    const wanted = previous || String(preferredValue || "");
    if (wanted && options.some((entry) => String(entry.value) === wanted)) select.value = wanted;
  }

  function setMultiSelectOptions(select, options, preserveSelection = true) {
    if (!select) return;
    const selected = preserveSelection
      ? new Set(Array.from(select.selectedOptions || []).map((option) => option.value))
      : new Set();
    select.replaceChildren();
    options.forEach(({ value, label, className = "" }) => {
      const option = document.createElement("option");
      option.value = String(value);
      option.textContent = String(label);
      option.selected = selected.has(option.value);
      if (className) option.className = className;
      select.appendChild(option);
    });
  }

  function getOperatorOptions(config = developerConfig) {
    return Object.entries(config?.operators || {})
      .filter(([, operator]) => operator?.enabled !== false && String(operator?.email || "").trim())
      .map(([id, operator]) => ({
        value: String(operator.email).trim(),
        label: `${operator?.displayName || id} · ${operator.email} · ${id}`
      }));
  }

  function getAgencyOptions(config = developerConfig, includeEmpty = false) {
    const options = Object.entries(config?.agencies || {})
      .filter(([, agency]) => agency?.enabled !== false)
      .map(([id, agency]) => ({
        value: String(agency?.displayName || id).trim(),
        label: `${agency?.displayName || id}${agency?.shortName ? ` · ${agency.shortName}` : ""} · ${id}`
      }));
    return includeEmpty ? [{ value: "", label: "— nessuna / rimuovi —" }, ...options] : options;
  }

  function getStatusOptions(config = developerConfig) {
    return (config?.ticketing?.statuses || [])
      .map((status) => String(status || "").trim()).filter(Boolean)
      .map((status) => ({ value: status, label: status }));
  }

  function getForwardTargetOptions(config = developerConfig) {
    return Object.entries(config?.forwardingTargets || {}).map(([id, target]) => ({
      value: id,
      label: `${target?.name || id}${target?.code ? ` · ${target.code}` : ""}`
    }));
  }

  function populateDeveloperForms(config = developerConfig, currentOperatorEmail = "") {
    const operators = getOperatorOptions(config);
    const statuses = getStatusOptions(config);
    const agencies = getAgencyOptions(config, true);
    const email = currentOperatorEmail || getCurrentOperatorEmail(config);
    setSelectOptions(byId("developer-create-operator"), operators, email, true);
    setSelectOptions(byId("developer-edit-caller"), operators, email, true);
    setSelectOptions(
      byId("developer-edit-operator"),
      [{ value: "", label: "— seleziona un operatore —" }, ...operators],
      "",
      true
    );
    setSelectOptions(byId("developer-create-status"), statuses, statuses[0]?.value || "", true);
    setSelectOptions(byId("developer-edit-status"), statuses, statuses[0]?.value || "", true);
    setSelectOptions(byId("developer-create-agency"), agencies, "", true);
    setSelectOptions(byId("developer-edit-agency"), agencies, "", true);
    setMultiSelectOptions(byId("developer-search-operators"), getOperatorOptions(config), true);
    setMultiSelectOptions(byId("developer-search-statuses"), statuses, true);
    setMultiSelectOptions(byId("developer-search-agencies"), getAgencyOptions(config), true);
    setSelectOptions(byId("developer-forward-target"), getForwardTargetOptions(config), "", true);

    const status = byId("developer-create-config-status");
    if (status) {
      status.textContent = config ? "Valori caricati dalla PluginConfig." : "Configurazione operativa non disponibile.";
      status.className = `developer-create-help${config ? "" : " status-error"}`;
    }
    populateGraphCategoryForm(config);
  }

  function getCurrentOperatorEmail(config = developerConfig) {
    const id = developerCurrentOperatorId;
    return String(config?.operators?.[id]?.email || Office.context?.mailbox?.userProfile?.emailAddress || "").trim();
  }

  function getM365AdminCenterMailbox(config = developerConfig) {
    return String(config?.ticketing?.email || "").trim();
  }

  function prefillCreateFromMessage() {
    const values = {
      "developer-create-message-id": currentMessageContext?.itemId || "",
      "developer-create-short-description": currentMessageContext?.subject || "",
      "developer-forward-conversation-id": currentMessageContext?.normalizedConversationId || "",
      "developer-graph-message-id": currentMessageContext?.itemId || ""
    };
    Object.entries(values).forEach(([id, value]) => {
      const input = byId(id);
      if (input && !input.value) input.value = value;
    });
  }

  function updateDeveloperActionUi() {
    const action = getDeveloperAction();
    const visibility = {
      "developer-get-form": action === "GetRecordByConversationId" || action === "GetRecordById",
      "developer-create-form": action === "createTicket",
      "developer-edit-form": action === "editTicket",
      "developer-search-form": action === "SearchTicketRecords",
      "developer-forward-form": action === "ForwardConversation",
      "developer-category-form": action === "ManageCategories",
      "developer-graph-category-form": action === "SetMessageCategories"
    };
    Object.entries(visibility).forEach(([id, visible]) => {
      const element = byId(id);
      if (element) element.hidden = !visible;
    });

    const label = byId("developer-action-input-label");
    const input = byId("developer-action-input");
    const modeRow = byId("developer-conversation-mode-row");
    if (modeRow) modeRow.hidden = action !== "GetRecordByConversationId";
    if (label && input) {
      if (action === "GetRecordById") {
        label.textContent = "Item ID";
        input.placeholder = "es. 9840";
        input.value = developerLastItemId || input.value || "";
      } else if (action === "GetRecordByConversationId") {
        label.textContent = "Conversation ID";
        input.placeholder = "Conversation ID del messaggio";
        input.value = currentMessageContext?.normalizedConversationId || input.value || "";
      }
    }

    if (action === "editTicket" && developerLastItemId) {
      const ticketId = byId("developer-edit-ticket-id");
      if (ticketId && !ticketId.value) ticketId.value = `SHP-${developerLastItemId}-1`;
    }

    prefillCreateFromMessage();
    populateDeveloperForms(developerConfig, getCurrentOperatorEmail());
    if (action === "ManageCategories") void refreshOfficeCategoryForm();
    if (action === "SetMessageCategories") populateGraphCategoryForm();
  }

  function setTicketButtonLoading(value) {
    const button = byId("developer-execute-button");
    if (button) {
      button.disabled = value;
      button.textContent = value ? "Esecuzione..." : "Esegui";
    }
  }

  function addDeveloperField(container, label, value) {
    const card = document.createElement("div");
    card.className = "developer-field-card";
    if (["Descrizione", "Note", "Note di risoluzione", "Campi", "Filtri applicati"].includes(label)) {
      card.classList.add("developer-log-field");
    }
    const name = document.createElement("span");
    name.textContent = label;
    const content = document.createElement("strong");
    let displayed = value;
    if (value && typeof value === "object") displayed = JSON.stringify(value, null, 2);
    content.textContent = displayed === undefined || displayed === null || displayed === "" ? "Non disponibile" : String(displayed);
    card.append(name, content);
    container.appendChild(card);
  }

  function getRecordResultEntries(parsed) {
    if (Array.isArray(parsed)) return parsed;
    if (Array.isArray(parsed?.items)) return parsed.items;
    if (Array.isArray(parsed?.records)) return parsed.records;
    if (parsed?.item && typeof parsed.item === "object") return [parsed.item];
    if (parsed?.fields && typeof parsed.fields === "object") return [{ itemId: parsed.itemId, fields: parsed.fields }];
    return [];
  }

  function getRecordItemId(record) {
    return record?.itemId ?? record?.id ?? record?.fields?.id ?? record?.fields?.ID ?? "";
  }

  function renderRecord(container, record) {
    const fields = record?.fields && typeof record.fields === "object" ? record.fields : {};
    addDeveloperField(container, "Item ID", getRecordItemId(record));
    const preferred = [
      "Subject", "From", "To", "CC", "BCC", "ReceivedDate", "MessageId",
      "Conversation_x0020_ID", "ConversationId", "ParentItemId", "HasAttachment",
      "Importance", "Status", "Created", "Modified"
    ];
    const rendered = new Set();
    preferred.forEach((key) => {
      if (Object.prototype.hasOwnProperty.call(fields, key)) {
        addDeveloperField(container, key, fields[key]);
        rendered.add(key);
      }
    });
    Object.keys(fields).filter((key) => !rendered.has(key)).sort((a, b) => a.localeCompare(b, "it"))
      .forEach((key) => addDeveloperField(container, key, fields[key]));
  }

  function renderRecordResultFields(parsed) {
    const container = byId("developer-ticket-fields");
    if (!container) return;
    container.replaceChildren();
    container.className = "developer-fields";
    addDeveloperField(container, "Esito", parsed?.status);
    addDeveloperField(container, "Trovato", parsed?.found === false ? "No" : "Sì");
    const records = getRecordResultEntries(parsed);
    addDeveloperField(container, "Record restituiti", records.length);
    if (!records.length) return;
    if (records.length === 1) {
      renderRecord(container, records[0]);
      return;
    }
    const selectorBox = document.createElement("div");
    selectorBox.className = "developer-ticket-selector";
    const label = document.createElement("label");
    label.textContent = "Record SharePoint";
    const select = document.createElement("select");
    records.forEach((record, index) => {
      const option = document.createElement("option");
      option.value = String(index);
      option.textContent = getRecordItemId(record) ? `Item ${getRecordItemId(record)}` : `Record ${index + 1}`;
      select.appendChild(option);
    });
    const host = document.createElement("div");
    const redraw = () => {
      host.replaceChildren();
      host.className = "developer-fields";
      renderRecord(host, records[Number(select.value) || 0]);
    };
    select.addEventListener("change", redraw);
    selectorBox.append(label, select);
    container.append(selectorBox, host);
    redraw();
  }

  function renderTicket(container, ticket) {
    if (!ticket || typeof ticket !== "object") return;
    TICKET_FIELDS.forEach(([key, label]) => addDeveloperField(container, label, ticket[key]));
  }

  function renderTicketResultFields(parsed, action) {
    const container = byId("developer-ticket-fields");
    if (!container) return;
    container.replaceChildren();
    container.className = "developer-fields";
    addDeveloperField(container, "Operazione", parsed?.action || action);
    addDeveloperField(container, "Esito", parsed?.status || parsed?.message);
    if (parsed?.reason) addDeveloperField(container, "Motivo", parsed.reason);
    if (parsed?.operator !== undefined) addDeveloperField(container, "Operatore ticket", parsed.operator);
    renderTicket(container, parsed?.ticket);
    if (parsed?.message && parsed?.status !== "ok") addDeveloperField(container, "Errore", parsed.message);
  }

  function renderSearchTicketRecordsFields(parsed) {
    const container = byId("developer-ticket-fields");
    if (!container) return;
    container.replaceChildren();
    container.className = "developer-fields";
    addDeveloperField(container, "Esito", parsed?.status);
    addDeveloperField(container, "Totale ticket", parsed?.totalCount ?? 0);
    addDeveloperField(container, "Ticket restituiti", parsed?.returnedCount ?? 0);
    addDeveloperField(container, "Troncato", parsed?.truncated ? "Sì" : "No");
    addDeveloperField(container, "Tempo", parsed?.elapsedMs == null ? null : `${parsed.elapsedMs} ms`);
    addDeveloperField(container, "Filtri applicati", parsed?.filters || {});
    const records = Array.isArray(parsed?.records) ? parsed.records : [];
    if (!records.length) return;
    const selectorBox = document.createElement("div");
    selectorBox.className = "developer-ticket-selector";
    const label = document.createElement("label");
    label.textContent = "Ticket";
    const select = document.createElement("select");
    records.forEach((ticket, index) => {
      const option = document.createElement("option");
      option.value = String(index);
      option.textContent = [ticket?.ticketId || `Ticket ${index + 1}`, ticket?.status, ticket?.operator].filter(Boolean).join(" · ");
      select.appendChild(option);
    });
    const host = document.createElement("div");
    const redraw = () => {
      host.replaceChildren();
      host.className = "developer-fields";
      renderTicket(host, records[Number(select.value) || 0]);
    };
    select.addEventListener("change", redraw);
    selectorBox.append(label, select);
    container.append(selectorBox, host);
    redraw();
  }

  function renderForwardResultFields(parsed) {
    const container = byId("developer-ticket-fields");
    if (!container) return;
    container.replaceChildren();
    container.className = "developer-fields";
    addDeveloperField(container, "Esito", parsed?.status || parsed?.message);
    addDeveloperField(container, "Conversation ID", parsed?.conversationId);
    addDeveloperField(container, "Destinazione", parsed?.groupName);
    addDeveloperField(container, "Item ID", parsed?.itemId);
    addDeveloperField(container, "Campi", parsed?.fields);
    if (parsed?.message && parsed?.status !== "ok") addDeveloperField(container, "Errore", parsed.message);
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
    addDeveloperField(container, "Operazione", action);
    addDeveloperField(container, "Action", parsed?.action);
    addDeveloperField(container, "Modalità", parsed?.mode);
    addDeveloperField(container, "Aggiornato", parsed?.updated);
    addDeveloperField(container, "Mail", parsed?.mail);
    addDeveloperField(container, "Message ID", parsed?.messageId);
    addDeveloperField(container, "Categorie", formatCategoryNameList(parsed?.categories));
    addDeveloperField(container, "Categorie prima", formatCategoryNameList(parsed?.categoriesBefore));
    addDeveloperField(container, "Categorie dopo", formatCategoryNameList(parsed?.categoriesAfter));
    addDeveloperField(container, "Master Category List", formatCategoryNameList(parsed?.masterCategories));
    if (parsed?.message && parsed?.status !== "ok") addDeveloperField(container, "Errore", parsed.message);
  }

  function isRecordDeveloperAction(action) {
    return action === "GetRecordByConversationId" || action === "GetRecordById";
  }
  function isTicketDeveloperAction(action) {
    return action === "createTicket" || action === "editTicket";
  }
  function isCategoryDeveloperAction(action) {
    return action === "ManageCategories" || action === "SetMessageCategories";
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
      response.textContent = JSON.stringify(parsed, null, 2);
    } catch {
      response.textContent = result.body || "(Risposta vuota)";
    }

    const returnedRecord = isRecordDeveloperAction(action) ? getRecordResultEntries(parsed)[0] : null;
    const returnedItemId = returnedRecord ? getRecordItemId(returnedRecord) : parsed?.itemId;
    if (returnedItemId) {
      developerLastItemId = String(returnedItemId);
      const ticketId = byId("developer-edit-ticket-id");
      if (ticketId && !ticketId.value) ticketId.value = `SHP-${developerLastItemId}-1`;
    }
    if (action === "GetRecordByConversationId" && getConversationRecordMode() === "first" && returnedItemId) {
      const parent = byId("developer-create-parent-item-id");
      if (parent) parent.value = String(returnedItemId);
    }

    const hasFields = Boolean(parsed && typeof parsed === "object");
    developerResultView = hasFields ? "fields" : "json";
    if (isRecordDeveloperAction(action) && parsed) renderRecordResultFields(parsed);
    else if (action === "SearchTicketRecords" && parsed) renderSearchTicketRecordsFields(parsed);
    else if (isTicketDeveloperAction(action) && parsed) renderTicketResultFields(parsed, action);
    else if (action === "ForwardConversation" && parsed) renderForwardResultFields(parsed);
    else if (isCategoryDeveloperAction(action) && parsed) renderCategoryFields(parsed, action);
    else byId("developer-ticket-fields")?.replaceChildren();
    applyDeveloperResultView();
  }

  function requiredFormValue(id, label) {
    const value = String(byId(id)?.value || "").trim();
    if (!value) throw new Error(`${label} è obbligatorio.`);
    return value;
  }

  function optionalTextValue(id) {
    const value = String(byId(id)?.value || "").trim();
    return value || null;
  }

  function assertPositiveInteger(value, label) {
    if (!/^\d+$/.test(String(value)) || Number(value) <= 0) throw new Error(`${label} deve essere un intero positivo.`);
    return Number(value);
  }

  function collectCreateTicketBody() {
    validateTicketingConfig(developerConfig);
    const parentItemId = requiredFormValue("developer-create-parent-item-id", "Parent Item ID");
    assertPositiveInteger(parentItemId, "Parent Item ID");
    const status = requiredFormValue("developer-create-status", "Stato");
    if (!developerConfig.ticketing.statuses.includes(status)) throw new Error(`Stato non ammesso: ${status}.`);
    const notes = String(byId("developer-create-notes")?.value || "");
    const resolutionNotes = String(byId("developer-create-resolution-notes")?.value || "");
    if (notes.length > TICKET_NOTES_MAX_LENGTH || resolutionNotes.length > TICKET_NOTES_MAX_LENGTH) {
      throw new Error(`Note e note di risoluzione non possono superare ${TICKET_NOTES_MAX_LENGTH} caratteri.`);
    }
    const body = {
      action: "createTicket",
      parentItemId,
      messageId: requiredFormValue("developer-create-message-id", "Message ID"),
      operator: requiredFormValue("developer-create-operator", "Operatore"),
      status
    };
    const optional = [
      ["ampReference", "developer-create-amp-reference"],
      ["agency", "developer-create-agency"],
      ["shortDescription", "developer-create-short-description"],
      ["description", "developer-create-description"],
      ["notes", "developer-create-notes"],
      ["resolutionNotes", "developer-create-resolution-notes"],
      ["elapsedTime", "developer-create-elapsed-time"]
    ];
    optional.forEach(([key, id]) => {
      const value = String(byId(id)?.value || "").trim();
      if (value) body[key] = value;
    });
    return body;
  }

  function collectEditTicketBody() {
    validateTicketingConfig(developerConfig);
    const body = {
      action: "editTicket",
      ticketId: requiredFormValue("developer-edit-ticket-id", "Ticket ID"),
      caller: requiredFormValue("developer-edit-caller", "Caller"),
      force: Boolean(byId("developer-edit-force")?.checked)
    };
    let changed = 0;
    EDIT_FIELDS.forEach(({ key, id, nullable }) => {
      if (!byId(`developer-edit-use-${id}`)?.checked) return;
      const value = String(byId(`developer-edit-${id}`)?.value || "").trim();
      if (!nullable && !value) throw new Error(`${key} non può essere vuoto.`);
      if (key === "status" && !developerConfig.ticketing.statuses.includes(value)) throw new Error(`Stato non ammesso: ${value}.`);
      if ((key === "notes" || key === "resolutionNotes") && value.length > TICKET_NOTES_MAX_LENGTH) {
        throw new Error(`${key} non può superare ${TICKET_NOTES_MAX_LENGTH} caratteri.`);
      }
      body[key] = value || null;
      changed += 1;
    });
    if (!changed) throw new Error("Attiva Aggiorna per almeno un campo modificabile.");
    return body;
  }

  function parseListInput(id) {
    const values = String(byId(id)?.value || "")
      .split(/[\n,]+/).map((value) => value.trim()).filter(Boolean);
    return Array.from(new Set(values));
  }

  function getSelectedValues(id) {
    return Array.from(byId(id)?.selectedOptions || []).map((option) => option.value);
  }

  function optionalDateTimeToIso(id, label) {
    const value = String(byId(id)?.value || "").trim();
    if (!value) return null;
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) throw new Error(`${label} non è una data valida.`);
    return date.toISOString();
  }

  function collectSearchTicketRecordsBody() {
    validateTicketingConfig(developerConfig);
    const body = {};
    const families = [
      ["ticketId", "ticketIds", parseListInput("developer-search-ticket-ids")],
      ["ampReference", "ampReferences", parseListInput("developer-search-amp-references")],
      ["messageId", "messageIds", parseListInput("developer-search-message-ids")],
      ["operator", "operators", getSelectedValues("developer-search-operators")],
      ["status", "statuses", getSelectedValues("developer-search-statuses")],
      ["agency", "agencies", getSelectedValues("developer-search-agencies")]
    ];
    families.forEach(([singular, plural, values]) => {
      if (values.length === 1) body[singular] = values[0];
      else if (values.length > 1) body[plural] = values;
    });
    const parent = String(byId("developer-search-parent-item-id")?.value || "").trim();
    if (parent) body.parentItemId = assertPositiveInteger(parent, "Parent Item ID");
    const dates = [
      ["createdFrom", "developer-search-created-from", "Created da"],
      ["createdTo", "developer-search-created-to", "Created a"],
      ["modifiedFrom", "developer-search-modified-from", "Modified da"],
      ["modifiedTo", "developer-search-modified-to", "Modified a"]
    ];
    dates.forEach(([key, id, label]) => {
      const value = optionalDateTimeToIso(id, label);
      if (value) body[key] = value;
    });
    if (body.createdFrom && body.createdTo && Date.parse(body.createdFrom) > Date.parse(body.createdTo)) throw new Error("Created da non può essere successivo a Created a.");
    if (body.modifiedFrom && body.modifiedTo && Date.parse(body.modifiedFrom) > Date.parse(body.modifiedTo)) throw new Error("Modified da non può essere successivo a Modified a.");
    const resultSize = String(byId("developer-search-result-size")?.value || "").trim();
    if (resultSize) {
      const parsed = Number(resultSize);
      if (!Number.isInteger(parsed) || parsed < 1 || parsed > 1000) throw new Error("Result size deve essere un intero da 1 a 1000.");
      body.resultSize = parsed;
    }
    return body;
  }

  function collectForwardConversationBody() {
    validateTicketingConfig(developerConfig);
    const targetId = requiredFormValue("developer-forward-target", "Destinazione");
    const target = developerConfig?.forwardingTargets?.[targetId];
    const groupName = String(target?.name || "").trim();
    if (!groupName) throw new Error(`Destinazione forwarding non valida: ${targetId}.`);
    return {
      conversationId: requiredFormValue("developer-forward-conversation-id", "Conversation ID"),
      groupName
    };
  }

  function getConversationRecordMode() {
    const mode = String(byId("developer-conversation-mode")?.value || "first").trim().toLowerCase();
    if (!["all", "first", "last"].includes(mode)) throw new Error("Mode non valido: usa all, first oppure last.");
    return mode;
  }

  function updateCharacterCount(inputId, counterId) {
    const length = String(byId(inputId)?.value || "").length;
    const counter = byId(counterId);
    if (counter) {
      counter.textContent = `${length}/${TICKET_NOTES_MAX_LENGTH}`;
      counter.className = length > TICKET_NOTES_MAX_LENGTH ? "status-error" : "";
    }
  }

  function updateEditFieldControl(id) {
    const checkbox = byId(`developer-edit-use-${id}`);
    const control = byId(`developer-edit-${id}`);
    if (control) control.disabled = !checkbox?.checked;
  }

  async function executeOfficeCategoryAction() {
    const action = String(byId("developer-category-action")?.value || "get");
    const item = Office.context?.mailbox?.item;
    if (!item || !currentMessageContext) throw new Error("Seleziona un messaggio Outlook.");
    const categories = uniqueCategoryNames(getSelectedValues("developer-category-categories"));
    const replace = Boolean(byId("developer-category-replace")?.checked);
    const clear = Boolean(byId("developer-category-clear")?.checked);
    const beforeRaw = await getCategories(item);
    const masterRaw = await getMasterCategories();
    const before = uniqueCategoryNames(beforeRaw);
    const master = uniqueCategoryNames(masterRaw);
    const beforeMap = new Map(before.map((value) => [categoryKey(value), value]));
    const masterMap = new Map(master.map((value) => [categoryKey(value), value]));
    if (action === "add") {
      if (!categories.length) throw new Error("Seleziona almeno una categoria da aggiungere.");
      const invalid = categories.filter((value) => !masterMap.has(categoryKey(value)));
      if (invalid.length) throw new Error(`Categorie non presenti nella Master Category List: ${invalid.join(", ")}.`);
      const canonical = categories.map((value) => masterMap.get(categoryKey(value)));
      if (replace) {
        if (before.length) await removeCategories(item, before);
        if (canonical.length) await addCategories(item, canonical);
      } else {
        await addCategories(item, canonical.filter((value) => !beforeMap.has(categoryKey(value))));
      }
    } else if (action === "remove") {
      const selected = clear ? before : categories.map((value) => beforeMap.get(categoryKey(value))).filter(Boolean);
      if (!clear && !selected.length) throw new Error("Seleziona almeno una categoria da rimuovere oppure attiva clear.");
      if (selected.length) await removeCategories(item, selected);
    } else if (action !== "get") {
      throw new Error("Action Office.js non valida.");
    }
    const afterRaw = await getCategories(item);
    const after = uniqueCategoryNames(afterRaw);
    updateCurrentMessageCategories(afterRaw);
    return {
      ok: true,
      status: 200,
      statusText: "OK",
      authenticationMode: "Office.js",
      body: JSON.stringify({ status: "ok", action, categoriesBefore: before, categoriesAfter: after, masterCategories: master })
    };
  }

  function updateOfficeCategoryControls() {
    const action = String(byId("developer-category-action")?.value || "get");
    const row = byId("developer-category-categories-row");
    const replaceRow = byId("developer-category-replace-row");
    const clearRow = byId("developer-category-clear-row");
    if (row) row.hidden = action === "get";
    if (replaceRow) replaceRow.hidden = action !== "add";
    if (clearRow) clearRow.hidden = action !== "remove";
    const select = byId("developer-category-categories");
    if (select) select.disabled = action === "get" || (action === "remove" && Boolean(byId("developer-category-clear")?.checked));
  }

  async function refreshOfficeCategoryForm({ preserveSelection = false } = {}) {
    const sequence = ++officeCategoryLoadSequence;
    const status = byId("developer-category-available");
    try {
      const item = Office.context?.mailbox?.item;
      if (!item) throw new Error("Seleziona un messaggio Outlook.");
      const [appliedRaw, masterRaw] = await Promise.all([getCategories(item), getMasterCategories()]);
      if (sequence !== officeCategoryLoadSequence) return;
      const applied = uniqueCategoryNames(appliedRaw);
      const master = uniqueCategoryNames(masterRaw);
      const action = String(byId("developer-category-action")?.value || "get");
      const options = action === "add"
        ? masterRaw.map((category) => ({ value: String(category?.displayName || category), label: `${category?.displayName || category}${category?.color ? ` · ${category.color}` : ""}` }))
        : action === "remove" ? applied.map((value) => ({ value, label: value })) : [];
      setMultiSelectOptions(byId("developer-category-categories"), options, preserveSelection);
      updateOfficeCategoryControls();
      if (status) status.textContent = `Applicate: ${applied.length ? applied.join(", ") : "nessuna"}. Master: ${master.length}.`;
    } catch (error) {
      if (sequence !== officeCategoryLoadSequence) return;
      if (status) {
        status.textContent = `Lettura Office.js non riuscita: ${error?.message || String(error)}`;
        status.className = "developer-create-help developer-category-summary status-error";
      }
    }
  }

  function graphCategoryStateMatches(mail, messageId) {
    return developerGraphCategoryState.mail === String(mail || "") && developerGraphCategoryState.messageId === String(messageId || "");
  }

  function resetGraphCategoryState() {
    graphCategoryLoadSequence += 1;
    developerGraphCategoryState = {
      mail: String(byId("developer-graph-mail")?.value || "").trim(),
      messageId: String(byId("developer-graph-message-id")?.value || "").trim(),
      categories: [], loaded: false, loading: false, error: ""
    };
  }

  function renderGraphCategoryStatus(config = developerConfig) {
    const status = byId("developer-graph-status");
    if (!status) return;
    const mail = String(byId("developer-graph-mail")?.value || "").trim();
    if (!mail) {
      status.textContent = "PluginConfig.ticketing.email è mancante o vuota.";
      status.className = "developer-create-help developer-category-summary status-error";
      return;
    }
    if (developerGraphCategoryState.loading) status.textContent = "Lettura Graph in corso…";
    else if (developerGraphCategoryState.error) status.textContent = `Lettura Graph non riuscita: ${developerGraphCategoryState.error}`;
    else if (developerGraphCategoryState.loaded) status.textContent = `Categorie lette da Graph: ${developerGraphCategoryState.categories.length}.`;
    else status.textContent = `Mailbox caricata. Categorie configurate: ${collectConfiguredCategories(config).length}.`;
    status.className = `developer-create-help developer-category-summary${developerGraphCategoryState.error ? " status-error" : ""}`;
  }

  function populateGraphCategoryForm(config = developerConfig) {
    const mail = byId("developer-graph-mail");
    if (mail) mail.value = getM365AdminCenterMailbox(config);
    const message = byId("developer-graph-message-id");
    if (message && !message.value) message.value = currentMessageContext?.itemId || "";
    updateGraphCategoryControls();
  }

  function populateGraphAddCategories(config = developerConfig, preserveSelection = true) {
    setMultiSelectOptions(byId("developer-graph-categories"), collectConfiguredCategories(config).map((value) => ({ value, label: value })), preserveSelection);
  }

  function populateGraphRemoveCategories(config = developerConfig, preserveSelection = true) {
    const configured = new Set(collectConfiguredCategories(config).map(categoryKey));
    setMultiSelectOptions(byId("developer-graph-categories"), developerGraphCategoryState.categories.map((value) => ({
      value,
      label: `${value} · ${configured.has(categoryKey(value)) ? "PluginConfig" : "non presente in PluginConfig"}`,
      className: configured.has(categoryKey(value)) ? "" : "category-not-configured"
    })), preserveSelection);
  }

  async function loadGraphCategoriesForRemove({ force = false } = {}) {
    const mail = requiredFormValue("developer-graph-mail", "Mail Function da PluginConfig");
    const messageId = requiredFormValue("developer-graph-message-id", "Message ID");
    if (!force && developerGraphCategoryState.loaded && graphCategoryStateMatches(mail, messageId)) return developerGraphCategoryState.categories;
    const sequence = ++graphCategoryLoadSequence;
    developerGraphCategoryState = { mail, messageId, categories: [], loaded: false, loading: true, error: "" };
    renderGraphCategoryStatus();
    try {
      const authentication = await window.WorkflowHealth.getAccessTokenSilently();
      if (!authentication?.accessToken) throw new Error("Token di accesso non disponibile.");
      const response = await fetch(DEVELOPER_ENDPOINTS.SetMessageCategories, {
        method: "POST",
        headers: { Authorization: `Bearer ${authentication.accessToken}`, Accept: "application/json", "Content-Type": "application/json" },
        body: JSON.stringify({ mail, messageId, action: "get" })
      });
      const text = await response.text();
      let parsed = {};
      try { parsed = JSON.parse(text || "{}"); } catch {}
      if (!response.ok) throw new Error(parsed?.message || `Graph get HTTP ${response.status}`);
      if (sequence !== graphCategoryLoadSequence) return [];
      developerGraphCategoryState = { mail, messageId, categories: uniqueCategoryNames(parsed?.categories), loaded: true, loading: false, error: "" };
      populateGraphRemoveCategories();
      renderGraphCategoryStatus();
      return developerGraphCategoryState.categories;
    } catch (error) {
      if (sequence !== graphCategoryLoadSequence) return [];
      developerGraphCategoryState = { mail, messageId, categories: [], loaded: false, loading: false, error: error?.message || String(error) };
      renderGraphCategoryStatus();
      throw error;
    }
  }

  function updateGraphCategoryControls() {
    const action = String(byId("developer-graph-action")?.value || "get");
    const isAdd = action === "add";
    const isRemove = action === "remove";
    const row = byId("developer-graph-categories-row");
    const replaceRow = byId("developer-graph-replace-row");
    const clearRow = byId("developer-graph-clear-row");
    const skipRow = byId("developer-graph-skip-final-get-row");
    const refresh = byId("developer-graph-refresh");
    if (row) row.hidden = action === "get";
    if (replaceRow) replaceRow.hidden = !isAdd;
    if (clearRow) clearRow.hidden = !isRemove;
    if (skipRow) skipRow.hidden = action === "get";
    if (refresh) refresh.hidden = !isRemove;
    if (isAdd) populateGraphAddCategories();
    else if (isRemove && developerGraphCategoryState.loaded) populateGraphRemoveCategories();
    else setMultiSelectOptions(byId("developer-graph-categories"), [], false);
    const select = byId("developer-graph-categories");
    if (select) select.disabled = action === "get" || (isRemove && Boolean(byId("developer-graph-clear")?.checked));
    renderGraphCategoryStatus();
  }

  function collectGraphCategoryBody() {
    const action = String(byId("developer-graph-action")?.value || "get");
    if (!["get", "add", "remove"].includes(action)) throw new Error("Action Graph non valida.");
    const body = {
      mail: requiredFormValue("developer-graph-mail", "Mail Function da PluginConfig"),
      messageId: requiredFormValue("developer-graph-message-id", "Message ID"),
      action
    };
    if (action === "get") return body;
    const clear = Boolean(byId("developer-graph-clear")?.checked);
    const replace = Boolean(byId("developer-graph-replace")?.checked);
    const categories = uniqueCategoryNames(getSelectedValues("developer-graph-categories"));
    if (action === "add") {
      if (!categories.length) throw new Error("Seleziona almeno una categoria.");
      body.categories = categories;
      body.replace = replace;
    } else {
      if (!clear && !categories.length) throw new Error("Seleziona almeno una categoria oppure attiva clear.");
      body.categories = clear ? [] : categories;
      body.clear = clear;
    }
    body.skipFinalGet = Boolean(byId("developer-graph-skip-final-get")?.checked);
    return body;
  }

  function updateGraphCategoryStateFromResponse(parsed, requestBody) {
    if (!parsed || !requestBody) return;
    const categories = parsed?.categories;
    if (!Array.isArray(categories)) {
      resetGraphCategoryState();
      renderGraphCategoryStatus();
      return;
    }
    developerGraphCategoryState = {
      mail: String(requestBody.mail || ""), messageId: String(requestBody.messageId || ""),
      categories: uniqueCategoryNames(categories), loaded: true, loading: false, error: ""
    };
    renderGraphCategoryStatus();
  }

  async function executeDeveloperAction() {
    const action = getDeveloperAction();
    const lookupValue = String(byId("developer-action-input")?.value || "").trim();
    resetTicketResult();
    setTicketButtonLoading(true);
    try {
      if (action === "ManageCategories") {
        const result = await executeOfficeCategoryAction();
        showTicketResult(result, action);
        await refreshOfficeCategoryForm();
        return;
      }
      if (isRecordDeveloperAction(action) && !lookupValue) throw new Error(action === "GetRecordById" ? "Inserisci un Item ID." : "Inserisci un Conversation ID.");
      const authentication = await window.WorkflowHealth.getAccessTokenSilently();
      if (!authentication?.accessToken) throw new Error("Token di accesso non disponibile.");
      const config = await ensureOperationalConfig(authentication.accessToken);
      refreshDeveloperToolsAccess(config);

      let url = DEVELOPER_ENDPOINTS[action];
      let request;
      let requestContext;
      if (isRecordDeveloperAction(action)) {
        const parameter = action === "GetRecordById" ? "itemId" : "conversationId";
        const query = new URLSearchParams({ [parameter]: lookupValue });
        if (action === "GetRecordByConversationId") query.set("mode", getConversationRecordMode());
        url = `${url}?${query.toString()}`;
        request = { method: "GET", headers: { Authorization: `Bearer ${authentication.accessToken}`, Accept: "application/json" } };
        requestContext = Object.fromEntries(query.entries());
      } else {
        let body;
        if (action === "createTicket") body = collectCreateTicketBody();
        else if (action === "editTicket") body = collectEditTicketBody();
        else if (action === "SearchTicketRecords") body = collectSearchTicketRecordsBody();
        else if (action === "ForwardConversation") body = collectForwardConversationBody();
        else if (action === "SetMessageCategories") {
          populateGraphCategoryForm(config);
          if (String(byId("developer-graph-action")?.value || "get") === "remove") await loadGraphCategoriesForRemove();
          body = collectGraphCategoryBody();
        } else throw new Error(`Azione non supportata: ${action}.`);
        requestContext = body;
        request = {
          method: "POST",
          headers: { Authorization: `Bearer ${authentication.accessToken}`, Accept: "application/json", "Content-Type": "application/json" },
          body: JSON.stringify(body)
        };
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
      if (action === "SetMessageCategories") updateGraphCategoryStateFromResponse(parsed, requestContext);
      if (action !== "SetMessageCategories") {
        const returnedRecord = isRecordDeveloperAction(action) ? getRecordResultEntries(parsed)[0] : null;
        currentSharePointData = {
          operation: action,
          conversationId: requestContext?.conversationId || parsed?.conversationId || null,
          itemId: getRecordItemId(returnedRecord) || parsed?.itemId || parsed?.ticket?.parentItemId || null,
          ok: result.ok,
          status: result.status,
          statusText: result.statusText,
          authenticationMode: result.authenticationMode,
          retrievedAt: new Date().toISOString(),
          request: requestContext,
          data: parsed,
          rawBody: parsed ? null : responseBody
        };
        publishSharePointData();
      }
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
    [
      ["developer-create-notes", "developer-create-notes-count"],
      ["developer-create-resolution-notes", "developer-create-resolution-notes-count"],
      ["developer-edit-notes", "developer-edit-notes-count"],
      ["developer-edit-resolution-notes", "developer-edit-resolution-notes-count"]
    ].forEach(([inputId, counterId]) => byId(inputId)?.addEventListener("input", () => updateCharacterCount(inputId, counterId)));
    EDIT_FIELDS.forEach(({ id }) => byId(`developer-edit-use-${id}`)?.addEventListener("change", () => updateEditFieldControl(id)));
    byId("developer-category-action")?.addEventListener("change", () => { updateOfficeCategoryControls(); void refreshOfficeCategoryForm(); });
    byId("developer-category-clear")?.addEventListener("change", updateOfficeCategoryControls);
    byId("developer-category-refresh")?.addEventListener("click", () => { void refreshOfficeCategoryForm({ preserveSelection: true }); });
    byId("developer-graph-action")?.addEventListener("change", () => {
      resetGraphCategoryState();
      updateGraphCategoryControls();
      if (String(byId("developer-graph-action")?.value || "get") === "remove") void loadGraphCategoriesForRemove({ force: true }).catch(() => {});
    });
    byId("developer-graph-clear")?.addEventListener("change", updateGraphCategoryControls);
    byId("developer-graph-refresh")?.addEventListener("click", () => { void loadGraphCategoriesForRemove({ force: true }).catch(() => {}); });
    byId("developer-graph-message-id")?.addEventListener("change", resetGraphCategoryState);
    byId(elementIds.idToggle)?.addEventListener("click", () => toggleIdentifierView("itemId"));
    byId(elementIds.conversationIdToggle)?.addEventListener("click", () => toggleIdentifierView("conversationId"));
    EDIT_FIELDS.forEach(({ id }) => updateEditFieldControl(id));
    [
      ["developer-create-notes", "developer-create-notes-count"],
      ["developer-create-resolution-notes", "developer-create-resolution-notes-count"],
      ["developer-edit-notes", "developer-edit-notes-count"],
      ["developer-edit-resolution-notes", "developer-edit-resolution-notes-count"]
    ].forEach(([inputId, counterId]) => updateCharacterCount(inputId, counterId));
    updateOfficeCategoryControls();
    updateGraphCategoryControls();
    updateDeveloperActionUi();
  }

  async function loadSelectedMessage() {
    const sequence = ++loadSequence;

    currentMessageIdentifiers = null;
    developerLastItemId = "";
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
      const messageDefaults = {
        "developer-create-parent-item-id": "",
        "developer-create-message-id": currentMessageContext?.itemId || "",
        "developer-create-short-description": currentMessageContext?.subject || "",
        "developer-create-description": "",
        "developer-edit-ticket-id": "",
        "developer-forward-conversation-id": currentMessageContext?.normalizedConversationId || "",
        "developer-graph-message-id": currentMessageContext?.itemId || ""
      };
      Object.entries(messageDefaults).forEach(([id, value]) => {
        const input = byId(id);
        if (input) input.value = value;
      });
      resetGraphCategoryState();
      updateGraphCategoryControls();
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
