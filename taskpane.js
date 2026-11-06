// Test-Plugin file version: 0.2-taskpane.js-devtools-5
const TEST_PLUGIN_BUILD = Object.freeze({
  application: '0.2',
  release: '0.2-devtools',
  files: Object.freeze({
    'taskpane.html': '0.2-taskpane.html-devtools-5',
    'taskpane.js': '0.2-taskpane.js-devtools-5',
    'taskpane.css': '0.2-taskpane.css-devtools-5',
    'auth-src.js': '0.2-auth-src.js-3',
    'auth-dialog.html': '0.2-auth-dialog.html-2',
    'auth-dialog-src.js': '0.2-auth-dialog-src.js-3',
    'auth-popup-redirect.html': '0.2-auth-popup-redirect.html-1',
    'auth-popup-redirect-src.js': '0.2-auth-popup-redirect-src.js-1',
  }),
  database: Object.freeze({ schemaVersion: 2, requiredDataVersion: 3 }),
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
  let currentSharePointData = null;
  let developerToolsAllowed = false;
  let developerLastItemId = "";
  let developerResultView = "fields";
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
    if (getDeveloperAction() !== "GetTicketByConversationId") return;
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
      `Gruppi: ${Object.keys(payload?.groups || {}).length}`
    ];
    setText("diagnostic-database", lines.join("\n"));
  }
  async function ensureOperationalConfig(accessToken) {
    const required = TEST_PLUGIN_BUILD.database;
    const cached = readOperationalConfigCache();
    window.WorkflowHealth?.recordDiagnostic?.("config_cache_read", { schemaVersion: cached?.schemaVersion ?? null, dataVersion: cached?.dataVersion ?? null });
    if (cached?.schemaVersion === required.schemaVersion && cached?.dataVersion === required.requiredDataVersion && cached?.payload) {
      window.WorkflowHealth?.recordDiagnostic?.("config_cache_valid", { dataVersion: cached.dataVersion });
      refreshOperationalConfigDiagnostics(cached, "cache valida");
      return cached.payload;
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
      const stored = { schemaVersion: payload.schemaVersion, dataVersion: payload.dataVersion, downloadedAt: new Date().toISOString(), payload };
      localStorage.setItem(OPERATIONAL_CONFIG_KEY, JSON.stringify(stored));
      window.WorkflowHealth?.recordDiagnostic?.("config_cache_saved", { schemaVersion: stored.schemaVersion, dataVersion: stored.dataVersion, agencies: Object.keys(payload.agencies || {}).length, operators: Object.keys(payload.operators || {}).length, groups: Object.keys(payload.groups || {}).length });
      refreshOperationalConfigDiagnostics(stored, "scaricata");
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
    developerToolsAllowed = access.allowed;
    const toggle = byId("email-details-toggle");
    const panel = byId("email-details-panel");
    if (toggle) toggle.hidden = !developerToolsAllowed;
    if (!developerToolsAllowed && panel) {
      panel.hidden = true;
      toggle?.setAttribute("aria-expanded", "false");
    }

    window.WorkflowHealth?.recordDiagnostic?.("developer_tools_access_evaluated", {
      allowed: developerToolsAllowed,
      operatorId: access.operator?.id || null,
      role: access.operator?.role || null,
      priority: Number.isFinite(Number(access.role?.priority)) ? Number(access.role.priority) : null,
      emailMatched: Boolean(access.operator)
    });
    return developerToolsAllowed;
  }

  async function initializeOperationalConfig() {
    const required = TEST_PLUGIN_BUILD.database;
    const cached = readOperationalConfigCache();
    const cacheIsValid = Boolean(
      cached?.schemaVersion === required.schemaVersion &&
      cached?.dataVersion === required.requiredDataVersion &&
      cached?.payload
    );

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

  const DEVELOPER_ENDPOINTS = Object.freeze({ GetTicketByConversationId: "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/GetTicketByConversationId", GetTicketById: "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/GetTicketById" });
  function resetTicketResult(){const panel=byId("developer-ticket-result");if(panel)panel.hidden=true;developerResultView="fields";["developer-result-action","developer-ticket-status","developer-ticket-auth-mode","developer-ticket-response"].forEach(id=>setText(id,"",""));byId("developer-ticket-fields")?.replaceChildren();}
  function getDeveloperAction(){return byId("developer-action-select")?.value||"GetTicketByConversationId";}
  function updateDeveloperActionUi(){const action=getDeveloperAction(),label=byId("developer-action-input-label"),input=byId("developer-action-input");if(!label||!input)return;if(action==="GetTicketById"){label.textContent="Item ID";input.placeholder="es. 9840";input.value=developerLastItemId||"";}else{label.textContent="Conversation ID";input.placeholder="Conversation ID del messaggio";input.value=currentMessageContext?.normalizedConversationId||"";}}
  function setTicketButtonLoading(value){const button=byId("developer-execute-button");if(button){button.disabled=value;button.textContent=value?"Esecuzione...":"Esegui";}}
  function parseAmpMetadata(value){if(!value)return{tickets:[],error:null};try{const parsed=typeof value==="string"?JSON.parse(value):value;return{tickets:Array.isArray(parsed?.tickets)?parsed.tickets:[],error:null};}catch(error){return{tickets:[],error:error?.message||String(error)};}}
  function findOperatorByLookupId(id){const config=readOperationalConfigCache()?.payload;return Object.entries(config?.operators||{}).find(([,op])=>String(op?.lookupId)===String(id))||null;}
  function addDeveloperField(container,label,value){const card=document.createElement("div");card.className="developer-field-card";const name=document.createElement("span");name.textContent=label;const content=document.createElement("strong");content.textContent=value==null||value===""?"Non disponibile":String(value);card.append(name,content);container.appendChild(card);}
  function renderSelectedAmpTicket(container,ticket){const details=document.createElement("div");details.className="developer-fields";addDeveloperField(details,"Ticket ID",ticket?.ticketId);addDeveloperField(details,"Stato AMP",ticket?.status);addDeveloperField(details,"Lock operatore",ticket?.lock?.operatorId||"Nessun lock");addDeveloperField(details,"Lock richiesto da",ticket?.lock?.requesterId);addDeveloperField(details,"Workflow lock",ticket?.lock?.workflow);addDeveloperField(details,"Lock acquisito",ticket?.lock?.acquiredAt);addDeveloperField(details,"Creato",ticket?.createdAt);addDeveloperField(details,"Ultima attività",ticket?.lastActivityAt);container.appendChild(details);}
  function renderTicketFields(parsed){const container=byId("developer-ticket-fields");if(!container)return;container.replaceChildren();container.className="developer-fields";const fields=parsed?.fields||{},itemId=parsed?.itemId||fields?.id||"",lookupId=fields?.AssegnataaLookupId,operator=findOperatorByLookupId(lookupId);addDeveloperField(container,"Item ID",itemId);addDeveloperField(container,"Gruppo",fields?.Gruppo);addDeveloperField(container,"Assegnatario",operator?`${operator[1].displayName} (${operator[0]} · LookupId ${lookupId})`:(lookupId?`LookupId ${lookupId}`:null));addDeveloperField(container,"Stato SharePoint",fields?._Status);const metadata=parseAmpMetadata(fields?.AMP_x002d_AVANADE);if(metadata.error){const e=document.createElement("div");e.className="developer-ticket-empty status-error";e.textContent=`Metadata AMP non valido: ${metadata.error}`;container.appendChild(e);return;}if(!metadata.tickets.length){const empty=document.createElement("div");empty.className="developer-ticket-empty";empty.textContent="Nessun ticket AMP presente.";container.appendChild(empty);return;}const selectorBox=document.createElement("div");selectorBox.className="developer-ticket-selector";const label=document.createElement("label");label.textContent="Ticket AMP";const selector=document.createElement("select");metadata.tickets.forEach((ticket,index)=>{const option=document.createElement("option");option.value=String(index);option.textContent=ticket?.ticketId||`Ticket ${index+1}`;selector.appendChild(option);});const host=document.createElement("div");const redraw=()=>{host.replaceChildren();renderSelectedAmpTicket(host,metadata.tickets[Number(selector.value)||0]);};selector.addEventListener("change",redraw);selectorBox.append(label,selector);container.append(selectorBox,host);redraw();}
  function applyDeveloperResultView(){const raw=byId("developer-ticket-response"),fields=byId("developer-ticket-fields"),button=byId("developer-result-view-toggle"),showJson=developerResultView==="json";if(raw)raw.hidden=!showJson;if(fields)fields.hidden=showJson;if(button)button.textContent=showJson?"Mostra campi":"Mostra JSON";}
  function showTicketResult(result,action){const panel=byId("developer-ticket-result"),status=byId("developer-ticket-status"),auth=byId("developer-ticket-auth-mode"),response=byId("developer-ticket-response");if(!panel||!status||!auth||!response)return;panel.hidden=false;setText("developer-result-action",action,"");status.textContent=`${result.status} ${result.statusText||""}`.trim();status.className=result.ok?"status-success":"status-error";auth.textContent=result.authenticationMode||"Non disponibile";let parsed=null;try{parsed=JSON.parse(result.body||"{}");response.textContent=JSON.stringify(parsed,null,2);}catch{response.textContent=result.body||"(Risposta vuota)";}if(parsed?.itemId)developerLastItemId=String(parsed.itemId);developerResultView=parsed?.fields?"fields":"json";if(parsed?.fields)renderTicketFields(parsed);else byId("developer-ticket-fields")?.replaceChildren();applyDeveloperResultView();}
  async function executeDeveloperAction(){const action=getDeveloperAction(),value=String(byId("developer-action-input")?.value||"").trim();resetTicketResult();setTicketButtonLoading(true);try{if(!value)throw new Error(action==="GetTicketById"?"Inserisci un Item ID.":"Inserisci un Conversation ID.");const authentication=await window.WorkflowHealth.getAccessTokenSilently();if(!authentication?.accessToken)throw new Error("Token di accesso non disponibile.");await ensureOperationalConfig(authentication.accessToken);const parameter=action==="GetTicketById"?"itemId":"conversationId";const url=`${DEVELOPER_ENDPOINTS[action]}?${parameter}=${encodeURIComponent(value)}`;const fetchResponse=await fetch(url,{method:"GET",headers:{Authorization:`Bearer ${authentication.accessToken}`,Accept:"application/json"}});const body=await fetchResponse.text();const result={ok:fetchResponse.ok,status:fetchResponse.status,statusText:fetchResponse.statusText,authenticationMode:authentication.acquisitionMode,body};showTicketResult(result,action);let parsed=null;try{parsed=JSON.parse(body||"{}");}catch{}currentSharePointData={conversationId:action==="GetTicketByConversationId"?value:parsed?.conversationId||null,ok:result.ok,status:result.status,statusText:result.statusText,authenticationMode:result.authenticationMode,retrievedAt:new Date().toISOString(),data:parsed,rawBody:parsed?null:body};publishSharePointData();}catch(error){showTicketResult({ok:false,status:"Errore",statusText:"",authenticationMode:"Non disponibile",body:error?.message||String(error)},action);}finally{setTicketButtonLoading(false);}}
  function registerDeveloperToolsControls(){byId("developer-action-select")?.addEventListener("change",updateDeveloperActionUi);byId("developer-execute-button")?.addEventListener("click",executeDeveloperAction);byId("developer-result-view-toggle")?.addEventListener("click",()=>{developerResultView=developerResultView==="fields"?"json":"fields";applyDeveloperResultView();});byId("close-developer-ticket-result")?.addEventListener("click",resetTicketResult);updateDeveloperActionUi();}
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
      syncDeveloperConversationId();
    } catch (error) {
      if (sequence !== loadSequence) {
        return;
      }

      currentMessageContext = null;
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
    // Nessuna delle due esegue automaticamente GetTicketByConversationId.
    void loadSelectedMessage();
    void initializeOperationalConfig();
    refreshAuthDiagnostics();
  });
})();
