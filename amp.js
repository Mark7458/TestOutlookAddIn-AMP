(() => {
  "use strict";

  const state = { initialized: false, unsubscribe: null };
  const byId = (id) => document.getElementById(id);

  function renderContext(context) {
    const output = byId("amp-message-context");
    const status = byId("amp-status");
    if (!output || !status) return;
    if (!context) {
      output.textContent = "Nessun messaggio selezionato.";
      status.textContent = "Seleziona un messaggio in Outlook.";
      return;
    }
    output.textContent = JSON.stringify(context, null, 2);
    status.textContent = `Contesto aggiornato: ${context.subject || "(nessun oggetto)"}`;
  }

  async function refresh() {
    const status = byId("amp-status");
    if (status) status.textContent = "Aggiornamento contesto in corso…";
    try {
      await window.TestPluginAPI?.refreshCurrentMessage?.();
      renderContext(window.TestPluginAPI?.getCurrentMessageContext?.() || null);
    } catch (error) {
      if (status) status.textContent = `Errore: ${error?.message || String(error)}`;
    }
  }

  function show() {
    document.getElementById("main-view")?.setAttribute("hidden", "");
    document.getElementById("header-view")?.setAttribute("hidden", "");
    const view = byId("amp-view");
    if (view) view.hidden = false;
    renderContext(window.TestPluginAPI?.getCurrentMessageContext?.() || null);
  }

  async function initialize() {
    if (state.initialized) return;
    const view = byId("amp-view");
    if (!view) return;
    const response = await fetch("amp-view.html", { cache: "no-store" });
    if (!response.ok) throw new Error(`Impossibile caricare amp-view.html (${response.status}).`);
    view.innerHTML = await response.text();
    byId("amp-back-button")?.addEventListener("click", () => window.TestPluginAPI?.showMainView?.());
    byId("amp-refresh-button")?.addEventListener("click", refresh);
    byId("open-amp-button")?.addEventListener("click", show);
    state.unsubscribe = window.TestPluginAPI?.onMessageChanged?.(renderContext) || null;
    state.initialized = true;
  }

  window.AmpModule = { initialize, show, refresh };
  document.addEventListener("DOMContentLoaded", () => initialize().catch((error) => console.error("AMP init error:", error)));
})();
