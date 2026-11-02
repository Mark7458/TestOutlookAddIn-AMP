// Test-Plugin file version: 0.2-auth-src.js-2
import {
  PublicClientApplication
} from "@azure/msal-browser";

const AUTH_CONFIG = {
  clientId:
    "17be48ae-1d05-4233-ba98-654102bf76e0",
  tenantId:
    "9284cd5c-ae26-4eb2-b079-a14b6df5f1aa",
  scope:
    "api://17be48ae-1d05-4233-ba98-654102bf76e0/Workflow.Access",
  redirectUri:
    "https://jakdrizzt.github.io/TestOutlookAddIn/auth-dialog.html",
  healthUrl:
    "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/health",
  dialogUrl:
    "https://jakdrizzt.github.io/TestOutlookAddIn/auth-dialog.html"
};

let msalInstance = null;
let dialogOperationSequence = 0;
let currentAuthentication = null;
let persistedBridgeRejectedThisRun = false;
const AUTH_DIAGNOSTIC_KEY = "outlook-sogei-taskpane-auth-diagnostics";
const MAX_AUTH_DIAGNOSTIC_EVENTS = 150;
const STORAGE_PROBE_TASKPANE_KEY = "outlook-sogei-storage-probe-taskpane";
const STORAGE_PROBE_DIALOG_KEY = "outlook-sogei-storage-probe-dialog";

const AUTH_FLOW = "popup-redirect-bridge";
const AUTH_BRIDGE_KEY = "outlook-sogei-taskpane-auth-bridge";
const UI_PREFERENCES_KEY = "test-plugin-ui-preferences";
function isDiagnosticsEnabled() { try { return Boolean(JSON.parse(localStorage.getItem(UI_PREFERENCES_KEY) || "{}").diagnosticsEnabled); } catch { return false; } }

function readAuthDiagnostics() {
  try {
    const parsed = JSON.parse(
      localStorage.getItem(AUTH_DIAGNOSTIC_KEY) || "[]"
    );
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function writeAuthDiagnostics(events) {
  localStorage.setItem(
    AUTH_DIAGNOSTIC_KEY,
    JSON.stringify(events.slice(-MAX_AUTH_DIAGNOSTIC_EVENTS))
  );
}

function recordAuthDiagnostic(event, details = null, source = "TASKPANE") {
  if (!isDiagnosticsEnabled()) return;
  const events = readAuthDiagnostics();
  events.push({
    time: new Date().toLocaleTimeString("it-IT", { hour12: false }),
    source,
    event,
    details
  });
  writeAuthDiagnostics(events);
  window.dispatchEvent(new CustomEvent("workflow-auth-diagnostics-changed"));
}

function runTaskpaneStorageProbe() {
  const probe = {
    source: "TASKPANE",
    timestamp: new Date().toISOString(),
    origin: window.location.origin
  };
  let writeSucceeded = false;
  let ownProbeVisible = false;
  let dialogProbeVisible = false;
  let dialogProbeSource = null;
  let dialogProbeTimestamp = null;

  try {
    localStorage.setItem(STORAGE_PROBE_TASKPANE_KEY, JSON.stringify(probe));
    writeSucceeded = true;
    ownProbeVisible = Boolean(localStorage.getItem(STORAGE_PROBE_TASKPANE_KEY));

    const dialogProbeRaw = localStorage.getItem(STORAGE_PROBE_DIALOG_KEY);
    dialogProbeVisible = Boolean(dialogProbeRaw);
    if (dialogProbeRaw) {
      try {
        const dialogProbe = JSON.parse(dialogProbeRaw);
        dialogProbeSource = dialogProbe?.source || null;
        dialogProbeTimestamp = dialogProbe?.timestamp || null;
      } catch {
        dialogProbeSource = "invalid-json";
      }
    }
  } catch (error) {
    recordAuthDiagnostic("storage_probe_taskpane_error", {
      name: error?.name || "Error",
      message: error?.message || String(error)
    });
    return;
  }

  recordAuthDiagnostic("storage_probe_taskpane", {
    writeSucceeded,
    ownProbeVisible,
    dialogProbeVisible,
    dialogProbeSource,
    dialogProbeTimestamp,
    origin: window.location.origin
  });
}

function savePersistedAuthentication(authentication, source = "unknown") {
  if (!authentication?.accessToken) return;
  const payload = { accessToken: authentication.accessToken, expiresOn: authentication.expiresOn || null, savedAt: new Date().toISOString() };
  try {
    localStorage.setItem(AUTH_BRIDGE_KEY, JSON.stringify(payload));
    recordAuthDiagnostic("persisted_auth_saved", { source, expiresOn: payload.expiresOn, savedAt: payload.savedAt, tokenLength: payload.accessToken.length });
  } catch (error) {
    recordAuthDiagnostic("persisted_auth_save_error", { source, name: error?.name || "Error", message: error?.message || String(error) });
  }
}
function readPersistedAuthentication() {
  try {
    const raw = localStorage.getItem(AUTH_BRIDGE_KEY);
    if (!raw) { recordAuthDiagnostic("persisted_auth_missing"); return null; }
    const parsed = JSON.parse(raw);
    const candidate = { accessToken: parsed?.accessToken || null, acquisitionMode: "persisted-bridge", expiresOn: parsed?.expiresOn || null };
    recordAuthDiagnostic("persisted_auth_found", { savedAt: parsed?.savedAt || null, expiresOn: candidate.expiresOn, tokenPresent: Boolean(candidate.accessToken), tokenLength: candidate.accessToken?.length || 0 });
    return candidate.accessToken ? candidate : null;
  } catch (error) {
    recordAuthDiagnostic("persisted_auth_read_error", { name: error?.name || "Error", message: error?.message || String(error) });
    return null;
  }
}
function clearPersistedAuthentication(reason = "unknown") {
  try { localStorage.removeItem(AUTH_BRIDGE_KEY); recordAuthDiagnostic("persisted_auth_cleared", { reason }); }
  catch (error) { recordAuthDiagnostic("persisted_auth_clear_error", { reason, name: error?.name || "Error", message: error?.message || String(error) }); }
}
function clearAuthenticationCache() {
  currentAuthentication = null;
  persistedBridgeRejectedThisRun = false;
  clearPersistedAuthentication("manualDiagnosticAction");
  recordAuthDiagnostic("manual_auth_cache_clear");
}
function resetDialogAuthenticationCache() {
  recordAuthDiagnostic("full_auth_reset_dialog_start");
  return new Promise((resolve, reject) => {
    if (!window.Office?.context?.ui || typeof Office.context.ui.displayDialogAsync !== "function") {
      reject(createDialogError("Office Dialog API non disponibile.", { errorCode: "office_dialog_not_available" }));
      return;
    }
    Office.context.ui.displayDialogAsync(
      `${AUTH_CONFIG.dialogUrl}?mode=reset`,
      { height: 10, width: 10, displayInIframe: false, promptBeforeOpen: false },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(createDialogError(result.error?.message || "Impossibile aprire il dialog reset.", { errorCode: result.error?.code || "dialog_reset_open_failed" }));
          return;
        }
        const dialog = result.value;
        let completed = false;
        const finish = (callback) => {
          if (completed) return;
          completed = true;
          try { dialog.close(); } catch {}
          callback();
        };
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (event) => {
          try {
            const message = JSON.parse(event.message);
            importDialogDiagnostics(message?.dialogDiagnostics);
            if (message?.type === "auth-reset-complete") {
              recordAuthDiagnostic("full_auth_reset_dialog_complete", { removedCount: message.removedCount || 0 });
              finish(() => resolve(message));
              return;
            }
            finish(() => reject(createDialogError(message?.message || "Reset dialog non completato.", { errorCode: message?.errorCode || "dialog_reset_failed" })));
          } catch (error) { finish(() => reject(error)); }
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (event) => {
          if (!completed) recordAuthDiagnostic("full_auth_reset_dialog_event", { error: event?.error || null });
        });
      }
    );
  });
}

async function resetAuthenticationCompletely() {
  clearAuthenticationCache();
  recordAuthDiagnostic("full_auth_reset_start");
  const result = await resetDialogAuthenticationCache();
  recordAuthDiagnostic("full_auth_reset_complete", { removedCount: result?.removedCount || 0 });
  return result;
}

function importDialogDiagnostics(entries) {
  if (!Array.isArray(entries)) return;

  entries.slice(-12).forEach((entry) => {
    const details = {
      operationId: entry?.details?.operationId ?? null,
      dialogMode: entry?.details?.dialogMode || null,
      ...(entry?.details || {})
    };
    recordAuthDiagnostic(
      entry.event || "dialog_event",
      details,
      "DIALOG"
    );
  });
}

function clearAuthDiagnostics() {
  localStorage.removeItem(AUTH_DIAGNOSTIC_KEY);
  window.dispatchEvent(new CustomEvent("workflow-auth-diagnostics-changed"));
}

function createDialogError(message, details = {}) {
  const error = new Error(message);
  error.name =
    details.name ||
    "DialogAuthenticationError";
  error.errorCode =
    details.errorCode ||
    null;
  error.subError =
    details.subError ||
    null;
  error.correlationId =
    details.correlationId ||
    null;
  error.authenticationPhase =
    details.authenticationPhase ||
    "officeDialog";
  error.dialogEvent =
    details.dialogEvent ||
    null;
  return error;
}

async function initializeMsal() {
  if (msalInstance) {
    return msalInstance;
  }

  runTaskpaneStorageProbe();
  recordAuthDiagnostic("auth_experiment_loaded", {
    flow: AUTH_FLOW
  });
  recordAuthDiagnostic("msal_initialize_start");

  msalInstance =
    new PublicClientApplication({
      auth: {
        clientId:
          AUTH_CONFIG.clientId,
        authority:
          `https://login.microsoftonline.com/${AUTH_CONFIG.tenantId}`,
        redirectUri:
          AUTH_CONFIG.redirectUri,
        navigateToLoginRequestUrl:
          false
      },
      cache: {
        cacheLocation:
          "localStorage",
        storeAuthStateInCookie:
          true
      }
    });

  await msalInstance.initialize();
  recordAuthDiagnostic("msal_initialize_complete", {
    accounts: msalInstance.getAllAccounts().length
  });
  return msalInstance;
}

/*
 * Cerca PRIMA un account e un token nella cache MSAL.
 * Restituisce null soltanto quando serve il flusso interattivo.
 */
function isAuthenticationUsable(authentication) {
  if (!authentication?.accessToken) {
    return false;
  }

  if (!authentication.expiresOn) {
    return true;
  }

  const expiresOn = new Date(authentication.expiresOn).getTime();
  const safetyWindowMs = 2 * 60 * 1000;

  return Number.isFinite(expiresOn) &&
    expiresOn - Date.now() > safetyWindowMs;
}

async function acquireTokenSilently() {
  const client =
    await initializeMsal();

  const accounts =
    client.getAllAccounts();

  recordAuthDiagnostic("silent_accounts", {
    count: accounts.length,
    usernames: accounts.map((account) => account.username).filter(Boolean).join(", ")
  });

  console.info(
    "Account MSAL disponibili nella cache:",
    accounts.length
  );

  if (accounts.length === 0) {
    recordAuthDiagnostic("silent_accounts_empty", {
      flow: AUTH_FLOW,
      nextStep: "interactiveConnectPlugin"
    });
    return null;
  }


  try {
    recordAuthDiagnostic("silent_start");

    const result =
      await client.acquireTokenSilent({
        account:
          accounts[0],
        scopes: [
          AUTH_CONFIG.scope
        ]
      });

    if (!result?.accessToken) {
      return null;
    }

    console.info(
      "Token MSAL ottenuto silenziosamente per:",
      result.account?.username ||
        accounts[0].username ||
        "account non disponibile"
    );

    recordAuthDiagnostic("silent_success", {
      username: result.account?.username || accounts[0].username || null,
      expiresOn: result.expiresOn?.toISOString?.() || null
    });

    currentAuthentication = {
      accessToken:
        result.accessToken,
      acquisitionMode:
        "silent",
      expiresOn:
        result.expiresOn || null
    };

    return currentAuthentication;
  } catch (error) {
    recordAuthDiagnostic("silent_error", {
      name: error?.name || null,
      errorCode: error?.errorCode || null,
      message: error?.message || String(error)
    });

    console.warn(
      "Token silent non disponibile; apro il dialogo interattivo.",
      {
        name:
          error?.name,
        errorCode:
          error?.errorCode,
        message:
          error?.message
      }
    );

    return null;
  }
}

function acquireTokenWithOfficeDialog(options = {}) {
  const operationId = ++dialogOperationSequence;
  const mode = options.mode === "recovery" ? "recovery" : "connect";
  const forceCloseAfterMs = Number(options.forceCloseAfterMs || 0);
  const dialogUrlBase = mode === "recovery" ? `${AUTH_CONFIG.dialogUrl}?mode=recovery` : AUTH_CONFIG.dialogUrl;
  const dialogUrl = `${dialogUrlBase}${dialogUrlBase.includes("?") ? "&" : "?"}operationId=${encodeURIComponent(operationId)}`;
  const dialogHeight = Number(options.height || (mode === "recovery" ? 10 : 65));
  const dialogWidth = Number(options.width || (mode === "recovery" ? 10 : 45));
  recordAuthDiagnostic("display_dialog_start", { operationId, mode, forceCloseAfterMs, dialogHeight, dialogWidth });

  return new Promise((resolve, reject) => {
    if (
      !window.Office ||
      !Office.context ||
      !Office.context.ui ||
      typeof Office.context.ui.displayDialogAsync !==
        "function"
    ) {
      reject(
        createDialogError(
          "Office Dialog API non è disponibile in questo client Outlook.",
          {
            errorCode:
              "office_dialog_not_available"
          }
        )
      );
      return;
    }

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      {
        height: dialogHeight,
        width: dialogWidth,
        displayInIframe: false,
        promptBeforeOpen: false
      },
      (result) => {
        if (
          result.status ===
          Office.AsyncResultStatus.Failed
        ) {
          recordAuthDiagnostic("display_dialog_error", {
            operationId,
            code: result.error?.code || null,
            message: result.error?.message || null
          });
          reject(
            createDialogError(
              result.error?.message ||
                "Impossibile aprire la finestra di autenticazione.",
              {
                errorCode:
                  result.error?.code ||
                  "dialog_open_failed",
                authenticationPhase:
                  "displayDialogAsync"
              }
            )
          );
          return;
        }

        const dialog = result.value;
        let completed = false;
        const dialogOpenedAt = Date.now();
        let dialog12006Count = 0;

        recordAuthDiagnostic("display_dialog_success", { operationId });

        let forcedCloseTimer = null;
        if (forceCloseAfterMs > 0) {
          recordAuthDiagnostic("bootstrap_wait_started", {
            operationId,
            forceCloseAfterMs
          });
          forcedCloseTimer = setTimeout(() => {
            if (completed || operationId !== dialogOperationSequence) return;
            recordAuthDiagnostic("bootstrap_parent_close", {
              operationId,
              elapsedMs: Date.now() - dialogOpenedAt
            });
            completeWithError(
              createDialogError(
                "Finestra bootstrap conclusa; avvio recovery.",
                {
                  errorCode: "bootstrap_window_elapsed",
                  authenticationPhase: "porkaroundBootstrap"
                }
              )
            );
          }, forceCloseAfterMs);
        }

        const authenticationTimeout =
          setTimeout(() => {
            recordAuthDiagnostic("dialog_timeout", { operationId });
            completeWithError(
              createDialogError(
                "Tempo massimo per l'autenticazione superato.",
                {
                  errorCode:
                    "authentication_timeout",
                  authenticationPhase:
                    "officeDialog"
                }
              )
            );
          }, 5 * 60 * 1000);

        function clearAuthenticationTimeout() {
          clearTimeout(authenticationTimeout);
          if (forcedCloseTimer) clearTimeout(forcedCloseTimer);
        }

        function closeDialog() {
          try {
            dialog.close();
          } catch {
            // Il dialogo potrebbe essere già chiuso.
          }
        }

        function completeWithError(error) {
          if (completed || operationId !== dialogOperationSequence) {
            return;
          }

          completed = true;
          dialogOperationSequence += 1;
          clearAuthenticationTimeout();
          closeDialog();
          recordAuthDiagnostic("dialog_complete_error", {
            operationId,
            name: error?.name || null,
            errorCode: error?.errorCode || null,
            message: error?.message || String(error)
          });
          reject(error);
        }

        function completeWithToken(tokenResult) {
          if (completed || operationId !== dialogOperationSequence) {
            return;
          }

          completed = true;
          dialogOperationSequence += 1;
          clearAuthenticationTimeout();
          closeDialog();
          currentAuthentication = tokenResult;
          savePersistedAuthentication(tokenResult, "dialogCompleteSuccess");
          recordAuthDiagnostic("dialog_complete_success", {
            operationId,
            acquisitionMode: tokenResult?.acquisitionMode || null,
            expiresOn: tokenResult?.expiresOn || null
          });
          resolve(tokenResult);
        }

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (event) => {
            let message;

            try {
              message =
                JSON.parse(event.message);
              recordAuthDiagnostic("dialog_message_received", {
                operationId,
                dialogOperationId: message?.dialogOperationId ?? null,
                dialogMode: message?.dialogMode || null,
                type: message?.type || null
              });
              if (message?.dialogOperationId != null && String(message.dialogOperationId) !== String(operationId)) {
                recordAuthDiagnostic("dialog_operation_mismatch", { operationId, dialogOperationId: message.dialogOperationId, dialogMode: message?.dialogMode || null });
              }
              importDialogDiagnostics(message?.dialogDiagnostics);
            } catch {
              completeWithError(
                createDialogError(
                  "Il dialogo ha restituito una risposta non valida.",
                  {
                    errorCode:
                      "invalid_dialog_message",
                    authenticationPhase:
                      "officeDialog"
                  }
                )
              );
              return;
            }

            if (
              message.type ===
                "auth-success" &&
              message.accessToken
            ) {
              recordAuthDiagnostic("dialog_auth_success_identity", {
                username: message.authenticationResult?.account?.username || null,
                homeAccountId: message.authenticationResult?.account?.homeAccountId || null,
                tenantId: message.authenticationResult?.tenantId || message.authenticationResult?.account?.tenantId || null
              });
              const finishAuthentication =
                async () => {
                  if (
                    message.authenticationResult
                      ?.account
                  ) {
                    try {
                      const client =
                        await initializeMsal();

                      const hydratedResult = {
                        ...message.authenticationResult,
                        expiresOn:
                          message.authenticationResult.expiresOn
                            ? new Date(
                                message.authenticationResult.expiresOn
                              )
                            : null,
                        extExpiresOn:
                          message.authenticationResult.extExpiresOn
                            ? new Date(
                                message.authenticationResult.extExpiresOn
                              )
                            : null
                      };

                      console.info("Cache MSAL task pane prima di hydrateCache:", client.getAllAccounts().length);
                      await client.hydrateCache(
                        hydratedResult,
                        { scopes: [AUTH_CONFIG.scope], account: hydratedResult.account }
                      );
                      const accountsAfterHydrate = client.getAllAccounts();
                      console.info("Cache MSAL task pane dopo hydrateCache:", accountsAfterHydrate.length,
                        accountsAfterHydrate.map((account) => account.username));
                      if (accountsAfterHydrate.length === 0) {
                        throw new Error("hydrateCache completato ma nessun account è presente nella cache MSAL del task pane.");
                      }

                      client.setActiveAccount(
                        hydratedResult.account
                      );

                      console.info(
                        "Cache MSAL del task pane sincronizzata dal dialogo per:",
                        hydratedResult.account
                          ?.username ||
                          "account non disponibile"
                      );
                    } catch (cacheError) {
                      console.warn(
                        "Impossibile sincronizzare la cache MSAL del task pane.",
                        {
                          name:
                            cacheError?.name,
                          errorCode:
                            cacheError?.errorCode,
                          message:
                            cacheError?.message
                        }
                      );
                    }
                  }

                  completeWithToken({
                    accessToken:
                      message.accessToken,
                    acquisitionMode:
                      message.acquisitionMode ||
                      "interactive",
                    expiresOn:
                      message.authenticationResult?.expiresOn ||
                      null
                  });
                };

              finishAuthentication().catch(
                completeWithError
              );
              return;
            }

            completeWithError(
              createDialogError(
                message.message ||
                  "Autenticazione non completata.",
                {
                  name:
                    message.name,
                  errorCode:
                    message.errorCode,
                  subError:
                    message.subError,
                  correlationId:
                    message.correlationId,
                  authenticationPhase:
                    message.authenticationPhase ||
                    "authDialog"
                }
              )
            );
          }
        );

        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          (event) => {
            recordAuthDiagnostic("dialog_event_received", {
              operationId,
              error: event?.error || null
            });
            console.warn(
              "Evento Office Dialog ricevuto:",
              {
                error:
                  event?.error
              }
            );

            if (event?.error === 12006) {
              dialog12006Count += 1;
              recordAuthDiagnostic("popup_dialog_closed_by_user", { operationId, count: dialog12006Count, elapsedMs: Date.now() - dialogOpenedAt });
              completeWithError(createDialogError("Finestra di autenticazione chiusa dall'utente.", { errorCode: "user_cancelled", authenticationPhase: "popupOfficeDialog", dialogEvent: 12006 }));
            }
          }
        );
      }
    );
  });
}

async function recoverConnection() {
  recordAuthDiagnostic("recover_connection_start", { mode: "persistent-small-dialog" });
  try {
    const authentication = await acquireTokenWithOfficeDialog({
      mode: "recovery",
      forceCloseAfterMs: 22000,
      height: 10,
      width: 10
    });
    recordAuthDiagnostic("recover_connection_complete", {
      acquisitionMode: authentication?.acquisitionMode || null
    });
    return authentication;
  } catch (error) {
    recordAuthDiagnostic("recover_connection_failed", {
      errorCode: error?.errorCode || null,
      message: error?.message || String(error)
    });
    throw createDialogError(
      "La sessione non è recuperabile. Riconnetti il plugin.",
      {
        errorCode: "interaction_required",
        authenticationPhase: "silentRecovery"
      }
    );
  }
}
async function probeCurrentAuthentication() {
  recordAuthDiagnostic("startup_auth_probe_start");

  if (isAuthenticationUsable(currentAuthentication)) {
    recordAuthDiagnostic("startup_auth_probe_memory_success", {
      expiresOn: currentAuthentication?.expiresOn || null
    });
    return { ...currentAuthentication, acquisitionMode: "memory" };
  }

  currentAuthentication = null;
  const authentication = await acquireTokenSilently();

  recordAuthDiagnostic(
    authentication ? "startup_auth_probe_success" : "startup_auth_probe_empty",
    { acquisitionMode: authentication?.acquisitionMode || null }
  );

  return authentication;
}

async function getAccessTokenSilently() {
  if (isAuthenticationUsable(currentAuthentication)) {
    recordAuthDiagnostic("token_memory_reused", {
      expiresOn: currentAuthentication?.expiresOn || null
    });
    return { ...currentAuthentication, acquisitionMode: "memory" };
  }
  const persistedAuthentication = readPersistedAuthentication();
  if (isAuthenticationUsable(persistedAuthentication)) {
    persistedBridgeRejectedThisRun = false;
    currentAuthentication = persistedAuthentication;
    recordAuthDiagnostic("persisted_auth_reused", { expiresOn: persistedAuthentication.expiresOn || null });
    return currentAuthentication;
  }
  if (persistedAuthentication) {
    persistedBridgeRejectedThisRun = true;
    recordAuthDiagnostic("persisted_auth_rejected_this_run", {
      reason: "notUsable",
      expiresOn: persistedAuthentication.expiresOn || null
    });
    clearPersistedAuthentication("notUsable");
  }

  currentAuthentication = null;
  const authentication = await acquireTokenSilently();
  if (authentication) {
    persistedBridgeRejectedThisRun = false;
    return authentication;
  }
  const silentAccountCount = msalInstance?.getAllAccounts?.().length || 0;
  if (!persistedAuthentication && silentAccountCount === 0) {
    recordAuthDiagnostic("zero_auth_state_detected", {
      persistedAuthAvailable: false,
      accountCount: silentAccountCount,
      reason: "emptyStateAfterReset"
    });
    recordAuthDiagnostic("initial_recovery_skipped", {
      flow: AUTH_FLOW,
      reason: "nothingToRecover"
    });
    recordAuthDiagnostic("interactive_auth_selected", {
      flow: AUTH_FLOW,
      nextStep: "connectPlugin"
    });
    return connectPlugin();
  }
  if (persistedBridgeRejectedThisRun) {
    persistedBridgeRejectedThisRun = false;
    recordAuthDiagnostic("expired_bridge_preflight_start", { flow: AUTH_FLOW });
    try {
      const recovered = await acquireTokenWithOfficeDialog({ mode: "recovery", forceCloseAfterMs: 3000, height: 10, width: 10 });
      recordAuthDiagnostic("expired_bridge_preflight_success", { acquisitionMode: recovered?.acquisitionMode || null });
      return recovered;
    } catch (error) {
      recordAuthDiagnostic("expired_bridge_preflight_empty", { errorCode: error?.errorCode || null });
      return connectPlugin();
    }
  }
  recordAuthDiagnostic("taskpane_silent_empty_recover_via_dialog");
  return recoverConnection();
}
async function connectPlugin() {
  recordAuthDiagnostic("connect_plugin_start");
  currentAuthentication = null;
  const cached = await acquireTokenSilently();
  if (cached) {
    recordAuthDiagnostic("connect_plugin_silent_success");
    return cached;
  }

  const interactiveWindowMs = 0;
  recordAuthDiagnostic("single_interactive_start", { interactiveWindowMs, mode: "popupInsideDialog" });
  try {
    const directResult = await acquireTokenWithOfficeDialog({ mode: "connect", forceCloseAfterMs: interactiveWindowMs });
    recordAuthDiagnostic("single_interactive_direct_success", { acquisitionMode: directResult?.acquisitionMode || null });
    return directResult;
  } catch (error) {
    const errorCode = error?.errorCode || null;
    recordAuthDiagnostic("single_interactive_released", { errorCode, message: error?.message || String(error) });
    if (["user_cancelled", "timed_out"].includes(String(errorCode || ""))) {
      recordAuthDiagnostic("popup_terminal_error_stop", { errorCode, reason: "noRecoveryForPopupCancellation" });
      throw error;
    }
  }

  recordAuthDiagnostic("persistent_recovery_start", { maxMs: 20000, height: 10, width: 10 });
  const authentication = await acquireTokenWithOfficeDialog({
    mode: "recovery",
    forceCloseAfterMs: 22000,
    height: 10,
    width: 10
  });
  recordAuthDiagnostic("persistent_recovery_success", { acquisitionMode: authentication?.acquisitionMode || null });
  return authentication;
}
async function callHealthOnDemand() {
  recordAuthDiagnostic("call_health_start");

  let authentication;
  if (isAuthenticationUsable(currentAuthentication)) {
    recordAuthDiagnostic("token_memory_reused", {
      expiresOn: currentAuthentication?.expiresOn || null
    });
    authentication = { ...currentAuthentication, acquisitionMode: "memory" };
  } else {
    authentication = await connectPlugin();
  }
  const healthStartedAt = Date.now();
  recordAuthDiagnostic("health_fetch_start", { authenticationMode: authentication.acquisitionMode });
  const response = await fetch(AUTH_CONFIG.healthUrl, { method: "GET", headers: { Authorization: `Bearer ${authentication.accessToken}`, Accept: "application/json" } });
  const body = await response.text();
  recordAuthDiagnostic("function_response", { status: response.status, ok: response.ok, authenticationMode: authentication.acquisitionMode, elapsedMs: Date.now() - healthStartedAt });
  if (response.status === 401 || response.status === 403) {
    throw createDialogError(`La Function ha rifiutato l'autenticazione (${response.status}).`, { errorCode: "interaction_required", authenticationPhase: "healthResponse" });
  }
  return { ok: response.ok, status: response.status, statusText: response.statusText, authenticationMode: authentication.acquisitionMode, body };
}

async function callHealth() {
  recordAuthDiagnostic("call_health_start");
  const authentication = await getAccessTokenSilently();
  const healthStartedAt = Date.now();
  recordAuthDiagnostic("health_fetch_start", { authenticationMode: authentication.acquisitionMode });
  const response = await fetch(AUTH_CONFIG.healthUrl, { method: "GET", headers: { Authorization: `Bearer ${authentication.accessToken}`, Accept: "application/json" } });
  const body = await response.text();
  recordAuthDiagnostic("function_response", { status: response.status, ok: response.ok, authenticationMode: authentication.acquisitionMode, elapsedMs: Date.now() - healthStartedAt });
  if (response.status === 401 || response.status === 403) {
    throw createDialogError(
      `La Function ha rifiutato l'autenticazione (${response.status}).`,
      {
        errorCode: "interaction_required",
        authenticationPhase: "healthResponse"
      }
    );
  }
  return { ok: response.ok, status: response.status, statusText: response.statusText, authenticationMode: authentication.acquisitionMode, body };
}
window.WorkflowHealth = { connectPlugin, recoverConnection, probeCurrentAuthentication, getAccessTokenSilently, callHealth, callHealthOnDemand, getDiagnostics: readAuthDiagnostics, clearDiagnostics: clearAuthDiagnostics, recordDiagnostic: (event, details = null) => recordAuthDiagnostic(event, details, "TASKPANE"), resetAuthenticationCompletely };
