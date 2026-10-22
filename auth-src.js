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
const AUTH_DIAGNOSTIC_KEY = "outlook-sogei-taskpane-auth-diagnostics";
const MAX_AUTH_DIAGNOSTIC_EVENTS = 500;

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

function importDialogDiagnostics(entries) {
  if (!Array.isArray(entries)) return;

  entries.slice(-60).forEach((entry) => {
    recordAuthDiagnostic(
      entry.event || "dialog_event",
      entry.details || null,
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
  const dialogUrl = mode === "recovery" ? `${AUTH_CONFIG.dialogUrl}?mode=recovery` : AUTH_CONFIG.dialogUrl;
  recordAuthDiagnostic("display_dialog_start", { operationId, mode, forceCloseAfterMs });

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
        height: 65,
        width: 45,
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
                type: message?.type || null
              });
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
              recordAuthDiagnostic("porkaround_12006_observed", {
                operationId,
                count: dialog12006Count,
                elapsedMs: Date.now() - dialogOpenedAt
              });
            }
          }
        );
      }
    );
  });
}

function delay(milliseconds) { return new Promise((resolve) => setTimeout(resolve, milliseconds)); }
async function runRecoveryLoop(options = {}) {
  const retryDelayMs = Number(options.retryDelayMs || 1000);
  const openedRecoveryMaxMs = Number(options.openedRecoveryMaxMs || 20000);
  const waitForDialogRelease = options.waitForDialogRelease !== false;
  let recoveryWindowStartedAt = null;
  let attempt = 0;
  recordAuthDiagnostic("recovery_watcher_start", { retryDelayMs, openedRecoveryMaxMs, waitForDialogRelease });
  while (true) {
    attempt += 1;
    recordAuthDiagnostic("recovery_attempt", { attempt });
    try {
      const authentication = await acquireTokenWithOfficeDialog({ mode: "recovery", forceCloseAfterMs: 3000 });
      recordAuthDiagnostic("recovery_success", { attempt, acquisitionMode: authentication?.acquisitionMode || null });
      return authentication;
    } catch (error) {
      const dialogBusy = error?.authenticationPhase === "displayDialogAsync";
      recordAuthDiagnostic("recovery_attempt_failed", { attempt, dialogBusy, errorCode: error?.errorCode || null, message: error?.message || String(error) });
      if (dialogBusy && waitForDialogRelease) { await delay(retryDelayMs); continue; }
      if (recoveryWindowStartedAt === null) { recoveryWindowStartedAt = Date.now(); recordAuthDiagnostic("recovery_window_started"); }
      if (Date.now() - recoveryWindowStartedAt >= openedRecoveryMaxMs) {
        throw createDialogError("Token non recuperabile entro la finestra di recovery di 20 secondi.", { errorCode: "recovery_timeout", authenticationPhase: "silentRecovery" });
      }
      await delay(retryDelayMs);
    }
  }
}
async function getAccessTokenSilently() {
  if (isAuthenticationUsable(currentAuthentication)) return { ...currentAuthentication, acquisitionMode: "memory" };
  currentAuthentication = null;
  const authentication = await acquireTokenSilently();
  if (authentication) return authentication;
  throw createDialogError("La sessione non è disponibile. Riconnetti il plugin.", { errorCode: "interaction_required", authenticationPhase: "silentOnly" });
}
async function recoverConnection() {
  recordAuthDiagnostic("recover_connection_start");
  const authentication = await runRecoveryLoop({ waitForDialogRelease: false, retryDelayMs: 1000, openedRecoveryMaxMs: 20000 });
  recordAuthDiagnostic("recover_connection_complete", { acquisitionMode: authentication?.acquisitionMode || null });
  return authentication;
}
async function connectPlugin() {
  recordAuthDiagnostic("connect_plugin_start");
  currentAuthentication = null;
  const cached = await acquireTokenSilently();
  if (cached) {
    recordAuthDiagnostic("connect_plugin_silent_success");
    return cached;
  }
  recordAuthDiagnostic("interactive_bootstrap_launch");
  try {
    const direct = await acquireTokenWithOfficeDialog({ mode: "connect" });
    recordAuthDiagnostic("interactive_bootstrap_direct_success", { acquisitionMode: direct?.acquisitionMode || null });
    return direct;
  } catch (error) {
    recordAuthDiagnostic("interactive_bootstrap_ended", { errorCode: error?.errorCode || null, message: error?.message || String(error) });
  }
  recordAuthDiagnostic("post_bootstrap_recovery_start");
  const authentication = await runRecoveryLoop({ waitForDialogRelease: false, retryDelayMs: 1000, openedRecoveryMaxMs: 20000 });
  recordAuthDiagnostic("connect_plugin_complete", { acquisitionMode: authentication?.acquisitionMode || null });
  return authentication;
}
async function callHealth() {
  recordAuthDiagnostic("call_health_start");
  const authentication = await getAccessTokenSilently();
  const healthStartedAt = Date.now();
  recordAuthDiagnostic("health_fetch_start", { authenticationMode: authentication.acquisitionMode });
  const response = await fetch(AUTH_CONFIG.healthUrl, { method: "GET", headers: { Authorization: `Bearer ${authentication.accessToken}`, Accept: "application/json" } });
  const body = await response.text();
  recordAuthDiagnostic("function_response", { status: response.status, ok: response.ok, authenticationMode: authentication.acquisitionMode, elapsedMs: Date.now() - healthStartedAt });
  return { ok: response.ok, status: response.status, statusText: response.statusText, authenticationMode: authentication.acquisitionMode, body };
}
window.WorkflowHealth = { connectPlugin, recoverConnection, getAccessTokenSilently, callHealth, getDiagnostics: readAuthDiagnostics, clearDiagnostics: clearAuthDiagnostics };
