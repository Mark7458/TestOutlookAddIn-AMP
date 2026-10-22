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
const MAX_AUTH_DIAGNOSTIC_EVENTS = 100;

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

function acquireTokenWithOfficeDialog() {
  const operationId = ++dialogOperationSequence;
  recordAuthDiagnostic("display_dialog_start", { operationId });

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
      AUTH_CONFIG.dialogUrl,
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
              const elapsedMs = Date.now() - dialogOpenedAt;
              recordAuthDiagnostic("porkaround_12006_observed", {
                operationId,
                count: dialog12006Count,
                elapsedMs
              });

              if (dialog12006Count >= 2 || elapsedMs >= 12000) {
                completeWithError(
                  createDialogError(
                    "Dialogo interattivo chiuso dopo il login; avvio del dialogo di recupero.",
                    {
                      errorCode: "dialog_closed_after_auth",
                      authenticationPhase: "porkaround",
                      dialogEvent: 12006
                    }
                  )
                );
              }
            }
          }
        );
      }
    );
  });
}

async function getAuthentication() {
  if (isAuthenticationUsable(currentAuthentication)) {
    recordAuthDiagnostic("memory_token_reused", {
      expiresOn: currentAuthentication.expiresOn || null
    });
    return {
      ...currentAuthentication,
      acquisitionMode: "memory"
    };
  }

  currentAuthentication = null;

  const silentAuthentication =
    await acquireTokenSilently();

  if (silentAuthentication) {
    return silentAuthentication;
  }

  /*
   * PORKAROUND:
   * il primo dialogo puo completare login/MFA ma perdere il bridge Office
   * durante il redirect. Se il primo dialogo termina con errore, proviamo
   * subito la cache MSAL. Se non e ancora utilizzabile, apriamo UN SOLO
   * secondo dialogo fresco. Nessun loop ulteriore e consentito.
   */
  try {
    return await acquireTokenWithOfficeDialog();
  } catch (firstDialogError) {
    recordAuthDiagnostic("porkaround_first_dialog_failed", {
      name: firstDialogError?.name || null,
      errorCode: firstDialogError?.errorCode || null,
      message: firstDialogError?.message || String(firstDialogError)
    });

    if (firstDialogError?.errorCode !== "dialog_closed_after_auth") {
      recordAuthDiagnostic("porkaround_silent_recovery_start");
      const recoveredAuthentication =
        await acquireTokenSilently();

      if (recoveredAuthentication) {
        recordAuthDiagnostic("porkaround_silent_recovery_success", {
          acquisitionMode: recoveredAuthentication.acquisitionMode || null,
          expiresOn: recoveredAuthentication.expiresOn || null
        });
        return recoveredAuthentication;
      }
    } else {
      recordAuthDiagnostic("porkaround_first_dialog_closed_after_token");
    }

    recordAuthDiagnostic("porkaround_second_dialog_start");
    try {
      const secondDialogAuthentication =
        await acquireTokenWithOfficeDialog();
      recordAuthDiagnostic("porkaround_second_dialog_success", {
        acquisitionMode: secondDialogAuthentication?.acquisitionMode || null,
        expiresOn: secondDialogAuthentication?.expiresOn || null
      });
      return secondDialogAuthentication;
    } catch (secondDialogError) {
      recordAuthDiagnostic("porkaround_second_dialog_failed", {
        name: secondDialogError?.name || null,
        errorCode: secondDialogError?.errorCode || null,
        message: secondDialogError?.message || String(secondDialogError)
      });
      throw secondDialogError;
    }
  }
}

async function callHealth() {
  recordAuthDiagnostic("call_health_start");

  const authentication =
    await getAuthentication();

  recordAuthDiagnostic("function_request_start", {
    acquisitionMode: authentication.acquisitionMode
  });

  const response = await fetch(
    AUTH_CONFIG.healthUrl,
    {
      method: "GET",
      headers: {
        Authorization:
          `Bearer ${authentication.accessToken}`,
        Accept:
          "application/json"
      }
    }
  );

  const body =
    await response.text();

  recordAuthDiagnostic("function_response", {
    status: response.status,
    ok: response.ok,
    authenticationMode: authentication.acquisitionMode
  });

  return {
    ok:
      response.ok,
    status:
      response.status,
    statusText:
      response.statusText,
    authenticationMode:
      authentication.acquisitionMode,
    body
  };
}

window.WorkflowHealth = {
  callHealth,
  getDiagnostics: readAuthDiagnostics,
  clearDiagnostics: clearAuthDiagnostics
};
