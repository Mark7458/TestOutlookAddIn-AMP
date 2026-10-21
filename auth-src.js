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

  console.info(
    "Account MSAL disponibili nella cache:",
    accounts.length
  );

  if (accounts.length === 0) {
    return null;
  }

  try {
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
          dialogOperationSequence += 1;
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

        const authenticationTimeout =
          setTimeout(() => {
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
          resolve(tokenResult);
        }

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (event) => {
            let message;

            try {
              message =
                JSON.parse(event.message);
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
            console.warn(
              "Evento Office Dialog ricevuto:",
              {
                error:
                  event?.error
              }
            );
          }
        );
      }
    );
  });
}

async function getAuthentication() {
  if (isAuthenticationUsable(currentAuthentication)) {
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

  return acquireTokenWithOfficeDialog();
}

async function callHealth() {
  const authentication =
    await getAuthentication();

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
  callHealth
};
