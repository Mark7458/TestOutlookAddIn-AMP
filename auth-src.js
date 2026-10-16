const AUTH_CONFIG = {
  healthUrl:
    "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/health",

  dialogUrl:
    "https://jakdrizzt.github.io/TestOutlookAddIn/auth-dialog.html"
};

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

function acquireTokenWithOfficeDialog() {
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
        displayInIframe: false
      },
      (result) => {
        if (
          result.status ===
          Office.AsyncResultStatus.Failed
        ) {
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

        /*
         * Se non riceviamo alcun messaggio dal dialogo
         * entro cinque minuti, terminiamo il tentativo.
         */
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
          if (completed) {
            return;
          }

          completed = true;

          clearAuthenticationTimeout();
          closeDialog();
          reject(error);
        }

        function completeWithToken(tokenResult) {
          if (completed) {
            return;
          }

          completed = true;

          clearAuthenticationTimeout();
          closeDialog();
          resolve(tokenResult);
        }

        /*
         * Questo è l'evento autorevole:
         * il dialogo invia un JSON tramite messageParent().
         */
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
              message.type === "auth-success" &&
              message.accessToken
            ) {
              completeWithToken({
                accessToken:
                  message.accessToken,

                acquisitionMode:
                  message.acquisitionMode ||
                  "interactive"
              });

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

        /*
         * In questo ambiente DialogEventReceived viene
         * emesso anche mentre la finestra di autenticazione
         * è ancora visibile.
         *
         * Non lo usiamo quindi per chiudere la Promise.
         * Lo registriamo soltanto a fini diagnostici.
         */
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

async function callHealth() {
  const authentication =
    await acquireTokenWithOfficeDialog();

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