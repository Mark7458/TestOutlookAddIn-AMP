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
    "https://jakdrizzt.github.io/TestOutlookAddIn/auth-dialog.html"
};

let msalInstance = null;

/*
 * Invia un oggetto JSON dal dialogo
 * al task pane principale.
 */
function sendMessageToParent(message) {
  Office.context.ui.messageParent(
    JSON.stringify(message)
  );
}

/*
 * Restituisce al task pane un errore
 * serializzato e leggibile.
 */
function sendErrorToParent(
  error,
  authenticationPhase
) {
  sendMessageToParent({
    type: "auth-error",

    authenticationPhase:
      authenticationPhase,

    name:
      error?.name ||
      "AuthenticationError",

    message:
      error?.message ||
      String(error),

    errorCode:
      error?.errorCode ||
      null,

    subError:
      error?.subError ||
      null,

    correlationId:
      error?.correlationId ||
      null
  });
}

/*
 * Crea una sola istanza MSAL.
 *
 * L'authority punta esplicitamente
 * al tenant Azure di test Avanade.
 *
 * L'account SOGEI verrà autenticato
 * come Guest in quel tenant.
 */
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
          "sessionStorage",

        storeAuthStateInCookie:
          true
      }
    });

  await msalInstance.initialize();

  return msalInstance;
}

/*
 * Restituisce il token al task pane.
 */
function sendTokenToParent(
  authenticationResult,
  acquisitionMode
) {
  if (!authenticationResult?.accessToken) {
    throw new Error(
      "MSAL non ha restituito un access token."
    );
  }

  sendMessageToParent({
    type:
      "auth-success",

    acquisitionMode:
      acquisitionMode,

    accessToken:
      authenticationResult.accessToken
  });
}

/*
 * Avvia o completa l'autenticazione.
 */
async function startAuthentication() {
  const client =
    await initializeMsal();

  /*
   * Se questa pagina è stata ricaricata
   * dopo il redirect Microsoft, qui MSAL
   * completa il flusso e recupera il token.
   */
  const redirectResult =
    await client.handleRedirectPromise();

  if (redirectResult?.accessToken) {
    sendTokenToParent(
      redirectResult,
      "interactive"
    );

    return;
  }

  /*
   * Se nel dialogo esiste già un account
   * nella cache MSAL, proviamo prima senza
   * mostrare una nuova autenticazione.
   */
  const accounts =
    client.getAllAccounts();

  if (accounts.length > 0) {
    try {
      const silentResult =
        await client.acquireTokenSilent({
          account:
            accounts[0],

          scopes: [
            AUTH_CONFIG.scope
          ]
        });

      sendTokenToParent(
        silentResult,
        "silent"
      );

      return;
    } catch (silentError) {
      console.warn(
        "Acquisizione silenziosa non riuscita.",
        silentError
      );
    }
  }

  /*
   * Nessun token utilizzabile:
   * navighiamo verso Microsoft Entra.
   *
   * select_account permette all'utente
   * di scegliere l'account SOGEI corretto.
   */
  await client.acquireTokenRedirect({
    scopes: [
      AUTH_CONFIG.scope
    ],

    redirectUri:
      AUTH_CONFIG.redirectUri,

    prompt:
      "select_account"
  });
}

/*
 * Office.js deve essere pronto prima
 * di poter usare messageParent().
 */
Office.onReady(() => {
  startAuthentication()
    .catch((error) => {
      console.error(
        "Errore nel dialogo di autenticazione.",
        error
      );

      sendErrorToParent(
        error,
        "authDialog"
      );
    });
});