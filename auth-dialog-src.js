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

const DIAGNOSTIC_KEY =
  "outlook-sogei-auth-diagnostics";
const MAX_DIAGNOSTIC_EVENTS = 60;
let msalInstance = null;

function readDiagnosticEvents() {
  try {
    const parsed = JSON.parse(
      localStorage.getItem(DIAGNOSTIC_KEY) || "[]"
    );
    return Array.isArray(parsed)
      ? parsed
      : [];
  } catch {
    return [];
  }
}

function writeDiagnosticEvents(events) {
  try {
    localStorage.setItem(
      DIAGNOSTIC_KEY,
      JSON.stringify(
        events.slice(-MAX_DIAGNOSTIC_EVENTS)
      )
    );
  } catch (error) {
    console.warn(
      "Impossibile salvare la diagnostica auth.",
      error
    );
  }
}

function refreshDiagnosticView() {
  const target =
    document.getElementById("auth-diagnostics");

  if (!target) {
    return;
  }

  const events = readDiagnosticEvents();

  target.textContent = events.length
    ? events.map((entry) => {
        const details = entry.details
          ? ` | ${Object.entries(entry.details)
              .map(([key, value]) =>
                `${key}=${String(value)}`
              )
              .join(" | ")}`
          : "";

        return `${entry.time} | ${entry.event}${details}`;
      }).join("\n")
    : "Nessun evento registrato.";

  target.scrollTop = target.scrollHeight;
}

function recordDiagnostic(event, details = null) {
  const events = readDiagnosticEvents();

  events.push({
    time: new Date().toLocaleTimeString(
      "it-IT",
      { hour12: false }
    ),
    event,
    details
  });

  writeDiagnosticEvents(events);
  refreshDiagnosticView();
}

function getSafeErrorDetails(error) {
  return {
    name:
      error?.name ||
      "Error",
    errorCode:
      error?.errorCode ||
      null,
    message:
      error?.message ||
      String(error),
    correlationId:
      error?.correlationId ||
      null
  };
}

/*
 * Invia un oggetto JSON dal dialogo
 * al task pane principale.
 */
function sendMessageToParent(message) {
  recordDiagnostic(
    "message_parent_start",
    {
      type:
        message?.type ||
        "unknown"
    }
  );

  Office.context.ui.messageParent(
    JSON.stringify(message)
  );

  recordDiagnostic(
    "message_parent_sent",
    {
      type:
        message?.type ||
        "unknown"
    }
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
  recordDiagnostic(
    "auth_error",
    {
      phase:
        authenticationPhase,
      ...getSafeErrorDetails(error)
    }
  );

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
 */
async function initializeMsal() {
  if (msalInstance) {
    return msalInstance;
  }

  recordDiagnostic(
    "msal_initialize_start"
  );

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

  recordDiagnostic(
    "msal_initialize_ok"
  );

  return msalInstance;
}

/*
 * Restituisce il token al task pane.
 * Il token NON viene scritto nella diagnostica.
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

  recordDiagnostic(
    "token_ready",
    {
      mode:
        acquisitionMode,
      tokenPresent:
        true,
      tokenLength:
        authenticationResult.accessToken.length,
      username:
        authenticationResult.account?.username ||
        null,
      homeAccountId:
        authenticationResult.account?.homeAccountId ||
        null,
      tenantId:
        authenticationResult.account?.tenantId ||
        null
    }
  );

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
  recordDiagnostic(
    "authentication_start",
    {
      page:
        window.location.href.split("#")[0]
    }
  );

  const client =
    await initializeMsal();

  recordDiagnostic(
    "handle_redirect_start"
  );

  let redirectResult;

  try {
    redirectResult =
      await client.handleRedirectPromise();

    recordDiagnostic(
      "handle_redirect_complete",
      {
        resultPresent:
          Boolean(redirectResult),
        tokenPresent:
          Boolean(redirectResult?.accessToken),
        username:
          redirectResult?.account?.username ||
          null
      }
    );
  } catch (error) {
    recordDiagnostic(
      "handle_redirect_error",
      getSafeErrorDetails(error)
    );
    throw error;
  }

  if (redirectResult?.accessToken) {
    sendTokenToParent(
      redirectResult,
      "interactive"
    );
    return;
  }

  const accounts =
    client.getAllAccounts();

  recordDiagnostic(
    "accounts_from_cache",
    {
      count:
        accounts.length,
      usernames:
        accounts
          .map((account) =>
            account.username || ""
          )
          .filter(Boolean)
          .join(", ")
    }
  );

  if (accounts.length > 0) {
    try {
      recordDiagnostic(
        "silent_token_start",
        {
          username:
            accounts[0].username ||
            null
        }
      );

      const silentResult =
        await client.acquireTokenSilent({
          account:
            accounts[0],
          scopes: [
            AUTH_CONFIG.scope
          ]
        });

      recordDiagnostic(
        "silent_token_complete",
        {
          tokenPresent:
            Boolean(
              silentResult?.accessToken
            )
        }
      );

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

      recordDiagnostic(
        "silent_token_error",
        getSafeErrorDetails(
          silentError
        )
      );
    }
  }

  recordDiagnostic(
    "acquire_token_redirect_start",
    {
      redirectUri:
        AUTH_CONFIG.redirectUri,
      scope:
        AUTH_CONFIG.scope
    }
  );

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
  refreshDiagnosticView();

  recordDiagnostic(
    "office_ready"
  );

  const clearButton =
    document.getElementById(
      "clear-auth-diagnostics"
    );

  if (clearButton) {
    clearButton.addEventListener(
      "click",
      () => {
        localStorage.removeItem(
          DIAGNOSTIC_KEY
        );
        refreshDiagnosticView();
      }
    );
  }

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
