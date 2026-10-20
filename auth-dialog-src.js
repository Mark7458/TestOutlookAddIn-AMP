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
const MESSAGE_PARENT_TIMEOUT_MS = 15000;
const MESSAGE_PARENT_POLL_MS = 100;

let msalInstance = null;

function readDiagnosticEvents() {
  try {
    const parsed = JSON.parse(
      localStorage.getItem(DIAGNOSTIC_KEY) ||
        "[]"
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
        events.slice(
          -MAX_DIAGNOSTIC_EVENTS
        )
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
    document.getElementById(
      "auth-diagnostics"
    );

  if (!target) {
    return;
  }

  const events = readDiagnosticEvents();

  target.textContent = events.length
    ? events.map((entry) => {
        const details = entry.details
          ? ` | ${Object.entries(
              entry.details
            )
              .map(([key, value]) =>
                `${key}=${String(value)}`
              )
              .join(" | ")}`
          : "";

        return `${entry.time} | ${entry.event}${details}`;
      }).join("\n")
    : "Nessun evento registrato.";

  target.scrollTop =
    target.scrollHeight;
}

function recordDiagnostic(
  event,
  details = null
) {
  const events =
    readDiagnosticEvents();

  events.push({
    time:
      new Date().toLocaleTimeString(
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

function bindDiagnosticControls() {
  refreshDiagnosticView();

  const clearButton =
    document.getElementById(
      "clear-auth-diagnostics"
    );

  if (!clearButton) {
    return;
  }

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

function delay(milliseconds) {
  return new Promise((resolve) => {
    setTimeout(resolve, milliseconds);
  });
}

/*
 * Office.onReady() nel dialog può restare appeso
 * dopo i redirect. Aspettiamo quindi direttamente
 * e con timeout la sola API che ci serve.
 */
async function waitForMessageParent() {
  const startedAt = Date.now();
  let attempts = 0;

  recordDiagnostic(
    "message_parent_wait_start",
    {
      timeoutMs:
        MESSAGE_PARENT_TIMEOUT_MS,
      pollMs:
        MESSAGE_PARENT_POLL_MS
    }
  );

  while (
    Date.now() - startedAt <
      MESSAGE_PARENT_TIMEOUT_MS
  ) {
    attempts += 1;

    if (
      typeof globalThis.Office !==
        "undefined" &&
      globalThis.Office?.context?.ui &&
      typeof globalThis.Office.context.ui
        .messageParent === "function"
    ) {
      recordDiagnostic(
        "message_parent_available",
        {
          attempts,
          elapsedMs:
            Date.now() - startedAt
        }
      );
      return;
    }

    await delay(
      MESSAGE_PARENT_POLL_MS
    );
  }

  throw new Error(
    "Office.context.ui.messageParent non è diventato disponibile entro il timeout."
  );
}

async function sendMessageToParent(message) {
  await waitForMessageParent();

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

async function sendErrorToParent(
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

  try {
    await sendMessageToParent({
      type:
        "auth-error",
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
  } catch (messageError) {
    recordDiagnostic(
      "message_parent_error_while_reporting_error",
      getSafeErrorDetails(
        messageError
      )
    );
  }
}

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

async function sendTokenToParent(
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
        authenticationResult
          .accessToken.length,
      username:
        authenticationResult
          .account?.username ||
        null,
      homeAccountId:
        authenticationResult
          .account?.homeAccountId ||
        null,
      tenantId:
        authenticationResult
          .account?.tenantId ||
        null
    }
  );

  await sendMessageToParent({
    type:
      "auth-success",
    acquisitionMode,
    accessToken:
      authenticationResult.accessToken
  });
}

async function startAuthentication() {
  recordDiagnostic(
    "script_start",
    {
      page:
        window.location.href
          .split("#")[0],
      hasHash:
        Boolean(
          window.location.hash
        ),
      hasHostInfo:
        new URL(window.location.href)
          .searchParams
          .has("_host_Info")
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
      await client
        .handleRedirectPromise();

    recordDiagnostic(
      "handle_redirect_complete",
      {
        resultPresent:
          Boolean(redirectResult),
        tokenPresent:
          Boolean(
            redirectResult?.accessToken
          ),
        username:
          redirectResult
            ?.account?.username ||
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
    await sendTokenToParent(
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
    let silentResult = null;

    /*
     * Il catch comprende SOLO acquireTokenSilent.
     * Un eventuale errore di messageParent non deve
     * trasformarsi in un nuovo redirect interattivo.
     */
    try {
      recordDiagnostic(
        "silent_token_start",
        {
          username:
            accounts[0].username ||
            null
        }
      );

      silentResult =
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

    if (silentResult?.accessToken) {
      await sendTokenToParent(
        silentResult,
        "silent"
      );
      return;
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

if (document.readyState === "loading") {
  document.addEventListener(
    "DOMContentLoaded",
    bindDiagnosticControls,
    { once: true }
  );
} else {
  bindDiagnosticControls();
}

recordDiagnostic(
  "bundle_loaded"
);

startAuthentication()
  .catch(async (error) => {
    console.error(
      "Errore nel dialogo di autenticazione.",
      error
    );

    await sendErrorToParent(
      error,
      "authDialog"
    );
  });
