import {
  createNestablePublicClientApplication,
  InteractionRequiredAuthError
} from "@azure/msal-browser";

const AUTH_CONFIG = {
  clientId: "17be48ae-1d05-4233-ba98-654102bf76e0",

  tenantId: "9284cd5c-ae26-4eb2-b079-a14b6df5f1aa",

  scope:
    "api://17be48ae-1d05-4233-ba98-654102bf76e0/Workflow.Access",

  healthUrl:
    "https://jakfunction-bzh9g2e6amhtdgff.italynorth-01.azurewebsites.net/api/health"
};

let msalInstance;

async function initializeMsal() {
  if (msalInstance) {
    return msalInstance;
  }

  msalInstance =
    await createNestablePublicClientApplication({
      auth: {
        clientId: AUTH_CONFIG.clientId,

        authority:
          `https://login.microsoftonline.com/${AUTH_CONFIG.tenantId}`
      },

      cache: {
        cacheLocation: "localStorage"
      }
    });

  return msalInstance;
}

async function getWorkflowAccessToken() {
  const client = await initializeMsal();

  const tokenRequest = {
    scopes: [AUTH_CONFIG.scope]
  };

  try {
    const result =
      await client.acquireTokenSilent(tokenRequest);

    return {
      accessToken: result.accessToken,
      acquisitionMode: "silent"
    };
  } catch (silentError) {
    if (
      !(silentError instanceof InteractionRequiredAuthError)
    ) {
      throw silentError;
    }

    const result =
      await client.acquireTokenPopup(tokenRequest);

    return {
      accessToken: result.accessToken,
      acquisitionMode: "interactive"
    };
  }
}

async function callHealth() {
  const authentication =
    await getWorkflowAccessToken();

  const response = await fetch(
    AUTH_CONFIG.healthUrl,
    {
      method: "GET",

      headers: {
        Authorization:
          `Bearer ${authentication.accessToken}`,

        Accept: "application/json"
      }
    }
  );

  const body = await response.text();

  return {
    ok: response.ok,
    status: response.status,
    statusText: response.statusText,
    authenticationMode:
      authentication.acquisitionMode,
    body
  };
}

window.WorkflowHealth = {
  callHealth
};