// Test-Plugin file version: 0.2-auth-popup-redirect-src.js-1
import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge";

(async () => {
  try {
    await broadcastResponseToMainFrame();
  } catch (error) {
    console.error("MSAL popup redirect bridge non riuscito.", error);
    const status = document.getElementById("bridge-status");
    if (status) status.textContent = "Impossibile completare il ritorno dell'autenticazione.";
  }
})();
