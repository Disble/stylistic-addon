import { BetterAuthAdapter } from "../adapters/auth/BetterAuthAdapter";
import { MASTRA_BASE_URL } from "../infrastructure/config";
import type { AuthBridgeSessionResponse } from "./auth-dialog.types";

/**
 * Posts a typed auth result from the dialog runtime back to the taskpane.
 *
 * This function intentionally fails loudly when Office.js is missing. Optional
 * chaining hid a real host integration bug during the OAuth implementation: the
 * dialog had a valid session but never notified the parent taskpane.
 */
function postDialogMessage(message: unknown): void {
  const ui = globalThis.Office?.context?.ui as
    | { messageParent?: (message: string) => void }
    | undefined;

  if (!ui?.messageParent) {
    throw new Error("Office Dialog API no está disponible para enviar el resultado al taskpane.");
  }

  ui.messageParent(JSON.stringify(message));
}

/**
 * Waits for Office.js before the dialog uses messageParent or redirects.
 *
 * The official Office fallback-auth sample wraps dialog work in `Office.onReady`.
 * We do the same before both phases: creating the provider URL and exchanging the
 * backend bridge code. That keeps the dialog runtime ready for the final
 * `messageParent` call regardless of provider redirects.
 */
function waitForOfficeReady(): Promise<void> {
  return new Promise((resolve, reject) => {
    const timeoutId = globalThis.setTimeout(() => {
      reject(new Error("Office.onReady no terminó dentro del timeout en auth-dialog."));
    }, 5_000);

    if (!globalThis.Office?.onReady) {
      globalThis.clearTimeout(timeoutId);
      reject(new Error("Office.js no expuso Office.onReady en auth-dialog."));
      return;
    }

    Office.onReady(() => {
      globalThis.clearTimeout(timeoutId);

      if (!globalThis.Office?.context?.ui?.messageParent) {
        reject(new Error("Office.context.ui.messageParent no está disponible en auth-dialog."));
        return;
      }

      resolve();
    });
  });
}

/**
 * Exchanges a one-time backend bridge code for the authenticated session.
 *
 * The backend creates this code at `/auth-complete` after Better Auth can read
 * the callback cookie. The add-in origin then consumes the code exactly once so
 * the taskpane receives a normal `AuthSession` without relying on cross-origin
 * cookies inside the Office WebView.
 */
async function exchangeBridgeCode(bridgeCode: string): Promise<unknown> {
  const url = new URL("/auth-bridge-session", MASTRA_BASE_URL);
  url.searchParams.set("code", bridgeCode);

  const response = await fetch(url.toString());
  const payload = (await response.json()) as AuthBridgeSessionResponse;
  if (!response.ok || !payload.session) {
    throw new Error(payload.error ?? "No se pudo recuperar la sesión del bridge.");
  }

  return payload.session;
}

/**
 * Runs the two-phase OAuth dialog bridge page.
 *
 * Phase 1 receives `callbackUrl`, creates the Google provider URL inside the
 * dialog runtime, and redirects. Phase 2 receives `authBridgeCode`, exchanges it
 * for the Better Auth bearer session, and posts the session to the parent
 * taskpane.
 */
async function runAuthDialog(): Promise<void> {
  const params = new URLSearchParams(globalThis.location.search);
  const callbackUrl = params.get("callbackUrl");
  const bridgeCode = params.get("authBridgeCode");

  await waitForOfficeReady();

  if (bridgeCode) {
    try {
      const session = await exchangeBridgeCode(bridgeCode);
      postDialogMessage({ type: "stylistic-auth-success", session });
    } catch (error) {
      postDialogMessage({
        type: "stylistic-auth-error",
        message: error instanceof Error ? error.message : "No se pudo recuperar la sesión.",
      });
    }
    return;
  }

  if (callbackUrl) {
    const signIn = await new BetterAuthAdapter().createSocialSignInRequest(callbackUrl);
    globalThis.location.assign(signIn.url);
    return;
  }

  postDialogMessage({
    type: "stylistic-auth-error",
    message: "El diálogo de autenticación no recibió una acción válida.",
  });
}

void runAuthDialog();
