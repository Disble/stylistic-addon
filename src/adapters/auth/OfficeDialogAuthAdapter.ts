import type { AuthSession } from "../../domain/auth/AuthSession.types";
import { AUTH_COMPLETE_URL, AUTH_DIALOG_PATH } from "../../infrastructure/config";
import type { DialogMessage } from "./OfficeDialogAuthAdapter.types";

/**
 * Coordinates OAuth through the Office Dialog API.
 *
 * Office requires the first dialog URL to share the taskpane origin, so the
 * taskpane opens local `auth-dialog.html` first. That page then redirects through
 * Better Auth/Google and eventually returns to the same local page with a
 * backend-issued one-time bridge code. Only the local dialog page calls
 * `Office.context.ui.messageParent`, matching the official fallback-auth sample
 * and avoiding cross-origin dialog messaging quirks.
 */
export class OfficeDialogAuthAdapter {
  /** Opens the provider sign-in dialog and resolves with the persisted session. */
  async signIn(): Promise<AuthSession> {
    const dialogUrl = this.buildDialogStartUrl();

    return new Promise<AuthSession>((resolve, reject) => {
      const ui = globalThis.Office?.context?.ui;
      if (!ui?.displayDialogAsync) {
        reject(new Error("Office Dialog API no está disponible en este host."));
        return;
      }

      ui.displayDialogAsync(
        dialogUrl,
        { height: 60, width: 40, displayInIframe: false },
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
            reject(
              new Error(result.error?.message ?? "No se pudo abrir el diálogo de autenticación.")
            );
            return;
          }

          const dialog = result.value;
          const close = () => dialog.close();

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (event) => {
            const message = this.parseDialogMessage((event as { message?: string }).message);
            close();

            if (!message) {
              reject(new Error("El diálogo de autenticación devolvió una respuesta inválida."));
              return;
            }

            if (message.type === "stylistic-auth-error") {
              reject(new Error(message.message));
              return;
            }

            resolve(message.session);
          });

          // Do not attach DialogEventReceived here. In real Word hosts, event
          // 12006 can be emitted while the dialog is still navigating and later
          // produces a valid DialogMessageReceived payload. The official Office
          // fallback-auth sample resolves the parent only from messages or from
          // displayDialogAsync failure; keep that discipline.
        }
      );
    });
  }

  /** Builds the backend callback URL that will mint the one-time bridge code. */
  private buildDialogCallbackUrl(): string {
    const url = new URL(AUTH_COMPLETE_URL);
    url.searchParams.set("parentOrigin", globalThis.location.origin);
    return url.toString();
  }

  /** Builds the first Office Dialog URL, which must be same-origin with taskpane. */
  private buildDialogStartUrl(): string {
    const url = new URL(AUTH_DIALOG_PATH, globalThis.location.origin);
    url.searchParams.set("callbackUrl", this.buildDialogCallbackUrl());
    return url.toString();
  }

  /** Parses the only message contract accepted from the dialog runtime. */
  private parseDialogMessage(raw: string | undefined): DialogMessage | undefined {
    if (!raw) {
      return undefined;
    }
    try {
      return JSON.parse(raw) as DialogMessage;
    } catch {
      return undefined;
    }
  }
}
