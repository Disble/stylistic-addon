import type { AuthSession } from "../../domain/auth/AuthSession.types";
import { AUTH_COMPLETE_URL, AUTH_DIALOG_PATH } from "../../infrastructure/config";

type DialogMessage = Readonly<
  | { type: "stylistic-auth-success"; session: AuthSession }
  | { type: "stylistic-auth-error"; message: string }
>;

/**
 * Coordinates OAuth through the Office Dialog API.
 *
 * Office requires the first dialog URL to share the taskpane origin. The dialog
 * page then redirects to Better Auth/provider and posts the final session back
 * with `Office.context.ui.messageParent`.
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

          // The official Office fallback-auth sample only resolves from
          // DialogMessageReceived. DialogEventReceived can surface transient host
          // events such as 12006 while the dialog is still navigating, which made
          // the taskpane mark auth as failed even though the dialog later produced
          // a valid bridge session.
        }
      );
    });
  }

  private buildDialogCallbackUrl(): string {
    const url = new URL(AUTH_COMPLETE_URL);
    url.searchParams.set("parentOrigin", globalThis.location.origin);
    return url.toString();
  }

  private buildDialogStartUrl(): string {
    const url = new URL(AUTH_DIALOG_PATH, globalThis.location.origin);
    url.searchParams.set("callbackUrl", this.buildDialogCallbackUrl());
    return url.toString();
  }

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
