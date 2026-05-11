/* global Office */

import { DOCUMENT_UUID_SETTINGS_KEY } from "../../infrastructure/config";

/** Resolves and persists the stable UUID associated with the active Word document. */
export class WordDocumentIdentityAdapter {
  /** Returns the persisted UUID, creating and saving one when missing or invalid. */
  async getDocumentUuid(): Promise<string> {
    const settings = this.getSettings();
    const persistedValue = settings.get(DOCUMENT_UUID_SETTINGS_KEY);
    const normalizedUuid = this.normalizePersistedUuid(persistedValue);

    if (normalizedUuid) {
      return normalizedUuid;
    }

    const documentUuid = this.createDocumentUuid();
    settings.set(DOCUMENT_UUID_SETTINGS_KEY, documentUuid);
    await this.saveSettings(settings);
    return documentUuid;
  }

  /** Reads the Office document settings boundary, failing closed when unsupported. */
  private getSettings(): Office.Settings {
    const settings = globalThis.Office?.context?.document?.settings;
    if (!settings) {
      throw new Error(
        "Tu versión de Office no soporta la identidad persistida de documento para Stylistic."
      );
    }
    return settings;
  }

  /** Normalizes an existing persisted UUID only when it already matches the contract. */
  private normalizePersistedUuid(value: unknown): string | undefined {
    if (typeof value !== "string") {
      return undefined;
    }

    const normalized = value.trim();
    return this.isUuid(normalized) ? normalized : undefined;
  }

  /** Persists the settings mutation and surfaces host failures as regular errors. */
  private saveSettings(settings: Office.Settings): Promise<void> {
    return new Promise((resolve, reject) => {
      settings.saveAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }

        reject(
          new Error(
            asyncResult.error?.message ??
              "No se pudo persistir el documentUuid en la configuración del documento."
          )
        );
      });
    });
  }

  /** Creates a v4 UUID using the host crypto API. */
  private createDocumentUuid(): string {
    if (typeof globalThis.crypto?.randomUUID === "function") {
      return globalThis.crypto.randomUUID();
    }

    if (typeof globalThis.crypto?.getRandomValues !== "function") {
      throw new Error("No hay una API criptográfica disponible para generar documentUuid.");
    }

    const bytes = globalThis.crypto.getRandomValues(new Uint8Array(16));
    bytes[6] = (bytes[6] & 0x0f) | 0x40;
    bytes[8] = (bytes[8] & 0x3f) | 0x80;

    const hex = Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0"));
    return [
      hex.slice(0, 4).join(""),
      hex.slice(4, 6).join(""),
      hex.slice(6, 8).join(""),
      hex.slice(8, 10).join(""),
      hex.slice(10, 16).join(""),
    ].join("-");
  }

  /** Validates that a persisted value still matches UUID formatting. */
  private isUuid(value: string): boolean {
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
  }
}
