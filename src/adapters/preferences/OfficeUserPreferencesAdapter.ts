/* global OfficeRuntime */

import type { IUserPreferencesPort } from "../../domain/ports";
import { ANALYSIS_PROFILE_STORAGE_KEY } from "../../infrastructure/config";

/**
 * Persists user-level preferences using OfficeRuntime.storage.
 *
 * OfficeRuntime.storage is per-user and cross-document, which matches the
 * semantics of "user preference" exactly — the chosen analysis profile follows
 * the user across files, not the document. There is intentionally no
 * localStorage fallback: hosts that cannot provide global add-in storage must
 * surface a controlled unsupported state, not silently weaken persistence.
 */
export class OfficeUserPreferencesAdapter implements IUserPreferencesPort {
  /** Reads the persisted analysis-profile id, returning undefined when absent or invalid. */
  async getAnalysisProfile(): Promise<string | undefined> {
    const raw = await this.getStorage().getItem(ANALYSIS_PROFILE_STORAGE_KEY);
    if (typeof raw !== "string" || raw.length === 0) {
      return undefined;
    }
    return raw;
  }

  /** Persists the analysis-profile id under the dedicated storage key. */
  async setAnalysisProfile(value: string): Promise<void> {
    await this.getStorage().setItem(ANALYSIS_PROFILE_STORAGE_KEY, value);
  }

  private getStorage(): OfficeRuntime.Storage {
    const storage = globalThis.OfficeRuntime?.storage;
    if (!storage) {
      throw new Error(
        "Tu versión de Office no soporta persistencia de preferencias para Stylistic."
      );
    }
    return storage;
  }
}
