/* global OfficeRuntime */

import type { AuthSession } from "../../domain/auth/AuthSession.types";
import type { IAuthSessionStoragePort } from "../../domain/ports";
import { AUTH_SESSION_STORAGE_KEY } from "../../infrastructure/config";

/**
 * Persists auth sessions using OfficeRuntime.storage.
 *
 * This adapter intentionally has no localStorage fallback: if the Office host
 * cannot provide global add-in storage, the taskpane should surface a controlled
 * unsupported-session state instead of silently weakening persistence semantics.
 */
export class OfficeAuthSessionStorageAdapter implements IAuthSessionStoragePort {
  /** Restores the persisted session from OfficeRuntime.storage. */
  async restore(): Promise<AuthSession | undefined> {
    const raw = await this.getStorage().getItem(AUTH_SESSION_STORAGE_KEY);
    if (!raw) {
      return undefined;
    }

    try {
      return this.parseSession(raw);
    } catch {
      await this.clear();
      return undefined;
    }
  }

  /** Persists the current session as a JSON string. */
  async persist(session: AuthSession): Promise<void> {
    await this.getStorage().setItem(AUTH_SESSION_STORAGE_KEY, JSON.stringify(session));
  }

  /** Removes the persisted session from OfficeRuntime.storage. */
  async clear(): Promise<void> {
    await this.getStorage().removeItem(AUTH_SESSION_STORAGE_KEY);
  }

  private getStorage(): OfficeRuntime.Storage {
    const storage = globalThis.OfficeRuntime?.storage;
    if (!storage) {
      throw new Error("Tu versión de Office no soporta persistencia de sesión para Stylistic.");
    }
    return storage;
  }

  private parseSession(raw: string): AuthSession | undefined {
    const parsed = JSON.parse(raw) as Partial<AuthSession>;
    if (!parsed.token || !parsed.user?.id) {
      return undefined;
    }
    return parsed as AuthSession;
  }
}
