/* global console */

import type { IUserCorrectionPreferencesPort } from "../../domain/ports";
import type { UserCorrectionPreferences } from "../../domain/user-preferences/UserCorrectionPreferences.types";
import { UserCorrectionPreferencesError } from "../../domain/user-preferences/UserCorrectionPreferencesError";
import type { HttpClient } from "../../infrastructure/http/HttpClient";
import { HttpError } from "../../infrastructure/http/HttpError";
import { USER_PREFERENCES_PATH } from "./BackendUserCorrectionPreferencesAdapter.constants";

/**
 * Backend-backed implementation of {@link IUserCorrectionPreferencesPort}.
 *
 * Owns the translation between transport-level signals (HTTP status, fetch
 * rejections) and domain-meaningful failure reasons. The presentation layer
 * never sees raw `HttpError` instances — only typed
 * {@link UserCorrectionPreferencesError} values.
 */
export class BackendUserCorrectionPreferencesAdapter implements IUserCorrectionPreferencesPort {
  constructor(private readonly httpClient: HttpClient) {}

  async load(): Promise<UserCorrectionPreferences> {
    try {
      return await this.httpClient.get<UserCorrectionPreferences>(USER_PREFERENCES_PATH);
    } catch (error) {
      throw this.toDomainError(error);
    }
  }

  async save(correctionInstructions: string | null): Promise<UserCorrectionPreferences> {
    try {
      return await this.httpClient.put<UserCorrectionPreferences>(USER_PREFERENCES_PATH, {
        correctionInstructions,
      });
    } catch (error) {
      throw this.toDomainError(error);
    }
  }

  private toDomainError(error: unknown): UserCorrectionPreferencesError {
    if (error instanceof UserCorrectionPreferencesError) {
      return error;
    }
    if (error instanceof HttpError) {
      console.error(
        `🔴 [UserCorrectionPreferencesAdapter] HTTP ${error.status} ${error.statusText} — body:`,
        error.body
      );
      if (error.status === 401) {
        return new UserCorrectionPreferencesError("unauthenticated");
      }
      if (error.status === 400) {
        return new UserCorrectionPreferencesError("invalid-request");
      }
      return new UserCorrectionPreferencesError("unknown", `HTTP ${error.status}`);
    }
    if (error instanceof TypeError) {
      console.error("🔴 [UserCorrectionPreferencesAdapter] Network error:", error);
      return new UserCorrectionPreferencesError("network", error.message);
    }
    if (error instanceof Error) {
      console.error("🔴 [UserCorrectionPreferencesAdapter] Unexpected error:", error);
      return new UserCorrectionPreferencesError("unknown", error.message);
    }
    console.error("🔴 [UserCorrectionPreferencesAdapter] Non-Error thrown:", error);
    return new UserCorrectionPreferencesError("unknown");
  }
}
