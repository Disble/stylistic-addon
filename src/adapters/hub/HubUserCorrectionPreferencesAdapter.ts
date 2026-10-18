/* global console */

import type { IUserCorrectionPreferencesPort } from "../../domain/ports";
import type { UserCorrectionPreferences } from "../../domain/user-preferences/UserCorrectionPreferences.types";
import { UserCorrectionPreferencesError } from "../../domain/user-preferences/UserCorrectionPreferencesError";
import type { HttpClient } from "../../infrastructure/http/HttpClient";
import { HttpError } from "../../infrastructure/http/HttpError";
import { HUB_PREFERENCES_PATH } from "./HubUserCorrectionPreferencesAdapter.constants";
import { readHubErrorDetail } from "./HubProblemParser.helpers";

/**
 * Hub-backed implementation of {@link IUserCorrectionPreferencesPort}
 * (`GET`/`PUT /preferences` on Stylistic Hub API v1).
 *
 * Owns the translation between transport-level signals (HTTP status, fetch
 * rejections) and domain-meaningful failure reasons. The presentation layer
 * never sees raw `HttpError` instances — only typed
 * {@link UserCorrectionPreferencesError} values.
 */
export class HubUserCorrectionPreferencesAdapter implements IUserCorrectionPreferencesPort {
  constructor(private readonly httpClient: HttpClient) {}

  async load(): Promise<UserCorrectionPreferences> {
    try {
      return await this.httpClient.get<UserCorrectionPreferences>(HUB_PREFERENCES_PATH);
    } catch (error) {
      throw this.toDomainError(error);
    }
  }

  async save(correctionInstructions: string | null): Promise<UserCorrectionPreferences> {
    try {
      return await this.httpClient.put<UserCorrectionPreferences>(HUB_PREFERENCES_PATH, {
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
      const detail = readHubErrorDetail(error);
      if (error.status === 401) {
        return new UserCorrectionPreferencesError("unauthenticated", detail);
      }
      if (error.status === 400) {
        return new UserCorrectionPreferencesError("invalid-request", detail);
      }
      return new UserCorrectionPreferencesError("unknown", detail);
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
