import type { UserCorrectionPreferencesErrorReason } from "./UserCorrectionPreferences.types";

/**
 * Structured failure thrown by adapters implementing
 * `IUserCorrectionPreferencesPort`.
 *
 * Reasons map roughly to backend error codes documented in the contract:
 * - `unauthenticated` → 401, the bearer token is missing or invalid.
 * - `invalid-request` → 400 with body validation issues.
 * - `network` → transport-level failure (no response, abort, fetch reject).
 * - `unknown` → unmapped error worth surfacing without leaking transport detail.
 */
export class UserCorrectionPreferencesError extends Error {
  public readonly reason: UserCorrectionPreferencesErrorReason;

  constructor(reason: UserCorrectionPreferencesErrorReason, message?: string) {
    super(message ?? reason);
    this.name = "UserCorrectionPreferencesError";
    this.reason = reason;
  }
}
