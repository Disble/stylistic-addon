import { UserCorrectionPreferencesError } from "../../../domain/user-preferences/UserCorrectionPreferencesError";
import {
  SETTINGS_SAVE_GENERIC_ERROR_MESSAGE,
  SETTINGS_SAVE_INVALID_REQUEST_MESSAGE,
  SETTINGS_SAVE_NETWORK_ERROR_MESSAGE,
  SETTINGS_SAVE_UNAUTHENTICATED_MESSAGE,
} from "./SettingsView.constants";

/**
 * Trims the textarea draft and converts a whitespace-only or empty value to
 * the `null` payload the backend expects to "clear" the persisted preferences.
 */
export function normalizeCorrectionInstructions(value: string): string | null {
  const trimmed = value.trim();
  return trimmed.length === 0 ? null : trimmed;
}

/**
 * Returns the saved value as a string suitable for the textarea draft.
 * `null` collapses to an empty string so the input remains a controlled value.
 */
export function correctionInstructionsToDraft(value: string | null): string {
  return value ?? "";
}

/** Maps a domain failure into the user-facing message rendered in the form. */
export function describeSaveError(error: unknown): string {
  if (error instanceof UserCorrectionPreferencesError) {
    switch (error.reason) {
      case "unauthenticated":
        return SETTINGS_SAVE_UNAUTHENTICATED_MESSAGE;
      case "invalid-request":
        return SETTINGS_SAVE_INVALID_REQUEST_MESSAGE;
      case "network":
        return SETTINGS_SAVE_NETWORK_ERROR_MESSAGE;
      default:
        return SETTINGS_SAVE_GENERIC_ERROR_MESSAGE;
    }
  }
  return SETTINGS_SAVE_GENERIC_ERROR_MESSAGE;
}
