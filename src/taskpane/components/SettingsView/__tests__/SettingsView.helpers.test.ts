import { describe, expect, it } from "vitest";
import { UserCorrectionPreferencesError } from "../../../../domain/user-preferences/UserCorrectionPreferencesError";
import {
  SETTINGS_SAVE_GENERIC_ERROR_MESSAGE,
  SETTINGS_SAVE_INVALID_REQUEST_MESSAGE,
  SETTINGS_SAVE_NETWORK_ERROR_MESSAGE,
  SETTINGS_SAVE_UNAUTHENTICATED_MESSAGE,
} from "../SettingsView.constants";
import {
  correctionInstructionsToDraft,
  describeSaveError,
  normalizeCorrectionInstructions,
} from "../SettingsView.helpers";

describe("normalizeCorrectionInstructions", () => {
  it("returns null for an empty string", () => {
    expect(normalizeCorrectionInstructions("")).toBeNull();
  });

  it("returns null for a whitespace-only value", () => {
    expect(normalizeCorrectionInstructions("   \n\t ")).toBeNull();
  });

  it("trims surrounding whitespace from a real value", () => {
    expect(normalizeCorrectionInstructions("  hola  ")).toBe("hola");
  });

  it("preserves internal whitespace as-is", () => {
    expect(normalizeCorrectionInstructions(" línea 1\nlínea 2 ")).toBe("línea 1\nlínea 2");
  });
});

describe("correctionInstructionsToDraft", () => {
  it("collapses null to an empty string for the controlled textarea", () => {
    expect(correctionInstructionsToDraft(null)).toBe("");
  });

  it("returns the persisted string unchanged", () => {
    expect(correctionInstructionsToDraft("Vigilá X.")).toBe("Vigilá X.");
  });
});

describe("describeSaveError", () => {
  it("maps unauthenticated to the session-expired message", () => {
    const error = new UserCorrectionPreferencesError("unauthenticated");
    expect(describeSaveError(error)).toBe(SETTINGS_SAVE_UNAUTHENTICATED_MESSAGE);
  });

  it("maps invalid-request to the validation message", () => {
    const error = new UserCorrectionPreferencesError("invalid-request");
    expect(describeSaveError(error)).toBe(SETTINGS_SAVE_INVALID_REQUEST_MESSAGE);
  });

  it("maps network errors to the connection message", () => {
    const error = new UserCorrectionPreferencesError("network");
    expect(describeSaveError(error)).toBe(SETTINGS_SAVE_NETWORK_ERROR_MESSAGE);
  });

  it("falls back to the generic message for unknown reasons", () => {
    const error = new UserCorrectionPreferencesError("unknown");
    expect(describeSaveError(error)).toBe(SETTINGS_SAVE_GENERIC_ERROR_MESSAGE);
  });

  it("falls back to the generic message for unrelated error types", () => {
    expect(describeSaveError(new Error("anything"))).toBe(SETTINGS_SAVE_GENERIC_ERROR_MESSAGE);
  });
});
