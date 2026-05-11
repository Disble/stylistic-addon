/**
 * User-level correction preferences — global guidance the user declares
 * explicitly to focus the corrector on known recurring issues.
 *
 * These preferences live at the backend boundary, not in the document. They
 * never participate in the document profile or the per-suggestion feedback
 * loop. The taskpane only owns the form state used to read and update them.
 */

/**
 * Snapshot returned by the backend for the user's stored preferences.
 *
 * `correctionInstructions === null` means the user has no global guidance set.
 * `correctionInstructionsMaxLength` is echoed by the backend so the UI can
 * derive the visual counter and validation thresholds without hard-coding the
 * limit in two places.
 */
export type UserCorrectionPreferences = Readonly<{
  correctionInstructions: string | null;
  correctionInstructionsMaxLength: number;
}>;

/**
 * Distinguishable failures returned by the user-correction-preferences port.
 *
 * Adapters MUST translate transport-level errors into one of these reasons so
 * the presentation layer can render a tailored message instead of inspecting
 * raw HTTP semantics.
 */
export type UserCorrectionPreferencesErrorReason =
  | "unauthenticated"
  | "invalid-request"
  | "network"
  | "unknown";
