import { CORRECTION_INSTRUCTIONS_MAX_LENGTH } from "../../../domain/user-preferences/UserCorrectionPreferences.constants";
import { DEFAULT_PROFILES } from "../../../infrastructure/config";
import type { AnalysisProfileOption } from "../AnalysisProfileSection";

/** Selectable analysis-profile options derived from the canonical profile list. */
export const ANALYSIS_PROFILE_OPTIONS: readonly AnalysisProfileOption[] = DEFAULT_PROFILES.map(
  (profile) => ({
    value: profile.id,
    label: profile.label,
  })
);

/**
 * Initial fallback for the correction-instructions max length used while the
 * GET request is in flight. The backend echoes the authoritative value.
 */
export const INITIAL_CORRECTION_INSTRUCTIONS_MAX_LENGTH = CORRECTION_INSTRUCTIONS_MAX_LENGTH;

/** Inline + status-bar message shown when the save flow fails generically. */
export const SETTINGS_SAVE_GENERIC_ERROR_MESSAGE = "No se pudo guardar la configuración.";

/** Status-bar success message shown after a clean save. */
export const SETTINGS_SAVE_SUCCESS_MESSAGE = "Configuración guardada.";

/** Specific message used when the backend rejects the request as invalid. */
export const SETTINGS_SAVE_INVALID_REQUEST_MESSAGE =
  "Las instrucciones no cumplen el formato esperado.";

/** Specific message used when the backend rejects the request as unauthenticated. */
export const SETTINGS_SAVE_UNAUTHENTICATED_MESSAGE = "Tu sesión expiró. Iniciá sesión nuevamente.";

/** Specific message used when the network is unreachable. */
export const SETTINGS_SAVE_NETWORK_ERROR_MESSAGE =
  "No pudimos conectar con el servidor. Verificá tu conexión.";
