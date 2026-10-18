import { DEFAULT_PROFILES } from "../../../infrastructure/config";
import type { AnalysisProfileOption } from "../AnalysisProfileSection";

export { CORRECTION_INSTRUCTIONS_MAX_LENGTH as INITIAL_CORRECTION_INSTRUCTIONS_MAX_LENGTH } from "../../../domain/user-preferences/UserCorrectionPreferences.constants";

/** Selectable analysis-profile options derived from the canonical profile list. */
export const ANALYSIS_PROFILE_OPTIONS: readonly AnalysisProfileOption[] = DEFAULT_PROFILES.map(
  (profile) => ({
    value: profile.id,
    label: profile.label,
  })
);

/** Inline + status-bar message shown when the save flow fails generically. */
export const SETTINGS_SAVE_GENERIC_ERROR_MESSAGE = "No se pudo guardar la configuración.";

/** Status-bar success message shown after a clean save. */
export const SETTINGS_SAVE_SUCCESS_MESSAGE = "Configuración guardada.";

/** Specific message used when the backend rejects the request as invalid. */
export const SETTINGS_SAVE_INVALID_REQUEST_MESSAGE =
  "Las instrucciones no cumplen el formato esperado.";

/** Specific message used when the hub rejects the task pane's token (missing or wrong). */
export const SETTINGS_SAVE_UNAUTHENTICATED_MESSAGE =
  "Stylistic Hub rechazó el token de este panel. Ábrelo de nuevo desde el hub.";

/** Specific message used when the network is unreachable. */
export const SETTINGS_SAVE_NETWORK_ERROR_MESSAGE =
  "No pudimos conectar con el servidor. Verificá tu conexión.";
