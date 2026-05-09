/**
 * Backend path for the user correction-preferences endpoint.
 *
 * Convention: Mastra-native endpoints (workflows, openapi.json, etc.) live
 * under `/api/...`. Custom non-Mastra endpoints — like this one — are mounted
 * at root without the `/api` prefix. Do not add `/api` here unless the backend
 * confirms the route was migrated under the Mastra-native namespace.
 */
export const USER_PREFERENCES_PATH = "/user/preferences";

/** Distinguishable backend error code returned for shape/length validation. */
export const INVALID_USER_PREFERENCES_REQUEST_ERROR = "invalid_user_preferences_request";

/** Distinguishable backend error code returned for invalid JSON bodies. */
export const INVALID_JSON_BODY_ERROR = "invalid_json_body";
