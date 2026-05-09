/**
 * Hard limit enforced by the backend on `correctionInstructions`.
 *
 * The backend echoes this in every response so the frontend can derive the
 * counter from the runtime value. This local constant mirrors that contract
 * for places where the runtime response is not yet available (e.g., initial
 * form scaffolding before the first GET resolves).
 */
export const CORRECTION_INSTRUCTIONS_MAX_LENGTH = 4000;
