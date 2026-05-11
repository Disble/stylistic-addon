import type { AnalysisProfileId } from "../domain/Profile.types";
import type { UserCorrectionPreferences } from "../domain/user-preferences/UserCorrectionPreferences.types";
import {
  getTaskpaneAuthToken,
  setTaskpaneAuthenticated,
  setTaskpaneAuthLoading,
  setTaskpaneSigningIn,
  setTaskpaneSigningOut,
  setTaskpaneUnauthenticated,
} from "./TaskpaneAuthStore";
import {
  getTaskpaneShellState,
  setTaskpaneCleanupLoading,
  setTaskpaneCleanupVisible,
  setTaskpaneDisableTrackChangesCtaVisible,
  setTaskpaneDisableTrackChangesLoading,
  showTaskpaneStatus,
} from "./TaskpaneShellStore";
import { toTaskpaneUserMessage } from "./TaskpaneError.helpers";
import type {
  SaveTaskpanePreferencesHandler,
  TaskpaneBootstrapRuntime,
} from "./TaskpaneBootstrap.types";

/** Narrows persisted storage values to supported analysis-profile identifiers. */
function isSupportedAnalysisProfile(
  value: string,
  supportedAnalysisProfiles: readonly AnalysisProfileId[]
): value is AnalysisProfileId {
  return supportedAnalysisProfiles.includes(value as AnalysisProfileId);
}

/** Loads the user's persisted analysis profile into the shell store. */
async function bootstrapAnalysisProfile(runtime: TaskpaneBootstrapRuntime): Promise<void> {
  try {
    const stored = await runtime.userPreferencesPort.getAnalysisProfile();
    if (stored && isSupportedAnalysisProfile(stored, runtime.supportedAnalysisProfiles)) {
      runtime.setSelectedGenero(stored);
    }
  } catch (error) {
    console.warn("⚠️ [Taskpane] No se pudo restaurar el perfil de análisis:", error);
  }
}

/** Restores and validates the persisted Better Auth session during taskpane bootstrap. */
async function bootstrapAuthSession(runtime: TaskpaneBootstrapRuntime): Promise<void> {
  setTaskpaneAuthLoading();
  try {
    const persisted = await runtime.authSessionStoragePort.restore();
    if (!persisted) {
      setTaskpaneUnauthenticated();
      return;
    }

    const validSession = await runtime.authPort.getSession(persisted.token);
    if (!validSession) {
      await runtime.authSessionStoragePort.clear();
      setTaskpaneUnauthenticated();
      return;
    }

    await runtime.authSessionStoragePort.persist(validSession);
    setTaskpaneAuthenticated(validSession);
  } catch (error) {
    setTaskpaneUnauthenticated(toTaskpaneUserMessage(error));
  }
}

/** Syncs cleanup CTA visibility with the authoritative document state. */
export async function refreshCleanupVisibility(runtime: TaskpaneBootstrapRuntime): Promise<void> {
  try {
    const { deletable } = await runtime.documentPort.getCleanupPreview();
    setTaskpaneCleanupVisible(deletable > 0);
  } catch (error) {
    console.warn("⚠️ [Taskpane] No se pudo calcular la visibilidad de limpieza:", error);
  }
}

/** Rehydrates Track Changes CTA visibility from the review-session mediator. */
async function refreshTrackChangesCtaVisibility(runtime: TaskpaneBootstrapRuntime): Promise<void> {
  try {
    const taskpaneState = await runtime.reviewSessionMediator.rehydrateTaskpaneState();
    setTaskpaneDisableTrackChangesCtaVisible(taskpaneState.showDisableTrackChangesCta);
  } catch (error) {
    console.warn(
      "⚠️ [Taskpane] No se pudo calcular la visibilidad del CTA de Track Changes:",
      error
    );
  }
}

/** Initializes taskpane-side subscriptions and store hydration once React is mounted. */
export function bootstrapTaskpane(runtime: TaskpaneBootstrapRuntime): void {
  void bootstrapAuthSession(runtime);
  void bootstrapAnalysisProfile(runtime);
  void refreshCleanupVisibility(runtime);
  void refreshTrackChangesCtaVisibility(runtime);
  runtime.documentPort.subscribeSelectionChanges(runtime.onSelectionSnapshot);
}

/** Loads the user's correction preferences from the backend-backed settings port. */
export async function handleLoadPreferences(
  runtime: TaskpaneBootstrapRuntime
): Promise<UserCorrectionPreferences> {
  return runtime.userCorrectionPreferencesPort.load();
}

/** Persists correction instructions plus the selected analysis profile atomically. */
export const handleSavePreferences = (
  runtime: TaskpaneBootstrapRuntime
): SaveTaskpanePreferencesHandler => {
  return async (correctionInstructions, analysisProfile) => {
    const result = await runtime.userCorrectionPreferencesPort.save(correctionInstructions);
    await runtime.userPreferencesPort.setAnalysisProfile(analysisProfile);
    runtime.setSelectedGenero(analysisProfile);
    return result;
  };
};

/** Starts the OAuth flow through the Office dialog bridge and persists the session. */
export async function handleSignIn(runtime: TaskpaneBootstrapRuntime): Promise<void> {
  setTaskpaneSigningIn(true);
  try {
    const session = await runtime.officeDialogAuthAdapter.signIn();
    await runtime.authSessionStoragePort.persist(session);
    setTaskpaneAuthenticated(session);
    showTaskpaneStatus("Sesión iniciada correctamente.", "success");
  } catch (error) {
    setTaskpaneUnauthenticated(toTaskpaneUserMessage(error));
  } finally {
    setTaskpaneSigningIn(false);
  }
}

/** Revokes the current Better Auth session and clears all local auth state. */
export async function handleSignOut(runtime: TaskpaneBootstrapRuntime): Promise<void> {
  const token = getTaskpaneAuthToken();
  setTaskpaneSigningOut(true);
  try {
    await runtime.authPort.signOut(token);
  } catch (error) {
    console.warn("⚠️ [Taskpane] No se pudo revocar sesión en backend:", error);
  } finally {
    await runtime.authSessionStoragePort.clear();
    setTaskpaneUnauthenticated();
    setTaskpaneSigningOut(false);
    showTaskpaneStatus("Sesión cerrada.", "success");
  }
}

/** Cleans up resolved Stylistic comments from the active document. */
export async function handleCleanup(runtime: TaskpaneBootstrapRuntime): Promise<void> {
  console.log("🧽 [Taskpane] Iniciando limpieza de comentarios resueltos...");
  if (getTaskpaneShellState().isCleanupLoading) {
    return;
  }

  setTaskpaneCleanupLoading(true);
  try {
    const { deleted, kept } = await runtime.documentPort.cleanupResolvedComments();
    console.log(`🧽 [Taskpane] Limpieza: ${deleted} eliminados, ${kept} conservados`);
    showTaskpaneStatus(`${deleted} comentario(s) eliminado(s), ${kept} conservado(s).`, "success");
    setTaskpaneCleanupVisible(kept > 0);
  } catch (error) {
    showTaskpaneStatus(toTaskpaneUserMessage(error), "error");
  } finally {
    setTaskpaneCleanupLoading(false);
  }
}

/** Disables Track Changes through the review-session workflow and refreshes CTA state. */
export async function handleDisableTrackChanges(runtime: TaskpaneBootstrapRuntime): Promise<void> {
  if (getTaskpaneShellState().isDisableTrackChangesLoading) {
    return;
  }

  setTaskpaneDisableTrackChangesLoading(true);
  try {
    const taskpaneState = await runtime.reviewSessionMediator.disableTrackChanges();
    setTaskpaneDisableTrackChangesCtaVisible(taskpaneState.showDisableTrackChangesCta);
    showTaskpaneStatus("Control de cambios desactivado.", "success");
  } catch (error) {
    showTaskpaneStatus(toTaskpaneUserMessage(error), "error");
  } finally {
    setTaskpaneDisableTrackChangesLoading(false);
  }
}
