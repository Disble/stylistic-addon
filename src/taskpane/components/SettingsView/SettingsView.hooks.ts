/* global console */

import * as React from "react";
import type { AnalysisProfileId } from "../../../domain/Profile.types";
import { showTaskpaneStatus, useTaskpaneShellStore } from "../../TaskpaneShellStore";
import {
  ANALYSIS_PROFILE_OPTIONS,
  INITIAL_CORRECTION_INSTRUCTIONS_MAX_LENGTH,
  SETTINGS_SAVE_SUCCESS_MESSAGE,
} from "./SettingsView.constants";
import {
  correctionInstructionsToDraft,
  describeSaveError,
  normalizeCorrectionInstructions,
} from "./SettingsView.helpers";
import { useSettingsViewStyles } from "./SettingsView.styles";
import type {
  SettingsViewClasses,
  SettingsViewHookDeps,
  SettingsViewState,
} from "./SettingsView.types";

/** Returns Griffel classes for the settings page. */
function useSettingsViewClasses(): SettingsViewClasses {
  const styles = useSettingsViewStyles();
  return {
    root: styles.root,
    header: styles.header,
    backButton: styles.backButton,
    title: styles.title,
    body: styles.body,
  };
}

/**
 * Aggregates the settings page state behind a draft-and-save model:
 *
 * - On mount, calls `loadPreferences` to hydrate the textarea draft and the
 *   max-length contract echoed by the backend. The persisted analysis-profile
 *   value comes from the shell store.
 * - Tracks both fields as a local draft. The shell store is NOT updated as the
 *   user types — it is only committed after a successful save.
 * - Load failures are intentionally swallowed (no inline error UI). The user
 *   would see a scary red message at a passive moment with no action to take —
 *   bad UX. If the backend really is unreachable, the Save click will surface
 *   a typed error at the moment the user actually attempts an action.
 *   Failures are logged to console for developer visibility only.
 * - `onSave` wraps the injected `savePreferences` callback, surfaces inline
 *   errors with a typed message, and emits a global status toast on success.
 * - The whole form is disabled while an analysis run is active so the persisted
 *   preference cannot diverge from the pipeline snapshot already in flight.
 */
export function useSettingsView(deps: SettingsViewHookDeps): SettingsViewState {
  const classes = useSettingsViewClasses();
  const persistedProfile = useTaskpaneShellStore((state) => state.selectedGenero);
  const isFormDisabled = useTaskpaneShellStore((state) => state.isAnalyzeLoading);

  const [state, setState] = React.useState<{
    profileDraftOverride: AnalysisProfileId | null;
    correctionInstructionsSnapshot: string | null;
    correctionInstructionsDraft: string;
    maxLength: number;
    isLoadingPreferences: boolean;
    isSaving: boolean;
    saveError: string | undefined;
  }>({
    profileDraftOverride: null,
    correctionInstructionsSnapshot: null,
    correctionInstructionsDraft: "",
    maxLength: INITIAL_CORRECTION_INSTRUCTIONS_MAX_LENGTH,
    isLoadingPreferences: true,
    isSaving: false,
    saveError: undefined,
  });

  const { loadPreferences, savePreferences } = deps;
  const profileDraft = state.profileDraftOverride ?? persistedProfile;

  React.useEffect(() => {
    let cancelled = false;
    setState((current) => ({ ...current, isLoadingPreferences: true }));

    void (async () => {
      let nextSnapshot: string | null | undefined;
      let nextDraft: string | undefined;
      let nextMaxLength: number | undefined;

      try {
        const preferences = await loadPreferences();
        nextSnapshot = preferences.correctionInstructions;
        nextDraft = correctionInstructionsToDraft(preferences.correctionInstructions);
        nextMaxLength = preferences.correctionInstructionsMaxLength;
      } catch (error) {
        // Swallowed by design — see hook JSDoc. Real failures resurface on Save.
        console.warn("⚠️ [Settings] No se pudieron precargar las preferencias:", error);
      } finally {
        if (!cancelled) {
          setState((current) => ({
            ...current,
            correctionInstructionsSnapshot:
              nextSnapshot === undefined ? current.correctionInstructionsSnapshot : nextSnapshot,
            correctionInstructionsDraft:
              nextDraft === undefined ? current.correctionInstructionsDraft : nextDraft,
            maxLength: nextMaxLength === undefined ? current.maxLength : nextMaxLength,
            isLoadingPreferences: false,
          }));
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [loadPreferences]);

  const onProfileChange = React.useCallback((value: AnalysisProfileId) => {
    setState((current) => ({ ...current, profileDraftOverride: value, saveError: undefined }));
  }, []);

  const onCorrectionInstructionsChange = React.useCallback((value: string) => {
    setState((current) => ({
      ...current,
      correctionInstructionsDraft: value,
      saveError: undefined,
    }));
  }, []);

  const isInstructionsDirty =
    normalizeCorrectionInstructions(state.correctionInstructionsDraft) !==
    state.correctionInstructionsSnapshot;
  const isProfileDirty = profileDraft !== persistedProfile;
  const isDirty = isInstructionsDirty || isProfileDirty;

  const onSave = React.useCallback(() => {
    if (state.isSaving) return;
    setState((current) => ({ ...current, isSaving: true, saveError: undefined }));

    void (async () => {
      try {
        const normalized = normalizeCorrectionInstructions(state.correctionInstructionsDraft);
        const result = await savePreferences(normalized, profileDraft);
        setState((current) => ({
          ...current,
          profileDraftOverride: null,
          correctionInstructionsSnapshot: result.correctionInstructions,
          correctionInstructionsDraft: correctionInstructionsToDraft(result.correctionInstructions),
          maxLength: result.correctionInstructionsMaxLength,
        }));
        showTaskpaneStatus(SETTINGS_SAVE_SUCCESS_MESSAGE, "success");
      } catch (error) {
        console.error("🔴 [Settings] Save failed:", error);
        const message = describeSaveError(error);
        setState((current) => ({ ...current, saveError: message }));
        showTaskpaneStatus(message, "error");
      } finally {
        setState((current) => ({ ...current, isSaving: false }));
      }
    })();
  }, [profileDraft, savePreferences, state.correctionInstructionsDraft, state.isSaving]);

  return {
    classes,
    analysisProfileOptions: ANALYSIS_PROFILE_OPTIONS,
    isFormDisabled,
    profileDraft,
    onProfileChange,
    correctionInstructionsDraft: state.correctionInstructionsDraft,
    correctionInstructionsMaxLength: state.maxLength,
    isLoadingPreferences: state.isLoadingPreferences,
    onCorrectionInstructionsChange,
    isDirty,
    isSaving: state.isSaving,
    saveError: state.saveError,
    onSave,
  };
}
