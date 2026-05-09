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

  const [profileDraft, setProfileDraft] = React.useState<AnalysisProfileId>(persistedProfile);
  const [correctionInstructionsSnapshot, setCorrectionInstructionsSnapshot] = React.useState<
    string | null
  >(null);
  const [correctionInstructionsDraft, setCorrectionInstructionsDraft] = React.useState<string>("");
  const [maxLength, setMaxLength] = React.useState<number>(
    INITIAL_CORRECTION_INSTRUCTIONS_MAX_LENGTH
  );
  const [isLoadingPreferences, setIsLoadingPreferences] = React.useState<boolean>(true);
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  const [saveError, setSaveError] = React.useState<string | undefined>(undefined);

  const { loadPreferences, savePreferences } = deps;

  React.useEffect(() => {
    let cancelled = false;
    setIsLoadingPreferences(true);

    void (async () => {
      try {
        const preferences = await loadPreferences();
        if (cancelled) return;
        setCorrectionInstructionsSnapshot(preferences.correctionInstructions);
        setCorrectionInstructionsDraft(
          correctionInstructionsToDraft(preferences.correctionInstructions)
        );
        setMaxLength(preferences.correctionInstructionsMaxLength);
      } catch (error) {
        // Swallowed by design — see hook JSDoc. Real failures resurface on Save.
        console.warn("⚠️ [Settings] No se pudieron precargar las preferencias:", error);
      } finally {
        if (!cancelled) {
          setIsLoadingPreferences(false);
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [loadPreferences]);

  React.useEffect(() => {
    setProfileDraft(persistedProfile);
  }, [persistedProfile]);

  const onProfileChange = React.useCallback((value: AnalysisProfileId) => {
    setProfileDraft(value);
    setSaveError(undefined);
  }, []);

  const onCorrectionInstructionsChange = React.useCallback((value: string) => {
    setCorrectionInstructionsDraft(value);
    setSaveError(undefined);
  }, []);

  const isInstructionsDirty =
    normalizeCorrectionInstructions(correctionInstructionsDraft) !== correctionInstructionsSnapshot;
  const isProfileDirty = profileDraft !== persistedProfile;
  const isDirty = isInstructionsDirty || isProfileDirty;

  const onSave = React.useCallback(() => {
    if (isSaving) return;
    setIsSaving(true);
    setSaveError(undefined);
    void (async () => {
      try {
        const normalized = normalizeCorrectionInstructions(correctionInstructionsDraft);
        const result = await savePreferences(normalized, profileDraft);
        setCorrectionInstructionsSnapshot(result.correctionInstructions);
        setCorrectionInstructionsDraft(
          correctionInstructionsToDraft(result.correctionInstructions)
        );
        setMaxLength(result.correctionInstructionsMaxLength);
        showTaskpaneStatus(SETTINGS_SAVE_SUCCESS_MESSAGE, "success");
      } catch (error) {
        console.error("🔴 [Settings] Save failed:", error);
        const message = describeSaveError(error);
        setSaveError(message);
        showTaskpaneStatus(message, "error");
      } finally {
        setIsSaving(false);
      }
    })();
  }, [correctionInstructionsDraft, isSaving, profileDraft, savePreferences]);

  return {
    classes,
    analysisProfileOptions: ANALYSIS_PROFILE_OPTIONS,
    isFormDisabled,
    profileDraft,
    onProfileChange,
    correctionInstructionsDraft,
    correctionInstructionsMaxLength: maxLength,
    isLoadingPreferences,
    onCorrectionInstructionsChange,
    isDirty,
    isSaving,
    saveError,
    onSave,
  };
}
