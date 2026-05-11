import type { AuthSession } from "../../../domain/auth/AuthSession.types";
import type { AnalysisProfileId } from "../../../domain/Profile.types";
import type { UserCorrectionPreferences } from "../../../domain/user-preferences/UserCorrectionPreferences.types";
import type { AnalysisProfileOption } from "../AnalysisProfileSection";

/** Loads the user's persisted correction preferences from the backend. */
export type SettingsLoadPreferences = () => Promise<UserCorrectionPreferences>;

/**
 * Persists the full settings form atomically:
 * - PUT the correction instructions to the backend
 * - persist the analysis profile to OfficeRuntime
 * - commit the new profile into the shell store on success
 *
 * Throws on any failure; partial commits are not allowed.
 */
export type SettingsSavePreferences = (
  correctionInstructions: string | null,
  analysisProfile: AnalysisProfileId
) => Promise<UserCorrectionPreferences>;

/** Props required to render the full settings view. */
export type SettingsViewProps = Readonly<{
  isSigningOut: boolean;
  onBack: () => void;
  onSignOut: () => Promise<void> | void;
  session?: AuthSession;
  loadPreferences: SettingsLoadPreferences;
  savePreferences: SettingsSavePreferences;
}>;

/** Griffel class slots consumed by the settings view. */
export type SettingsViewClasses = Readonly<{
  root: string;
  header: string;
  backButton: string;
  title: string;
  body: string;
}>;

/** Constructor dependencies passed to the settings-view hook. */
export type SettingsViewHookDeps = Readonly<{
  loadPreferences: SettingsLoadPreferences;
  savePreferences: SettingsSavePreferences;
}>;

/** View model consumed by the settings-page component. */
export type SettingsViewState = Readonly<{
  classes: SettingsViewClasses;
  analysisProfileOptions: readonly AnalysisProfileOption[];
  isFormDisabled: boolean;
  profileDraft: AnalysisProfileId;
  onProfileChange: (value: AnalysisProfileId) => void;
  correctionInstructionsDraft: string;
  correctionInstructionsMaxLength: number;
  isLoadingPreferences: boolean;
  onCorrectionInstructionsChange: (value: string) => void;
  isDirty: boolean;
  isSaving: boolean;
  saveError?: string;
  onSave: () => void;
}>;
