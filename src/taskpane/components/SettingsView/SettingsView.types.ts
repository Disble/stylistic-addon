import type { AuthSession } from "../../../domain/auth/AuthSession.types";
import type { AnalysisProfileId } from "../../../domain/Profile.types";
import type { AnalysisProfileOption } from "../AnalysisProfileSection";

/** Props required to render the full settings view. */
export type SettingsViewProps = Readonly<{
  isSigningOut: boolean;
  onBack: () => void;
  onSignOut: () => Promise<void> | void;
  session?: AuthSession;
}>;

/** Griffel class slots consumed by the settings view. */
export type SettingsViewClasses = Readonly<{
  root: string;
  header: string;
  backButton: string;
  title: string;
  body: string;
}>;

/** View model consumed by the settings-page component. */
export type SettingsViewState = Readonly<{
  classes: SettingsViewClasses;
  analysisProfileOptions: readonly AnalysisProfileOption[];
  isAnalysisProfileDisabled: boolean;
  selectedGenero: AnalysisProfileId;
  handleGeneroChange: (value: AnalysisProfileId) => void;
}>;
