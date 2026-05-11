import * as React from "react";
import { Button } from "@fluentui/react-components";
import { ArrowLeftRegular } from "@fluentui/react-icons";
import { AccountSettings } from "../AccountSettings";
import { AnalysisProfileSection } from "../AnalysisProfileSection";
import { CorrectionInstructionsSection } from "../CorrectionInstructionsSection";
import { SettingsSaveBar } from "../SettingsSaveBar";
import type { SettingsViewProps } from "./SettingsView.types";
import { useSettingsView } from "./SettingsView.hooks";

/**
 * Renders the secondary settings page accessed from the main toolbar gear.
 *
 * Hosts user-level configuration that does not belong in the per-run flow
 * (account info, analysis-profile selection, global correction instructions).
 * Uses an explicit draft-and-save model: nothing is persisted until the user
 * clicks the bottom Save bar. While an analysis run is active, the form is
 * disabled so the persisted preferences cannot diverge from the pipeline
 * snapshot already in flight.
 */
export function SettingsView({
  isSigningOut,
  loadPreferences,
  onBack,
  onSignOut,
  savePreferences,
  session,
}: SettingsViewProps): React.JSX.Element {
  const {
    classes,
    analysisProfileOptions,
    isFormDisabled,
    profileDraft,
    onProfileChange,
    correctionInstructionsDraft,
    correctionInstructionsMaxLength,
    isLoadingPreferences,
    onCorrectionInstructionsChange,
    isDirty,
    isSaving,
    saveError,
    onSave,
  } = useSettingsView({ loadPreferences, savePreferences });

  return (
    <section className={classes.root} aria-label="Settings">
      <header className={classes.header}>
        <Button
          appearance="subtle"
          aria-label="Volver"
          className={classes.backButton}
          data-testid="settings-back-button"
          icon={<ArrowLeftRegular aria-hidden="true" />}
          onClick={onBack}
          type="button"
        />
        <span className={classes.title}>Settings</span>
        <span aria-hidden="true" />
      </header>
      <div className={classes.body}>
        <AccountSettings
          isSigningOut={isSigningOut}
          onSignOut={onSignOut}
          session={session}
        />
        <AnalysisProfileSection
          isDisabled={isFormDisabled || isSaving}
          onGeneroChange={onProfileChange}
          options={analysisProfileOptions}
          selectedGenero={profileDraft}
        />
        <CorrectionInstructionsSection
          isDisabled={isFormDisabled || isSaving || isLoadingPreferences}
          isLoading={isLoadingPreferences}
          maxLength={correctionInstructionsMaxLength}
          onChange={onCorrectionInstructionsChange}
          value={correctionInstructionsDraft}
        />
        <SettingsSaveBar
          isDirty={isDirty && !isFormDisabled}
          isSaving={isSaving}
          onSave={onSave}
          saveError={saveError}
        />
      </div>
    </section>
  );
}
