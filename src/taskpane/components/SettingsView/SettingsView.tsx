import * as React from "react";
import { Button } from "@fluentui/react-components";
import { ArrowLeftRegular } from "@fluentui/react-icons";
import { AccountSettings } from "../AccountSettings";
import { AnalysisProfileSection } from "../AnalysisProfileSection";
import type { SettingsViewProps } from "./SettingsView.types";
import { useSettingsView } from "./useSettingsView";

/**
 * Renders the secondary settings page accessed from the main toolbar gear.
 * Hosts user-level configuration that does not belong in the per-run flow:
 * account info, analysis-profile selection, etc.
 */
export function SettingsView({
  isSigningOut,
  onBack,
  onSignOut,
  session,
}: SettingsViewProps): React.JSX.Element {
  const {
    classes,
    analysisProfileOptions,
    isAnalysisProfileDisabled,
    selectedGenero,
    handleGeneroChange,
  } = useSettingsView();

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
          isDisabled={isAnalysisProfileDisabled}
          onGeneroChange={handleGeneroChange}
          options={analysisProfileOptions}
          selectedGenero={selectedGenero}
        />
      </div>
    </section>
  );
}
