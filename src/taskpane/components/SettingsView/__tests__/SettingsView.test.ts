import * as React from "react";
import { describe, expect, it, vi } from "vitest";
import type { AnalysisProfileSectionProps } from "../../AnalysisProfileSection";
import type { CorrectionInstructionsSectionProps } from "../../CorrectionInstructionsSection";
import type { SettingsSaveBarProps } from "../../SettingsSaveBar";
import { SettingsView } from "../SettingsView";
import { useSettingsView } from "../SettingsView.hooks";
import type { SettingsViewState } from "../SettingsView.types";

vi.mock("../../AccountSettings", () => ({
  AccountSettings: (props: unknown) => React.createElement("account-settings", props),
}));

vi.mock("../../AnalysisProfileSection", () => ({
  AnalysisProfileSection: (props: AnalysisProfileSectionProps) =>
    React.createElement("analysis-profile-section", props),
}));

vi.mock("../../CorrectionInstructionsSection", () => ({
  CorrectionInstructionsSection: (props: CorrectionInstructionsSectionProps) =>
    React.createElement("correction-instructions-section", props),
}));

vi.mock("../../SettingsSaveBar", () => ({
  SettingsSaveBar: (props: SettingsSaveBarProps) => React.createElement("settings-save-bar", props),
}));

vi.mock("../SettingsView.hooks", () => ({
  useSettingsView: vi.fn(),
}));

type BodyChildren = readonly [
  React.ReactElement,
  React.ReactElement<AnalysisProfileSectionProps>,
  React.ReactElement<CorrectionInstructionsSectionProps>,
  React.ReactElement<SettingsSaveBarProps>,
];

function getBodyChildren(element: React.ReactElement): BodyChildren {
  const body = React.Children.toArray(element.props.children)[1] as React.ReactElement<{
    children?: React.ReactNode;
  }>;
  const children = React.Children.toArray(body.props.children);
  if (children.length !== 4) {
    throw new Error(`Expected 4 settings body children, received ${children.length}.`);
  }

  return [
    children[0] as React.ReactElement,
    children[1] as React.ReactElement<AnalysisProfileSectionProps>,
    children[2] as React.ReactElement<CorrectionInstructionsSectionProps>,
    children[3] as React.ReactElement<SettingsSaveBarProps>,
  ];
}

function makeState(overrides: Partial<SettingsViewState> = {}): SettingsViewState {
  return {
    classes: {
      root: "root",
      header: "header",
      backButton: "backButton",
      title: "title",
      body: "body",
    },
    analysisProfileOptions: [{ value: "general", label: "General" }],
    isFormDisabled: false,
    profileDraft: "general",
    onProfileChange: vi.fn(),
    correctionInstructionsDraft: "",
    correctionInstructionsMaxLength: 4000,
    isLoadingPreferences: false,
    onCorrectionInstructionsChange: vi.fn(),
    isDirty: false,
    isSaving: false,
    saveError: undefined,
    onSave: vi.fn(),
    ...overrides,
  };
}

const NOOP_PROPS = {
  isSigningOut: false,
  onBack: vi.fn(),
  onSignOut: vi.fn(),
  session: undefined,
  loadPreferences: vi.fn(),
  savePreferences: vi.fn(),
};

describe("SettingsView", () => {
  it("propagates the analysis-profile draft and disabled flag to the selector", () => {
    vi.mocked(useSettingsView).mockReturnValue(
      makeState({ isFormDisabled: true, profileDraft: "general" })
    );

    const view = SettingsView(NOOP_PROPS);
    const [, profile] = getBodyChildren(view);

    expect(profile.props.selectedGenero).toBe("general");
    expect(profile.props.isDisabled).toBe(true);
  });

  it("propagates the correction-instructions draft and max length to the section", () => {
    vi.mocked(useSettingsView).mockReturnValue(
      makeState({
        correctionInstructionsDraft: "Vigilá X.",
        correctionInstructionsMaxLength: 4000,
        isLoadingPreferences: true,
      })
    );

    const view = SettingsView(NOOP_PROPS);
    const [, , instructions] = getBodyChildren(view);

    expect(instructions.props.value).toBe("Vigilá X.");
    expect(instructions.props.maxLength).toBe(4000);
    expect(instructions.props.isLoading).toBe(true);
    expect(instructions.props.isDisabled).toBe(true);
  });

  it("disables the save bar when there are no pending changes", () => {
    vi.mocked(useSettingsView).mockReturnValue(makeState({ isDirty: false }));

    const view = SettingsView(NOOP_PROPS);
    const [, , , saveBar] = getBodyChildren(view);

    expect(saveBar.props.isDirty).toBe(false);
  });

  it("treats the form as not-dirty while disabled by an active analysis run", () => {
    vi.mocked(useSettingsView).mockReturnValue(makeState({ isDirty: true, isFormDisabled: true }));

    const view = SettingsView(NOOP_PROPS);
    const [, , , saveBar] = getBodyChildren(view);

    expect(saveBar.props.isDirty).toBe(false);
  });

  it("forwards the save error to the save bar", () => {
    vi.mocked(useSettingsView).mockReturnValue(
      makeState({ isDirty: true, saveError: "No se pudo guardar." })
    );

    const view = SettingsView(NOOP_PROPS);
    const [, , , saveBar] = getBodyChildren(view);

    expect(saveBar.props.saveError).toBe("No se pudo guardar.");
  });
});
