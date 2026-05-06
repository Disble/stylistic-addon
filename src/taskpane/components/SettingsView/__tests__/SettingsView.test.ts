import * as React from "react";
import { describe, expect, it, vi } from "vitest";
import type { AnalysisProfileSectionProps } from "../../AnalysisProfileSection";
import { SettingsView } from "../SettingsView";
import { useSettingsView } from "../useSettingsView";

vi.mock("../../AccountSettings", () => ({
  AccountSettings: (props: unknown) => React.createElement("account-settings", props),
}));

vi.mock("../../AnalysisProfileSection", () => ({
  AnalysisProfileSection: (props: AnalysisProfileSectionProps) =>
    React.createElement("analysis-profile-section", props),
}));

vi.mock("../useSettingsView", () => ({
  useSettingsView: vi.fn(),
}));

function getAnalysisProfileElement(element: React.ReactElement): React.ReactElement {
  const body = React.Children.toArray(element.props.children)[1] as React.ReactElement<{
    children?: React.ReactNode;
  }>;
  const bodyChildren = React.Children.toArray(body.props.children);
  return bodyChildren[1] as React.ReactElement;
}

describe("SettingsView", () => {
  it("passes the loading-derived disabled state to the analysis profile selector", () => {
    vi.mocked(useSettingsView).mockReturnValue({
      classes: {
        root: "root",
        header: "header",
        backButton: "backButton",
        title: "title",
        body: "body",
      },
      analysisProfileOptions: [{ value: "general", label: "General" }],
      isAnalysisProfileDisabled: true,
      selectedGenero: "general",
      handleGeneroChange: vi.fn(),
    });

    const view = SettingsView({
      isSigningOut: false,
      onBack: vi.fn(),
      onSignOut: vi.fn(),
      session: undefined,
    });

    const analysisProfile = getAnalysisProfileElement(view);
    expect(analysisProfile.props.isDisabled).toBe(true);
    expect(analysisProfile.props.selectedGenero).toBe("general");
  });
});
