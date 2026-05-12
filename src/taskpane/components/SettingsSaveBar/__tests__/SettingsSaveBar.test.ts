import * as React from "react";
import { describe, expect, it, vi } from "vitest";

vi.mock("@fluentui/react-components", () => ({
  Button: (props: Record<string, unknown>) => React.createElement("button", props),
  Spinner: (props: Record<string, unknown>) => React.createElement("spinner", props),
}));

vi.mock("../SettingsSaveBar.hooks", () => ({
  useSettingsSaveBar: () => ({
    root: "root",
    error: "error",
    actions: "actions",
    saveButton: "saveButton",
  }),
}));

import { SettingsSaveBar } from "../SettingsSaveBar";

function findSaveButton(element: React.ReactElement): React.ReactElement {
  const children = React.Children.toArray(element.props.children);

  for (const child of children) {
    if (!React.isValidElement<{ children?: React.ReactNode; className?: string }>(child)) {
      continue;
    }

    if (child.props.className !== "actions") {
      continue;
    }

    const buttonChildren = React.Children.toArray(child.props.children);
    const firstButton = buttonChildren[0];
    if (firstButton && React.isValidElement(firstButton)) {
      return firstButton;
    }
  }

  throw new Error("SettingsSaveBar action row did not render a save button.");
}

function findError(element: React.ReactElement): React.ReactElement | undefined {
  const children = React.Children.toArray(element.props.children);

  for (const child of children) {
    if (
      React.isValidElement<{ "data-testid"?: string }>(child) &&
      child.props["data-testid"] === "settings-save-error"
    ) {
      return child;
    }
  }

  return undefined;
}

describe("SettingsSaveBar", () => {
  it("disables the save button when there are no pending changes", () => {
    const view = SettingsSaveBar({
      isDirty: false,
      isSaving: false,
      onSave: vi.fn(),
    });

    const button = findSaveButton(view);
    expect(button.props.disabled).toBe(true);
  });

  it("enables the save button when the form is dirty and not saving", () => {
    const view = SettingsSaveBar({
      isDirty: true,
      isSaving: false,
      onSave: vi.fn(),
    });

    const button = findSaveButton(view);
    expect(button.props.disabled).toBe(false);
  });

  it("disables the save button while a save is in flight", () => {
    const view = SettingsSaveBar({
      isDirty: true,
      isSaving: true,
      onSave: vi.fn(),
    });

    const button = findSaveButton(view);
    expect(button.props.disabled).toBe(true);
    expect(button.props.children).toBe("Guardando...");
  });

  it("renders the save error inline when provided", () => {
    const view = SettingsSaveBar({
      isDirty: true,
      isSaving: false,
      saveError: "No se pudo guardar.",
      onSave: vi.fn(),
    });

    const error = findError(view);
    expect(error).toBeDefined();
    expect(error?.props.children).toBe("No se pudo guardar.");
  });

  it("invokes onSave when the button is clicked", () => {
    const onSave = vi.fn();
    const view = SettingsSaveBar({
      isDirty: true,
      isSaving: false,
      onSave,
    });

    const button = findSaveButton(view);
    (button.props as { onClick: () => void }).onClick();

    expect(onSave).toHaveBeenCalledTimes(1);
  });
});
