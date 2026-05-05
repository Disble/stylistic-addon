import * as React from "react";
import { vi } from "vitest";

/**
 * Centralized Fluent UI v9 mocks for taskpane tests.
 *
 * Tests in this folder use `renderToStaticMarkup` to inspect the HTML output of
 * React components. Real Fluent UI loads heavy modules and injects styles via
 * Griffel which makes the entrypoint test suite slow and noisy. These mocks
 * replace each Fluent component with a minimal DOM equivalent that preserves
 * the props relevant for assertions (aria-*, data-*, className, ...).
 *
 * Side-effect import: `import "./TaskpaneFluentMocks";` registers all mocks.
 */

vi.mock("@fluentui/react-components", () => {
  const Passthrough = function Passthrough({ children }: { children?: unknown }) {
    return children as unknown;
  };
  const renderAs = (tag: string, displayName: string) => {
    const Component = function FluentMockComponent({ children, icon, ...rest }: any) {
      const props = { ...rest } as Record<string, unknown>;
      const childNodes = icon ? [icon, children] : [children];
      return React.createElement(tag, props, ...childNodes);
    };
    Component.displayName = displayName;
    return Component;
  };
  const TextareaMock = function TextareaMock({
    value,
    onChange,
    children: _children,
    ...rest
  }: any) {
    return React.createElement("textarea", {
      ...rest,
      value,
      onChange: (event: any) => {
        if (typeof onChange === "function") {
          onChange(event, { value: event?.target?.value ?? "" });
        }
      },
    });
  };
  const DropdownMock = function DropdownMock({
    children,
    onOptionSelect,
    selectedOptions,
    value,
    placeholder: _placeholder,
    ...rest
  }: any) {
    return React.createElement(
      "div",
      {
        ...rest,
        role: "combobox",
        "data-selected-value": Array.isArray(selectedOptions) ? selectedOptions[0] : "",
        "data-display-value": value,
        onChange: (event: any) => {
          if (typeof onOptionSelect === "function") {
            onOptionSelect(event, {
              optionValue: event?.target?.value ?? "",
              optionText: event?.target?.value ?? "",
              selectedOptions: [event?.target?.value ?? ""],
            });
          }
        },
      },
      children
    );
  };
  const OptionMock = function OptionMock({ children, value, ...rest }: any) {
    return React.createElement(
      "span",
      { ...rest, role: "option", "data-option-value": value },
      children
    );
  };
  const FieldMock = function FieldMock({ children, label, validationState, ...rest }: any) {
    const labelNode = label
      ? React.createElement("span", { "data-field-label": "true" }, label)
      : null;
    return React.createElement(
      "div",
      { ...rest, "data-field": "true", "data-validation": validationState ?? "" },
      labelNode,
      children
    );
  };
  const ProgressBarMock = function ProgressBarMock({ value, max, ...rest }: any) {
    const numericMax = typeof max === "number" && max > 0 ? max : 1;
    const numericValue = typeof value === "number" ? value : 0;
    const percent = Math.round((numericValue / numericMax) * 100);
    return React.createElement("div", {
      ...rest,
      role: "progressbar",
      "aria-valuenow": numericValue,
      "aria-valuemax": numericMax,
      "data-progress-percent": percent,
    });
  };
  const SpinnerMock = function SpinnerMock({ size: _size, label, ...rest }: any) {
    return React.createElement(
      "span",
      { ...rest, role: "status", "data-spinner": "true" },
      label ?? null
    );
  };
  return {
    FluentProvider: Passthrough,
    MessageBar: renderAs("div", "MessageBar"),
    MessageBarBody: renderAs("div", "MessageBarBody"),
    MessageBarTitle: renderAs("strong", "MessageBarTitle"),
    Card: renderAs("section", "Card"),
    CardHeader: renderAs("header", "CardHeader"),
    CardFooter: renderAs("footer", "CardFooter"),
    Badge: renderAs("span", "Badge"),
    Button: renderAs("button", "Button"),
    Textarea: TextareaMock,
    Dropdown: DropdownMock,
    Option: OptionMock,
    Field: FieldMock,
    ProgressBar: ProgressBarMock,
    Spinner: SpinnerMock,
    Body1: renderAs("span", "Body1"),
    Body1Strong: renderAs("strong", "Body1Strong"),
    Body2: renderAs("p", "Body2"),
    Caption1: renderAs("span", "Caption1"),
    Caption1Strong: renderAs("span", "Caption1Strong"),
    Subtitle2: renderAs("h3", "Subtitle2"),
    Title2: renderAs("h2", "Title2"),
    webLightTheme: {},
    makeStyles: () => () =>
      new Proxy(
        {},
        {
          get: (_target, key) => (typeof key === "string" ? key : ""),
        }
      ),
    mergeClasses: (...classes: unknown[]) => classes.filter(Boolean).join(" "),
    tokens: new Proxy(
      {},
      {
        get: () => "",
      }
    ),
  };
});

vi.mock("@fluentui/react-icons", () => ({
  TextEditStyleRegular: () => null,
  CheckmarkRegular: () => null,
  DismissRegular: () => null,
  CommentRegular: () => null,
  CommentMultipleRegular: () => null,
  WandRegular: () => null,
  BroomRegular: () => null,
  EraserRegular: () => null,
  DocumentEditRegular: () => null,
  ArrowSyncRegular: () => null,
  ChevronDownRegular: () => null,
  HistoryDismissRegular: () => null,
}));
