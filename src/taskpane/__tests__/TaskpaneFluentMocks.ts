import * as React from "react";
import { vi } from "vitest";
import type { MockProps } from "./TaskpaneFluentMocks.types";

/** Returns only string class names so mocked mergeClasses never stringifies objects. */
function isNonEmptyClassName(value: unknown): value is string {
  return typeof value === "string" && value.length > 0;
}

/** Minimal passthrough mock used for Fluent providers that only render children. */
function Passthrough({ children }: MockProps): React.ReactNode {
  return children;
}

/** Creates a simple DOM-backed mock for one Fluent component. */
function renderAs(tag: string, displayName: string) {
  const Component = function FluentMockComponent({ children, icon, ...rest }: MockProps) {
    const props = { ...rest };
    const childNodes = icon ? [icon, children] : [children];
    return React.createElement(tag, props, ...childNodes);
  };
  Component.displayName = displayName;
  return Component;
}

/** Mocks Fluent Textarea while preserving the two-argument onChange contract. */
function TextareaMock({
  value,
  onChange,
  children: _children,
  ...rest
}: MockProps): React.JSX.Element {
  return React.createElement("textarea", {
    ...rest,
    value,
    onChange: (event: { target?: { value?: string } }) => {
      if (typeof onChange === "function") {
        onChange(event, { value: event?.target?.value ?? "" });
      }
    },
  });
}

/** Mocks Fluent Dropdown while preserving option-select payloads. */
function DropdownMock({
  children,
  onOptionSelect,
  selectedOptions,
  value,
  placeholder: _placeholder,
  ...rest
}: MockProps): React.JSX.Element {
  return React.createElement(
    "div",
    {
      ...rest,
      role: "combobox",
      "data-selected-value": Array.isArray(selectedOptions) ? selectedOptions[0] : "",
      "data-display-value": value,
      onChange: (event: { target?: { value?: string } }) => {
        if (typeof onOptionSelect === "function") {
          const optionValue = event?.target?.value ?? "";
          onOptionSelect(event, {
            optionValue,
            optionText: optionValue,
            selectedOptions: [optionValue],
          });
        }
      },
    },
    children
  );
}

/** Mocks Fluent Option as a plain span annotated with its option value. */
function OptionMock({ children, value, ...rest }: MockProps): React.JSX.Element {
  return React.createElement(
    "span",
    { ...rest, role: "option", "data-option-value": value },
    children
  );
}

/** Mocks Fluent Field while surfacing label and validation state metadata. */
function FieldMock({ children, label, validationState, ...rest }: MockProps): React.JSX.Element {
  const labelNode = label
    ? React.createElement("span", { "data-field-label": "true" }, label)
    : null;
  return React.createElement(
    "div",
    { ...rest, "data-field": "true", "data-validation": validationState ?? "" },
    labelNode,
    children
  );
}

/** Mocks Fluent ProgressBar with a DOM-only progress representation. */
function ProgressBarMock({ value, max, ...rest }: MockProps): React.JSX.Element {
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
}

/** Mocks Fluent Spinner as a status node with optional inline label. */
function SpinnerMock({ size: _size, label, ...rest }: MockProps): React.JSX.Element {
  return React.createElement(
    "span",
    { ...rest, role: "status", "data-spinner": "true" },
    label ?? null
  );
}

/** Returns a stable class string without Object stringification noise. */
function mergeMockClasses(...classes: unknown[]): string {
  return classes.filter(isNonEmptyClassName).join(" ");
}

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
    Skeleton: renderAs("div", "Skeleton"),
    SkeletonItem: renderAs("span", "SkeletonItem"),
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
    mergeClasses: mergeMockClasses,
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
