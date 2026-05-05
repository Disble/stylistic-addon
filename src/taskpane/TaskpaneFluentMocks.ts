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
  return {
    FluentProvider: Passthrough,
    MessageBar: Passthrough,
    MessageBarBody: Passthrough,
    MessageBarTitle: Passthrough,
    Card: renderAs("section", "Card"),
    CardHeader: renderAs("header", "CardHeader"),
    CardFooter: renderAs("footer", "CardFooter"),
    Badge: renderAs("span", "Badge"),
    Button: renderAs("button", "Button"),
    Textarea: TextareaMock,
    Body1: renderAs("span", "Body1"),
    Body1Strong: renderAs("strong", "Body1Strong"),
    Body2: renderAs("p", "Body2"),
    Caption1: renderAs("span", "Caption1"),
    Caption1Strong: renderAs("span", "Caption1Strong"),
    Subtitle2: renderAs("h3", "Subtitle2"),
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
  ArrowRightRegular: () => null,
}));
