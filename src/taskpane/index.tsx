import * as React from "react";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { createRoot } from "react-dom/client";
import { bootstrapTaskpane } from "./taskpane";
import { App } from "./components/App";

/**
 * React bootstrap entrypoint for the Stylistic taskpane.
 * Renders the React shell first, then binds the legacy composition root to the rendered DOM.
 */
function renderTaskpaneShell(): void {
  const office = globalThis.Office;
  const rootElement = document.getElementById("container");

  if (!office?.onReady || !rootElement) {
    return;
  }

  const root = createRoot(rootElement);

  office.onReady((info) => {
    const wordHost = office.HostType?.Word ?? "Word";
    if (info.host !== wordHost) {
      return;
    }

    root.render(
      <React.StrictMode>
        <FluentProvider theme={webLightTheme}>
          <App />
        </FluentProvider>
      </React.StrictMode>
    );

    queueMicrotask(() => {
      bootstrapTaskpane(document, office);
    });
  });
}

renderTaskpaneShell();
