import * as React from "react";
import { describe, expect, it } from "vitest";
import { renderToStaticMarkup } from "react-dom/server";
import { HeroEmptyState } from "../HeroEmptyState";

describe("HeroEmptyState", () => {
  it("invites the user to start in neutral Spanish", () => {
    const markup = renderToStaticMarkup(React.createElement(HeroEmptyState, null));

    expect(markup).toContain(
      "Analiza el documento completo o selecciona un fragmento para empezar."
    );
    expect(markup).not.toContain("Analizá");
    expect(markup).not.toContain("seleccioná");
  });
});
