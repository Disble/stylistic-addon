import * as React from "react";
import { describe, expect, it, vi } from "vitest";
import { renderToStaticMarkup } from "react-dom/server";
import { AnalyzeSection } from "../AnalyzeSection";

describe("AnalyzeSection", () => {
  it("tells the user where the analyzed text goes before they analyze", () => {
    const markup = renderToStaticMarkup(
      React.createElement(AnalyzeSection, { isLoading: false, onAnalyze: vi.fn() })
    );

    expect(markup).toContain('data-testid="analyze-privacy-notice"');
    expect(markup).toContain(
      "El texto que analizas y el perfil de estilo del documento solo se envían al proveedor que elegiste en Stylistic Hub. No se envía nada más a ningún lugar."
    );
    expect(markup).toContain("Analizar y sugerir");
  });
});
