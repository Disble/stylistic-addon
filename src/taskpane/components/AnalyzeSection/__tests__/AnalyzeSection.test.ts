import * as React from "react";
import { describe, expect, it, vi } from "vitest";
import { renderToStaticMarkup } from "react-dom/server";
import { AnalyzeSection } from "../AnalyzeSection";

describe("AnalyzeSection", () => {
  it("shows only the analyze action, with no privacy notice under it", () => {
    const markup = renderToStaticMarkup(
      React.createElement(AnalyzeSection, { isLoading: false, onAnalyze: vi.fn() })
    );

    expect(markup).toContain("Analizar y sugerir");
    expect(markup).not.toContain('data-testid="analyze-privacy-notice"');
    expect(markup).not.toContain("No se envía nada más a ningún lugar");
  });
});
