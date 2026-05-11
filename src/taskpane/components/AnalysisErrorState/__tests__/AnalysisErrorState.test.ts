import * as React from "react";
import { describe, expect, it, vi } from "vitest";
import { renderToStaticMarkup } from "react-dom/server";
import { AnalysisErrorState } from "../AnalysisErrorState";

function renderAnalysisErrorState(retryKind: "full-retry" | "retry-query"): string {
  return renderToStaticMarkup(
    React.createElement(AnalysisErrorState, {
      error: {
        message: "Mensaje de error",
        retryKind,
        visible: true,
      },
      onRetryAnalysis: vi.fn(),
      onRetryAnalysisQuery: vi.fn(),
    })
  );
}

describe("AnalysisErrorState", () => {
  it("renders nothing when the error surface is hidden", () => {
    const markup = renderToStaticMarkup(
      React.createElement(AnalysisErrorState, {
        error: { message: "", retryKind: "full-retry", visible: false },
        onRetryAnalysis: vi.fn(),
        onRetryAnalysisQuery: vi.fn(),
      })
    );

    expect(markup).toBe("");
  });

  it("renders retry-query copy when the frontend can re-poll the same run", () => {
    const markup = renderAnalysisErrorState("retry-query");

    expect(markup).toContain('data-testid="analysis-error-state"');
    expect(markup).toContain('data-retry-kind="retry-query"');
    expect(markup).toContain("Reintentar consulta");
    expect(markup).toContain("sin reenviar el texto al backend");
    expect(markup).toContain('data-testid="analysis-error-retry-button"');
  });

  it("renders full-retry copy when the user must resubmit the analysis", () => {
    const markup = renderAnalysisErrorState("full-retry");

    expect(markup).toContain('data-testid="analysis-error-state"');
    expect(markup).toContain('data-retry-kind="full-retry"');
    expect(markup).toContain("Reintentar análisis");
    expect(markup).toContain("volver a enviar la petición al backend");
    expect(markup).toContain('data-testid="analysis-error-retry-button"');
  });

  it("renders children inside the actions area to colocate selection chips with the retry CTA", () => {
    const markup = renderToStaticMarkup(
      React.createElement(
        AnalysisErrorState,
        {
          error: { message: "Mensaje de error", retryKind: "full-retry", visible: true },
          onRetryAnalysis: vi.fn(),
          onRetryAnalysisQuery: vi.fn(),
        },
        React.createElement("span", { "data-testid": "selection-slot" }, "Selección activa")
      )
    );

    expect(markup).toContain('data-testid="selection-slot"');
    expect(markup).toContain("Selección activa");
  });
});
