import * as React from "react";
import { Button, Spinner } from "@fluentui/react-components";
import { WandRegular } from "@fluentui/react-icons";
import type { AnalyzeSectionProps } from "./AnalyzeSection.types";
import { useAnalyzeSection } from "./useAnalyzeSection";

/** Renders the primary "analyze" CTA with a wand icon and inline spinner. */
export function AnalyzeSection({ isLoading, onAnalyze }: AnalyzeSectionProps): React.JSX.Element {
  const classes = useAnalyzeSection();
  const icon = isLoading ? (
    <Spinner size="tiny" data-testid="analyze-spinner" />
  ) : (
    <WandRegular aria-hidden="true" />
  );

  return (
    <div className={classes.root}>
      <Button
        appearance="primary"
        aria-label="Analizar y sugerir"
        className={classes.button}
        data-testid="analyze-button"
        disabled={isLoading}
        icon={icon}
        onClick={() => {
          void onAnalyze();
        }}
        type="button"
      >
        {isLoading ? "Analizando..." : "Analizar y sugerir"}
      </Button>
    </div>
  );
}
