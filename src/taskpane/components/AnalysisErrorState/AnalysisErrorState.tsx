import * as React from "react";
import { Button, Caption1 } from "@fluentui/react-components";
import { AnalysisHeroDocumentIllustration } from "../AnalysisHeroDocumentIllustration";
import { AnalysisHeroShell } from "../AnalysisHeroShell";
import { ANALYSIS_ERROR_TITLE_ID } from "./AnalysisErrorState.constants";
import { resolveRetryPresentation } from "./AnalysisErrorState.helpers";
import type { AnalysisErrorStateProps } from "./AnalysisErrorState.types";
import { useAnalysisErrorState } from "./AnalysisErrorState.hooks";

/**
 * Hero-shaped analysis-error surface that takes over the workflow viewport when
 * an analysis attempt fails and there is no progress or results competing for
 * attention. Mirrors the idle hero composition: centered illustration, Fluent
 * typography stack, and a primary retry CTA chosen from the retry kind so the
 * user either resubmits the analysis or only re-polls a known backend run.
 */
export function AnalysisErrorState({
  error,
  onRetryAnalysis,
  onRetryAnalysisQuery,
  children,
}: AnalysisErrorStateProps): React.JSX.Element | null {
  const classes = useAnalysisErrorState();
  if (!error.visible) {
    return null;
  }

  const presentation = resolveRetryPresentation(error.retryKind);
  const handleRetry = error.retryKind === "retry-query" ? onRetryAnalysisQuery : onRetryAnalysis;

  return (
    <AnalysisHeroShell
      actions={
        <>
          <Button
            appearance="primary"
            className={classes.retryButton}
            data-testid="analysis-error-retry-button"
            onClick={() => {
              void handleRetry();
            }}
            size="large"
            type="button"
          >
            {presentation.actionLabel}
          </Button>
          {children}
        </>
      }
      dataRetryKind={error.retryKind}
      illustration={
        <AnalysisHeroDocumentIllustration
          documentClassName={classes.illustrationDoc}
          svgClassName={classes.illustration}
        >
          <rect fill="currentColor" height="5" opacity="0.30" rx="2.5" width="68" x="68" y="54" />
          <rect fill="currentColor" height="5" opacity="0.18" rx="2.5" width="92" x="68" y="68" />
          <rect fill="currentColor" height="5" opacity="0.18" rx="2.5" width="56" x="68" y="82" />

          <rect fill="currentColor" height="5" opacity="0.18" rx="2.5" width="92" x="68" y="106" />
          <rect fill="currentColor" height="5" opacity="0.18" rx="2.5" width="78" x="68" y="120" />
          <rect fill="currentColor" height="5" opacity="0.18" rx="2.5" width="44" x="68" y="134" />

          <g className={classes.illustrationAlert}>
            <circle cx="170" cy="44" fill="white" r="26" />
            <circle cx="170" cy="44" fill="currentColor" r="22" />
            <rect fill="white" height="16" rx="2" width="4" x="168" y="32" />
            <circle cx="170" cy="56" fill="white" r="2.5" />
          </g>
        </AnalysisHeroDocumentIllustration>
      }
      message={error.message}
      messageSupplement={<Caption1 className={classes.guidance}>{presentation.guidance}</Caption1>}
      sectionLabelledById={ANALYSIS_ERROR_TITLE_ID}
      testId="analysis-error-state"
      title={presentation.title}
    />
  );
}
