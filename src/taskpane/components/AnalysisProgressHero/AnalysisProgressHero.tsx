import * as React from "react";
import { Button, Field, ProgressBar, mergeClasses } from "@fluentui/react-components";
import { AnalysisHeroDocumentIllustration } from "../AnalysisHeroDocumentIllustration";
import { AnalysisHeroShell } from "../AnalysisHeroShell";
import { ANALYSIS_PROGRESS_TITLE_ID } from "./AnalysisProgressHero.constants";
import type { AnalysisProgressHeroProps } from "./AnalysisProgressHero.types";
import { useAnalysisProgressHero } from "./AnalysisProgressHero.hooks";

/**
 * Hero-shaped analysis-in-progress surface shown while the pipeline is working
 * and no result cards exist yet. Mirrors the idle / error hero composition:
 * centered animated illustration, Fluent typography stack, contextual progress
 * bar, and a stacked action area where Cancel and Retry-query CTAs live as
 * full-width buttons aligned with the rest of the workflow surfaces.
 */
export function AnalysisProgressHero({
  progress,
  onCancelAnalysis,
  onRetryAnalysisQuery,
}: AnalysisProgressHeroProps): React.JSX.Element | null {
  const classes = useAnalysisProgressHero();
  if (!progress.visible) {
    return null;
  }

  const total = progress.total > 0 ? progress.total : 1;
  const value = Math.min(progress.current, total);
  const canCancel =
    progress.asyncSession.phase === "polling" && progress.asyncSession.activeRuns.length > 0;
  const canRetryQuery =
    progress.asyncSession.phase === "retryable-failure" &&
    progress.asyncSession.retryableRuns.length > 0;

  return (
    <AnalysisHeroShell
      actions={
        <>
          {canRetryQuery && (
            <Button
              appearance="primary"
              className={classes.primaryButton}
              data-testid="analysis-progress-hero-retry-query-button"
              onClick={() => {
                void onRetryAnalysisQuery();
              }}
              size="large"
              type="button"
            >
              Reintentar consulta
            </Button>
          )}
          {canCancel && (
            <Button
              appearance={canRetryQuery ? "subtle" : "secondary"}
              className={classes.primaryButton}
              data-testid="analysis-progress-hero-cancel-button"
              onClick={() => {
                void onCancelAnalysis();
              }}
              size="large"
              type="button"
            >
              Cancelar
            </Button>
          )}
        </>
      }
      illustration={
        <AnalysisHeroDocumentIllustration
          documentClassName={classes.illustrationDoc}
          svgClassName={classes.illustration}
        >
          <rect
            className={mergeClasses(classes.illustrationLine, classes.illustrationLineDelay1)}
            fill="currentColor"
            height="5"
            opacity="0.30"
            rx="2.5"
            width="68"
            x="68"
            y="54"
          />
          <rect
            className={mergeClasses(classes.illustrationLine, classes.illustrationLineDelay2)}
            fill="currentColor"
            height="5"
            opacity="0.18"
            rx="2.5"
            width="92"
            x="68"
            y="68"
          />
          <rect
            className={mergeClasses(classes.illustrationLine, classes.illustrationLineDelay3)}
            fill="currentColor"
            height="5"
            opacity="0.18"
            rx="2.5"
            width="56"
            x="68"
            y="82"
          />

          <rect
            className={mergeClasses(classes.illustrationLine, classes.illustrationLineDelay4)}
            fill="currentColor"
            height="5"
            opacity="0.18"
            rx="2.5"
            width="92"
            x="68"
            y="106"
          />
          <rect
            className={mergeClasses(classes.illustrationLine, classes.illustrationLineDelay5)}
            fill="currentColor"
            height="5"
            opacity="0.18"
            rx="2.5"
            width="78"
            x="68"
            y="120"
          />
          <rect
            className={mergeClasses(classes.illustrationLine, classes.illustrationLineDelay6)}
            fill="currentColor"
            height="5"
            opacity="0.18"
            rx="2.5"
            width="44"
            x="68"
            y="134"
          />

          <g className={classes.illustrationWand}>
            <line
              stroke="currentColor"
              strokeLinecap="round"
              strokeWidth="4"
              x1="190"
              x2="148"
              y1="22"
              y2="64"
            />
            <path
              d="M 191 12 L 194.5 19 L 201.5 22 L 194.5 25 L 191 32 L 187.5 25 L 180.5 22 L 187.5 19 Z"
              fill="currentColor"
            />
          </g>

          <g className={mergeClasses(classes.illustrationSparkle, classes.illustrationSparkle1)}>
            <path
              d="M218 56 L219.2 60.5 L223.5 61.7 L219.2 62.9 L218 67.4 L216.8 62.9 L212.5 61.7 L216.8 60.5 Z"
              fill="currentColor"
            />
          </g>
          <g className={mergeClasses(classes.illustrationSparkle, classes.illustrationSparkle2)}>
            <path
              d="M28 96 L29 99.5 L32.5 100.5 L29 101.5 L28 105 L27 101.5 L23.5 100.5 L27 99.5 Z"
              fill="currentColor"
            />
          </g>
          <g className={mergeClasses(classes.illustrationSparkle, classes.illustrationSparkle3)}>
            <circle cx="206" cy="108" fill="currentColor" r="2.5" />
          </g>
        </AnalysisHeroDocumentIllustration>
      }
      message={progress.message}
      messageSupplement={
        <Field className={classes.progressField} validationState="none">
          <ProgressBar
            aria-label="Progreso del análisis"
            data-testid="analysis-progress-hero-bar"
            max={total}
            value={value}
          />
        </Field>
      }
      sectionLabelledById={ANALYSIS_PROGRESS_TITLE_ID}
      testId="analysis-progress-hero"
      title="Analizando tu texto"
    />
  );
}
