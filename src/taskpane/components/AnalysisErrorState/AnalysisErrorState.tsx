import * as React from "react";
import { Body1, Button, Caption1, Title3 } from "@fluentui/react-components";
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

  const isRetryQuery = error.retryKind === "retry-query";
  const title = isRetryQuery
    ? "No pudimos recuperar el resultado"
    : "El análisis no pudo completarse";
  const guidance = isRetryQuery
    ? "Reintentá la consulta del mismo run sin reenviar el texto al backend."
    : "Reintentá el análisis completo para volver a enviar la petición al backend.";
  const actionLabel = isRetryQuery ? "Reintentar consulta" : "Reintentar análisis";
  const handleRetry = isRetryQuery ? onRetryAnalysisQuery : onRetryAnalysis;

  return (
    <section
      aria-labelledby="analysis-error-state-title"
      className={classes.root}
      data-retry-kind={error.retryKind}
      data-testid="analysis-error-state"
    >
      <div className={classes.illustrationWrapper}>
        <svg
          aria-hidden="true"
          className={classes.illustration}
          fill="none"
          role="img"
          viewBox="0 0 240 200"
          xmlns="http://www.w3.org/2000/svg"
        >
          <ellipse cx="120" cy="184" fill="currentColor" opacity="0.08" rx="78" ry="6" />

          <g className={classes.illustrationDoc}>
            <rect
              fill="currentColor"
              height="130"
              opacity="0.06"
              rx="8"
              width="118"
              x="64"
              y="38"
            />
            <rect
              fill="white"
              height="130"
              rx="8"
              stroke="currentColor"
              strokeWidth="1.5"
              width="118"
              x="56"
              y="32"
            />

            <rect
              fill="currentColor"
              height="5"
              opacity="0.30"
              rx="2.5"
              width="68"
              x="68"
              y="54"
            />
            <rect
              fill="currentColor"
              height="5"
              opacity="0.18"
              rx="2.5"
              width="92"
              x="68"
              y="68"
            />
            <rect
              fill="currentColor"
              height="5"
              opacity="0.18"
              rx="2.5"
              width="56"
              x="68"
              y="82"
            />

            <rect
              fill="currentColor"
              height="5"
              opacity="0.18"
              rx="2.5"
              width="92"
              x="68"
              y="106"
            />
            <rect
              fill="currentColor"
              height="5"
              opacity="0.18"
              rx="2.5"
              width="78"
              x="68"
              y="120"
            />
            <rect
              fill="currentColor"
              height="5"
              opacity="0.18"
              rx="2.5"
              width="44"
              x="68"
              y="134"
            />
          </g>

          <g className={classes.illustrationAlert}>
            <circle cx="170" cy="44" fill="white" r="26" />
            <circle cx="170" cy="44" fill="currentColor" r="22" />
            <rect fill="white" height="16" rx="2" width="4" x="168" y="32" />
            <circle cx="170" cy="56" fill="white" r="2.5" />
          </g>
        </svg>
        <div className={classes.copy}>
          <Title3 as="h2" className={classes.title} id="analysis-error-state-title">
            {title}
          </Title3>
          <Body1 className={classes.message}>{error.message}</Body1>
          <Caption1 className={classes.guidance}>{guidance}</Caption1>
        </div>
      </div>
      <div className={classes.actions}>
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
          {actionLabel}
        </Button>
        {children}
      </div>
    </section>
  );
}
