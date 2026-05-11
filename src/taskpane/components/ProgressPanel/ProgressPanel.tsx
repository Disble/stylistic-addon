import * as React from "react";
import { Button, Caption1, Field, ProgressBar, Tooltip } from "@fluentui/react-components";
import { ArrowClockwiseRegular, DismissRegular } from "@fluentui/react-icons";
import type { ProgressPanelProps } from "./ProgressPanel.types";
import { useProgressPanel } from "./ProgressPanel.hooks";

/**
 * Compact progress strip rendered above the partial results list. Shows the
 * message + icon-only cancel/retry actions on the top row and the determinate
 * progress bar beneath, so the user keeps scrolling through suggestions without
 * a centered button taking over the layout.
 */
export function ProgressPanel({
  progress,
  onCancelAnalysis,
  onRetryAnalysisQuery,
}: ProgressPanelProps): React.JSX.Element | null {
  const classes = useProgressPanel();
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
    <div className={classes.root} data-testid="progress-panel">
      <div className={classes.header}>
        <Caption1 className={classes.message} data-testid="progress-message">
          {progress.message}
        </Caption1>
        <div className={classes.actions}>
          {canRetryQuery && (
            <Tooltip content="Reintentar consulta" relationship="label" withArrow>
              <Button
                appearance="subtle"
                aria-label="Reintentar consulta"
                data-testid="retry-analysis-query-button"
                icon={<ArrowClockwiseRegular />}
                onClick={() => {
                  void onRetryAnalysisQuery();
                }}
                size="small"
                type="button"
              />
            </Tooltip>
          )}
          {canCancel && (
            <Tooltip content="Cancelar análisis" relationship="label" withArrow>
              <Button
                appearance="subtle"
                aria-label="Cancelar análisis"
                data-testid="cancel-analysis-button"
                icon={<DismissRegular />}
                onClick={() => {
                  void onCancelAnalysis();
                }}
                size="small"
                type="button"
              />
            </Tooltip>
          )}
        </div>
      </div>
      <Field validationState="none">
        <ProgressBar
          aria-label="Progreso del análisis"
          data-testid="progress-bar"
          max={total}
          value={value}
        />
      </Field>
    </div>
  );
}
