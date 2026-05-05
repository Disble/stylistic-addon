import * as React from "react";
import { Caption1, Field, ProgressBar } from "@fluentui/react-components";
import type { ProgressPanelProps } from "./ProgressPanel.types";
import { useProgressPanel } from "./useProgressPanel";

/** Renders the determinate pipeline progress bar with its message. */
export function ProgressPanel({ progress }: ProgressPanelProps): React.JSX.Element | null {
  const classes = useProgressPanel();
  if (!progress.visible) {
    return null;
  }

  const total = progress.total > 0 ? progress.total : 1;
  const value = Math.min(progress.current, total);

  return (
    <div className={classes.root} data-testid="progress-panel">
      <Field validationState="none">
        <ProgressBar
          aria-label="Progreso del análisis"
          data-testid="progress-bar"
          max={total}
          value={value}
        />
      </Field>
      <Caption1 className={classes.message} data-testid="progress-message">
        {progress.message}
      </Caption1>
    </div>
  );
}
