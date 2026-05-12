import * as React from "react";
import { Body1, Title3 } from "@fluentui/react-components";
import type { AnalysisHeroShellProps } from "./AnalysisHeroShell.types";
import { useAnalysisHeroShell } from "./AnalysisHeroShell.hooks";

/** Renders the shared shell used by taskpane analysis hero states. */
export function AnalysisHeroShell({
  sectionLabelledById,
  testId,
  title,
  message,
  illustration,
  messageSupplement,
  actions,
  dataRetryKind,
}: AnalysisHeroShellProps): React.JSX.Element {
  const classes = useAnalysisHeroShell();

  return (
    <section
      aria-labelledby={sectionLabelledById}
      className={classes.root}
      {...(dataRetryKind ? { "data-retry-kind": dataRetryKind } : {})}
      data-testid={testId}
    >
      <div className={classes.illustrationWrapper}>
        {illustration}
        <div className={classes.copy}>
          <Title3 as="h2" className={classes.title} id={sectionLabelledById}>
            {title}
          </Title3>
          <Body1 className={classes.message}>{message}</Body1>
          {messageSupplement}
        </div>
      </div>
      {actions ? <div className={classes.actions}>{actions}</div> : null}
    </section>
  );
}
