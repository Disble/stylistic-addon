import * as React from "react";
import { Button, Spinner } from "@fluentui/react-components";
import { HistoryDismissRegular } from "@fluentui/react-icons";
import type { DisableTrackChangesSectionProps } from "./DisableTrackChangesSection.types";
import { useDisableTrackChangesSection } from "./DisableTrackChangesSection.hooks";

/** Renders the "disable Track Changes" CTA, warning-styled via palette tokens. */
export function DisableTrackChangesSection({
  isLoading,
  isVisible,
  onDisableTrackChanges,
}: DisableTrackChangesSectionProps): React.JSX.Element | null {
  const classes = useDisableTrackChangesSection();
  if (!isVisible) {
    return null;
  }

  const icon = isLoading ? (
    <Spinner size="tiny" data-testid="disable-track-changes-spinner" />
  ) : (
    <HistoryDismissRegular aria-hidden="true" />
  );

  return (
    <div className={classes.root}>
      <Button
        appearance="secondary"
        aria-label="Desactivar control de cambios"
        className={classes.button}
        data-testid="disable-track-changes-button"
        disabled={isLoading}
        icon={icon}
        onClick={() => {
          void onDisableTrackChanges();
        }}
        type="button"
      >
        {isLoading ? "Desactivando..." : "Desactivar control de cambios"}
      </Button>
    </div>
  );
}
