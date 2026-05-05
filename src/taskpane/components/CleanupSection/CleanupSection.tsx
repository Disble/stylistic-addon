import * as React from "react";
import { Button, Spinner } from "@fluentui/react-components";
import { BroomRegular } from "@fluentui/react-icons";
import type { CleanupSectionProps } from "./CleanupSection.types";
import { useCleanupSection } from "./useCleanupSection";

/** Renders the resolved-comments cleanup CTA. */
export function CleanupSection({
  isLoading,
  isVisible,
  onCleanup,
}: CleanupSectionProps): React.JSX.Element | null {
  const classes = useCleanupSection();
  if (!isVisible) {
    return null;
  }

  const icon = isLoading ? (
    <Spinner size="tiny" data-testid="cleanup-spinner" />
  ) : (
    <BroomRegular aria-hidden="true" />
  );

  return (
    <div className={classes.root}>
      <Button
        appearance="secondary"
        aria-label="Limpiar comentarios resueltos"
        className={classes.button}
        data-testid="cleanup-button"
        disabled={isLoading}
        icon={icon}
        onClick={() => {
          void onCleanup();
        }}
        type="button"
      >
        {isLoading ? "Limpiando..." : "Limpiar comentarios resueltos"}
      </Button>
    </div>
  );
}
