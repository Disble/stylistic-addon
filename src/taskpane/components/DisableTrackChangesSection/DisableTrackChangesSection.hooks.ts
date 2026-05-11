import type { DisableTrackChangesSectionClasses } from "./DisableTrackChangesSection.types";
import { useDisableTrackChangesSectionStyles } from "./DisableTrackChangesSection.styles";

/** Returns Griffel classes for the disable-track-changes section. */
export function useDisableTrackChangesSection(): DisableTrackChangesSectionClasses {
  const styles = useDisableTrackChangesSectionStyles();
  return {
    root: styles.root,
    button: styles.button,
  };
}
