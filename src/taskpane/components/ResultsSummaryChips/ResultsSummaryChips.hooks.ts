import type { ResultsSummaryChipsClasses } from "./ResultsSummaryChips.types";
import { useResultsSummaryChipsStyles } from "./ResultsSummaryChips.styles";

/** Returns Griffel classes for the chips toolbar. */
export function useResultsSummaryChips(): ResultsSummaryChipsClasses {
  const styles = useResultsSummaryChipsStyles();
  return {
    root: styles.root,
    chip: styles.chip,
    chipActive: styles.chipActive,
    chipLabel: styles.chipLabel,
    chipCount: styles.chipCount,
  };
}
