import * as React from "react";
import type { ResultsSummaryChipsProps } from "./ResultsSummaryChips.types";
import {
  buildResultsSummaryChips,
  resolveChipClassName,
} from "./ResultsSummaryChips.helpers";
import {
  useResultsSummaryChips,
} from "./ResultsSummaryChips.hooks";

/** Renders the sticky chips toolbar that filters the results list. */
export function ResultsSummaryChips({
  activeFilter,
  counts,
  onFilterChange,
  summaryText,
}: ResultsSummaryChipsProps): React.JSX.Element {
  const classes = useResultsSummaryChips();
  const chips = buildResultsSummaryChips(counts);

  return (
    <div
      aria-label={summaryText || "Filtrar sugerencias"}
      className={classes.root}
      data-testid="results-summary-chips"
      role="group"
    >
      {chips.map((chip) => {
        const isActive = chip.filter === activeFilter;
        return (
          <button
            aria-label={chip.ariaLabel}
            aria-pressed={isActive}
            className={resolveChipClassName(classes, isActive)}
            data-active={isActive ? "true" : "false"}
            data-filter={chip.filter}
            data-testid={`results-chip-${chip.filter}`}
            key={chip.filter}
            onClick={() => onFilterChange(chip.filter)}
            type="button"
          >
            <span className={classes.chipLabel}>{chip.label}</span>
            <span className={classes.chipCount}>{chip.count}</span>
          </button>
        );
      })}
    </div>
  );
}
