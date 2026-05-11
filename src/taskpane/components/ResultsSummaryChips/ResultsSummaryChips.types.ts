import type { ResultsPanelChipCounts, ResultsPanelFilter } from "../../ResultsPanelFilters.types";

/** Props required to render results-summary filter chips. */
export type ResultsSummaryChipsProps = Readonly<{
  activeFilter: ResultsPanelFilter;
  counts: ResultsPanelChipCounts;
  onFilterChange: (filter: ResultsPanelFilter) => void;
  summaryText: string;
}>;

/** Griffel class slots consumed by the results-summary chips. */
export type ResultsSummaryChipsClasses = Readonly<{
  root: string;
  chip: string;
  chipActive: string;
  chipLabel: string;
  chipCount: string;
}>;

/** View model for one summary chip rendered above the results list. */
export type ResultsSummaryChipDescriptor = Readonly<{
  filter: ResultsPanelFilter;
  label: string;
  count: number;
  ariaLabel: string;
  tone: "neutral" | "danger" | "warning" | "brand" | "success";
}>;
