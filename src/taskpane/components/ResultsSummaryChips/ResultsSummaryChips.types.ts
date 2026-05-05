import type { ResultsPanelChipCounts, ResultsPanelFilter } from "../../ResultsPanelStore";

export type ResultsSummaryChipsProps = Readonly<{
  activeFilter: ResultsPanelFilter;
  counts: ResultsPanelChipCounts;
  onFilterChange: (filter: ResultsPanelFilter) => void;
  summaryText: string;
}>;

export type ResultsSummaryChipsClasses = Readonly<{
  root: string;
  chip: string;
  chipActive: string;
  chipLabel: string;
  chipCount: string;
}>;

export type ResultsSummaryChipDescriptor = Readonly<{
  filter: ResultsPanelFilter;
  label: string;
  count: number;
  ariaLabel: string;
  tone: "neutral" | "danger" | "warning" | "brand" | "success";
}>;
