import { mergeClasses } from "@fluentui/react-components";
import type {
  ResultsSummaryChipDescriptor,
  ResultsSummaryChipsClasses,
} from "./ResultsSummaryChips.types";
import type { ResultsPanelChipCounts } from "../../ResultsPanelFilters.types";

/** Returns the list of chip descriptors to render, hiding state buckets with zero counts. */
export function buildResultsSummaryChips(
  counts: ResultsPanelChipCounts
): readonly ResultsSummaryChipDescriptor[] {
  const chips: ResultsSummaryChipDescriptor[] = [
    {
      filter: "all",
      label: "Todas",
      count: counts.all,
      ariaLabel: `Mostrar todas las sugerencias (${counts.all})`,
      tone: "neutral",
    },
  ];
  if (counts.high > 0) {
    chips.push({
      filter: "high",
      label: "Alta",
      count: counts.high,
      ariaLabel: `Mostrar sugerencias de severidad alta (${counts.high})`,
      tone: "danger",
    });
  }
  if (counts.medium > 0) {
    chips.push({
      filter: "medium",
      label: "Media",
      count: counts.medium,
      ariaLabel: `Mostrar sugerencias de severidad media (${counts.medium})`,
      tone: "warning",
    });
  }
  if (counts.low > 0) {
    chips.push({
      filter: "low",
      label: "Baja",
      count: counts.low,
      ariaLabel: `Mostrar sugerencias de severidad baja (${counts.low})`,
      tone: "brand",
    });
  }
  if (counts.accepted > 0) {
    chips.push({
      filter: "accepted",
      label: "Aceptadas",
      count: counts.accepted,
      ariaLabel: `Mostrar sugerencias aceptadas (${counts.accepted})`,
      tone: "success",
    });
  }
  if (counts.rejected > 0) {
    chips.push({
      filter: "rejected",
      label: "Rechazadas",
      count: counts.rejected,
      ariaLabel: `Mostrar sugerencias rechazadas (${counts.rejected})`,
      tone: "neutral",
    });
  }
  if (counts.failed > 0) {
    chips.push({
      filter: "failed",
      label: "No encontradas",
      count: counts.failed,
      ariaLabel: `Mostrar sugerencias no encontradas (${counts.failed})`,
      tone: "warning",
    });
  }
  return chips;
}

/** Returns the merged className for a chip given its active state. */
export function resolveChipClassName(
  classes: ResultsSummaryChipsClasses,
  isActive: boolean
): string {
  return isActive ? mergeClasses(classes.chip, classes.chipActive) : classes.chip;
}
