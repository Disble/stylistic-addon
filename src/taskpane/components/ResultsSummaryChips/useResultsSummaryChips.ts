import { makeStyles, mergeClasses, tokens } from "@fluentui/react-components";
import type {
  ResultsSummaryChipDescriptor,
  ResultsSummaryChipsClasses,
} from "./ResultsSummaryChips.types";
import type { ResultsPanelChipCounts } from "../../ResultsPanelStore";

const useResultsSummaryChipsStyles = makeStyles({
  root: {
    position: "sticky",
    top: 0,
    zIndex: 1,
    display: "flex",
    flexWrap: "wrap",
    columnGap: tokens.spacingHorizontalXS,
    rowGap: tokens.spacingVerticalXS,
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalS,
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottomWidth: "1px",
    borderBottomStyle: "solid",
    borderBottomColor: tokens.colorNeutralStroke2,
  },
  chip: {
    display: "inline-flex",
    alignItems: "center",
    columnGap: tokens.spacingHorizontalXXS,
    paddingTop: tokens.spacingVerticalXXS,
    paddingBottom: tokens.spacingVerticalXXS,
    paddingLeft: tokens.spacingHorizontalS,
    paddingRight: tokens.spacingHorizontalS,
    borderTopWidth: "1px",
    borderRightWidth: "1px",
    borderBottomWidth: "1px",
    borderLeftWidth: "1px",
    borderTopStyle: "solid",
    borderRightStyle: "solid",
    borderBottomStyle: "solid",
    borderLeftStyle: "solid",
    borderTopColor: tokens.colorNeutralStroke1,
    borderRightColor: tokens.colorNeutralStroke1,
    borderBottomColor: tokens.colorNeutralStroke1,
    borderLeftColor: tokens.colorNeutralStroke1,
    borderTopLeftRadius: tokens.borderRadiusCircular,
    borderTopRightRadius: tokens.borderRadiusCircular,
    borderBottomLeftRadius: tokens.borderRadiusCircular,
    borderBottomRightRadius: tokens.borderRadiusCircular,
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    cursor: "pointer",
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
    ":focus-visible": {
      outlineWidth: "2px",
      outlineStyle: "solid",
      outlineColor: tokens.colorStrokeFocus2,
    },
  },
  chipActive: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    borderTopColor: tokens.colorBrandBackground,
    borderRightColor: tokens.colorBrandBackground,
    borderBottomColor: tokens.colorBrandBackground,
    borderLeftColor: tokens.colorBrandBackground,
    ":hover": {
      backgroundColor: tokens.colorBrandBackgroundHover,
    },
  },
  chipLabel: {
    fontWeight: tokens.fontWeightSemibold,
  },
  chipCount: {
    fontVariantNumeric: "tabular-nums",
    opacity: 0.85,
  },
});

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
