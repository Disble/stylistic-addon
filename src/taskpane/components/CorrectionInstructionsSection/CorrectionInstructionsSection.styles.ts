import { makeStyles, tokens } from "@fluentui/react-components";

/** Griffel styles for the correction-instructions section. */
export const useCorrectionInstructionsSectionStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalXS,
    marginBottom: tokens.spacingVerticalM,
  },
  header: {
    display: "flex",
    alignItems: "center",
    columnGap: tokens.spacingHorizontalXS,
  },
  title: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  description: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
    marginBottom: tokens.spacingVerticalXS,
  },
  textarea: {
    width: "100%",
    minHeight: "120px",
    fontFamily: tokens.fontFamilyBase,
  },
  footer: {
    display: "flex",
    justifyContent: "flex-end",
  },
  counter: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    fontVariantNumeric: "tabular-nums",
  },
  counterLimit: {
    color: tokens.colorPaletteRedForeground1,
    fontWeight: tokens.fontWeightSemibold,
  },
});
