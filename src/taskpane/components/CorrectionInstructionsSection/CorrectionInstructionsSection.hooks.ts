import { useCorrectionInstructionsSectionStyles } from "./CorrectionInstructionsSection.styles";
import type {
  CorrectionInstructionsSectionClasses,
  CorrectionInstructionsSectionState,
} from "./CorrectionInstructionsSection.types";

/** Returns Griffel classes for the correction-instructions section. */
function useCorrectionInstructionsSectionClasses(): CorrectionInstructionsSectionClasses {
  const styles = useCorrectionInstructionsSectionStyles();
  return {
    root: styles.root,
    header: styles.header,
    title: styles.title,
    description: styles.description,
    textarea: styles.textarea,
    footer: styles.footer,
    counter: styles.counter,
    counterLimit: styles.counterLimit,
  };
}

/**
 * Aggregates the state needed to render the correction-instructions section:
 * pre-built classes, the live counter label, and an `isAtLimit` flag the
 * component uses to switch the counter color when the user is at the limit.
 */
export function useCorrectionInstructionsSection(
  value: string,
  maxLength: number
): CorrectionInstructionsSectionState {
  const classes = useCorrectionInstructionsSectionClasses();
  const length = value.length;
  return {
    classes,
    isAtLimit: length >= maxLength,
    counterLabel: `${length} / ${maxLength}`,
  };
}
