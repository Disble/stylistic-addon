import * as React from "react";
import { Field, Spinner, Textarea } from "@fluentui/react-components";
import {
  CORRECTION_INSTRUCTIONS_ARIA_LABEL,
  CORRECTION_INSTRUCTIONS_DESCRIPTION,
  CORRECTION_INSTRUCTIONS_PLACEHOLDER,
  CORRECTION_INSTRUCTIONS_TITLE,
} from "./CorrectionInstructionsSection.constants";
import { useCorrectionInstructionsSection } from "./CorrectionInstructionsSection.hooks";
import type { CorrectionInstructionsSectionProps } from "./CorrectionInstructionsSection.types";

/**
 * Renders the global correction-instructions form section: a labelled textarea
 * and a live `X / max` counter. The component is purely presentational — the
 * form draft, dirty tracking, and persistence live in the parent settings view.
 *
 * Load errors are intentionally NOT rendered here. The parent hook swallows
 * load failures because surfacing a passive error at form open is bad UX —
 * actionable feedback is reserved for the Save click.
 */
export function CorrectionInstructionsSection({
  value,
  maxLength,
  isDisabled,
  isLoading,
  onChange,
}: CorrectionInstructionsSectionProps): React.JSX.Element {
  const { classes, isAtLimit, counterLabel } = useCorrectionInstructionsSection(value, maxLength);

  return (
    <section className={classes.root} aria-label={CORRECTION_INSTRUCTIONS_ARIA_LABEL}>
      <div className={classes.header}>
        <span className={classes.title}>{CORRECTION_INSTRUCTIONS_TITLE}</span>
        {isLoading && <Spinner size="tiny" aria-label="Cargando instrucciones" />}
      </div>
      <p className={classes.description}>{CORRECTION_INSTRUCTIONS_DESCRIPTION}</p>
      <Field>
        <Textarea
          aria-label={CORRECTION_INSTRUCTIONS_ARIA_LABEL}
          className={classes.textarea}
          data-testid="correction-instructions-textarea"
          disabled={isDisabled}
          maxLength={maxLength}
          onChange={(_event, data) => onChange(data.value)}
          placeholder={CORRECTION_INSTRUCTIONS_PLACEHOLDER}
          resize="vertical"
          value={value}
        />
      </Field>
      <div className={classes.footer}>
        <span
          aria-live="polite"
          className={isAtLimit ? `${classes.counter} ${classes.counterLimit}` : classes.counter}
          data-testid="correction-instructions-counter"
        >
          {counterLabel}
        </span>
      </div>
    </section>
  );
}
