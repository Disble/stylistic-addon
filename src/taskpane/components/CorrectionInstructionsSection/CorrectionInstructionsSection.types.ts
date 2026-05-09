/** Props consumed by the correction-instructions form section. */
export type CorrectionInstructionsSectionProps = Readonly<{
  value: string;
  maxLength: number;
  isDisabled: boolean;
  isLoading: boolean;
  onChange: (value: string) => void;
}>;

/** Griffel class slots consumed by the correction-instructions section. */
export type CorrectionInstructionsSectionClasses = Readonly<{
  root: string;
  header: string;
  title: string;
  description: string;
  textarea: string;
  footer: string;
  counter: string;
  counterLimit: string;
}>;

/** View-model returned by the section's hook for rendering decisions. */
export type CorrectionInstructionsSectionState = Readonly<{
  classes: CorrectionInstructionsSectionClasses;
  isAtLimit: boolean;
  counterLabel: string;
}>;
