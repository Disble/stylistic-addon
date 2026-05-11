/** Diagnostic payload for one apply-suggestion range candidate. */
export interface ApplySuggestionRangeCandidateDiagnostics {
  /** Plain `Range.text` exposed by Word for the inspected candidate. */
  text: string;

  /** Current reviewed text exposed by Word. */
  current: string;

  /** Original reviewed text exposed by Word. */
  original: string;

  /** Whether the candidate satisfies the current-side-only predicate. */
  passes: boolean;
}

/** Diagnostic payload for pre-mutation apply-suggestion scope. */
export interface ApplySuggestionPreMutationScopeDiagnostics {
  /** Mutation classification selected for the suggestion. */
  changeType: string;

  /** Word document change-tracking mode before the mutation. */
  changeTrackingMode: string;

  /** Suggestion anchor being mutated. */
  anchor: string;

  /** Suggested text that will be inserted. */
  suggestedText: string;

  /** Length of the containing paragraph before mutation. */
  paragraphLength: number;

  /** Short preview of the containing paragraph before mutation. */
  paragraphPreview: string;
}
