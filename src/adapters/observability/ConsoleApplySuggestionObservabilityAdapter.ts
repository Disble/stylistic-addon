/* global console */

import type { Suggestion } from "../../domain/suggestion/Suggestion.types";

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

/** Best-effort console diagnostics for apply-suggestion Word mutations. */
export class ConsoleApplySuggestionObservabilityAdapter {
  /** Emits the prepared application order for a batch. */
  logPreparedApplicationOrder(
    suggestions: Array<Pick<Suggestion, "id" | "type" | "positionHint">>,
  ): void {
    console.log(
      "📝 [BatchApplyOrchestrator] Orden de aplicación preparado",
      suggestions.map((suggestion) => ({
        id: suggestion.id,
        type: suggestion.type,
        hasPositionHint: suggestion.positionHint !== undefined,
        positionHint: suggestion.positionHint,
      })),
    );
  }

  /** Emits the pre-mutation Word scope for one suggestion. */
  logPreMutationScope(
    commandId: string,
    diagnostics: ApplySuggestionPreMutationScopeDiagnostics,
  ): void {
    console.log(
      `🔎 [ApplySuggestionCommand] "${commandId}": pre-mutation scope`,
      diagnostics,
    );
  }

  /** Emits the resolved mutation target shape for one suggestion. */
  logMutationTargetResolved(
    commandId: string,
    diagnostics: { usesOperationalWrapper: boolean; insertLocation: string },
  ): void {
    console.log(
      `🔎 [ApplySuggestionCommand] "${commandId}": mutation target resolved`,
      diagnostics,
    );
  }

  /** Emits that Word `insertText` was issued and resolver diagnostics follow. */
  logInsertTextIssued(commandId: string): void {
    console.log(
      `🔎 [ApplySuggestionCommand] "${commandId}": insertText issued; entering annotation resolver`,
    );
  }

  /** Emits a post-mutation failure where metadata range isolation failed. */
  warnPostMutationIsolationFailure(commandId: string): void {
    console.warn(
      `⚠️ [ApplySuggestionCommand] "${commandId}": mutation completed but annotation range was not isolated`,
    );
  }

  /** Emits parent content-control state before wrapper resolution. */
  logParentContentControl(
    commandId: string,
    diagnostics: {
      hasParentContentControl: boolean;
      tag: string;
      title: string;
    },
  ): void {
    console.log(
      `🔎 [ApplySuggestionCommand] "${commandId}": parent content control before wrapper resolution`,
      diagnostics,
    );
  }

  /** Emits operational wrapper creation diagnostics. */
  logCreatingOperationalWrapper(
    commandId: string,
    diagnostics: { previousTrackingMode: string },
  ): void {
    console.log(
      `🔎 [ApplySuggestionCommand] "${commandId}": creating operational wrapper`,
      diagnostics,
    );
  }

  /** Emits a non-operational Stylistic content-control coverage warning. */
  warnNonOperationalStylisticContentControl(
    commandId: string,
    diagnostics: { existingTag: string },
  ): void {
    console.warn(
      `⚠️ [ApplySuggestionCommand] "${commandId}": anchor covered by non-operational Stylistic CC`,
      diagnostics,
    );
  }

  /** Emits an operational wrapper identity mismatch warning. */
  warnOperationalWrapperIdentityMismatch(
    commandId: string,
    diagnostics: { existingTag: string; title: string },
  ): void {
    console.warn(
      `⚠️ [ApplySuggestionCommand] "${commandId}": operational wrapper identity mismatch`,
      diagnostics,
    );
  }

  /** Emits operational wrapper reuse diagnostics. */
  logReusingOperationalWrapper(
    commandId: string,
    diagnostics: { existingTag: string; title: string },
  ): void {
    console.log(
      `🔎 [ApplySuggestionCommand] "${commandId}": reusing operational wrapper`,
      diagnostics,
    );
  }

  /** Emits the reviewed-text shape of one resolver candidate. */
  logResolverCandidate(
    commandId: string,
    label: string,
    diagnostics: ApplySuggestionRangeCandidateDiagnostics,
  ): void {
    console.log(
      `🔎 [ApplySuggestionCommand] "${commandId}": resolver ${label}`,
      diagnostics,
    );
  }

  /** Emits that one resolver candidate could not be located. */
  logResolverCandidateNotFound(commandId: string, label: string): void {
    console.log(
      `🔎 [ApplySuggestionCommand] "${commandId}": resolver ${label} candidate not found`,
    );
  }

  /** Emits which resolver candidate was selected. */
  logResolverSelected(commandId: string, label: string): void {
    console.log(
      `🔎 [ApplySuggestionCommand] "${commandId}": resolver selected ${label}`,
    );
  }
}

/** Shared console apply-suggestion observability adapter. */
export const applySuggestionObservability =
  new ConsoleApplySuggestionObservabilityAdapter();
