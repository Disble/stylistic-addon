/* global Word, console */

/**
 * BatchApplyOrchestrator — coordinates the sequential application of multiple
 * suggestions as tracked changes in Word.
 *
 * Each suggestion runs in its own `Word.run` context (per-suggestion isolation)
 * to avoid stale ranges after OOXML insertions. Suggestions are applied in
 * reverse document order (end-first) to prevent Content Control boundary
 * interference. Command failures are aggregated so later suggestions can still
 * run (partial success philosophy).
 *
 * Dependencies are injected via the constructor so the orchestrator has no
 * direct coupling to Office.js or WordAdapter internals.
 *
 * @module BatchApplyOrchestrator
 */

import type { DocumentReviewUiState } from "../../domain/review/DocumentReviewStateMachine";
import type {
  ApplySuggestionsResult,
  CommandResult,
  DocumentReviewState,
  ProgressCallback,
  Suggestion,
  SuggestionApplicationFailure,
  SuggestionApplicationFailureReason,
} from "../../domain/types";
import { ApplySuggestionCommand } from "./ApplySuggestionCommand";
import { getDefaultTextLocator } from "./WordTextLocatorContext";

const textLocator = getDefaultTextLocator();

/** Injected capabilities required by the orchestrator. */
type BatchApplyDependencies = {
  /** Enables Track Changes lazily; returns `true` when newly activated. */
  ensureTrackChangesActive: () => Promise<boolean>;
  /** Returns the current document-derived review state. */
  getDocumentReviewState: () => Promise<DocumentReviewState>;
  /** Derives the explicit UI state from a review snapshot. */
  deriveDocumentState: (state: DocumentReviewState) => DocumentReviewUiState;
};

/**
 * Orchestrates batch suggestion application.
 *
 * Usage:
 * ```ts
 * const orchestrator = new BatchApplyOrchestrator(dependencies);
 * const result = await orchestrator.run(suggestions, onProgress);
 * ```
 */
export class BatchApplyOrchestrator {
  constructor(private readonly deps: BatchApplyDependencies) {}

  /**
   * Applies suggestions as tracked changes, returning an aggregate result.
   */
  async run(
    suggestions: Suggestion[],
    onProgress?: ProgressCallback,
  ): Promise<ApplySuggestionsResult> {
    console.log(
      `📝 [BatchApplyOrchestrator] applySuggestions: ${suggestions.length} sugerencias`,
    );

    if (suggestions.length === 0) {
      const pendingAfter = await this.deps.getDocumentReviewState();
      return {
        successCount: 0,
        failedSuggestions: [],
        pendingAfter,
        documentState: this.deps.deriveDocumentState(pendingAfter),
        trackChangesActivatedForBatch: false,
      };
    }

    const sortedSuggestions = this.sortByDocumentPosition(suggestions);

    const failedSuggestions: SuggestionApplicationFailure[] = [];
    let successCount = 0;
    let trackChangesPrepared = false;
    let trackChangesActivatedForBatch = false;

    for (const suggestion of sortedSuggestions) {
      const trackChangesState = await this.prepareTrackChangesForSuggestion(
        suggestion,
        trackChangesPrepared,
      );
      trackChangesPrepared = trackChangesState.trackChangesPrepared;
      if (trackChangesState.activatedForBatch) {
        trackChangesActivatedForBatch = true;
      }

      const commandResult = await this.executeSuggestionCommand(suggestion);
      successCount = this.registerSuggestionOutcome(
        suggestion,
        commandResult,
        failedSuggestions,
        successCount,
      );

      this.rebasePendingSnapshotHints(
        sortedSuggestions,
        suggestion.id,
        commandResult,
      );

      this.reportApplyProgress(
        onProgress,
        successCount + failedSuggestions.length,
        suggestions.length,
      );
    }

    console.log(
      `📝 [BatchApplyOrchestrator] Completado: ${successCount} éxitos, ${failedSuggestions.length} fallos`,
    );

    const pendingAfter = await this.deps.getDocumentReviewState();

    return {
      successCount,
      failedSuggestions,
      pendingAfter,
      documentState: this.deps.deriveDocumentState(pendingAfter),
      trackChangesActivatedForBatch,
    };
  }

  // -------------------------------------------------------------------------
  // Internal helpers
  // -------------------------------------------------------------------------

  /**
   * Sorts suggestions by snapshot-derived document position when available,
   * falling back to reverse array order only for legacy suggestion batches.
   */
  private sortByDocumentPosition(suggestions: Suggestion[]): Suggestion[] {
    if (suggestions.length <= 1) return suggestions;

    const allHavePositionHints = suggestions.every(
      (suggestion) => suggestion.positionHint?.source === "snapshot",
    );

    if (allHavePositionHints) {
      return [...suggestions].sort((left, right) => {
        const leftRequiresReread =
          left.positionHint?.requiresLocalReread === true;
        const rightRequiresReread =
          right.positionHint?.requiresLocalReread === true;

        if (leftRequiresReread !== rightRequiresReread) {
          return leftRequiresReread ? 1 : -1;
        }

        const endDifference =
          (right.positionHint?.end ?? 0) - (left.positionHint?.end ?? 0);

        if (endDifference !== 0) {
          return endDifference;
        }

        return (
          (right.positionHint?.start ?? 0) - (left.positionHint?.start ?? 0)
        );
      });
    }

    return [...suggestions].reverse();
  }

  /**
   * Activates Track Changes once per batch when the current suggestion requires it.
   */
  private async prepareTrackChangesForSuggestion(
    suggestion: Suggestion,
    trackChangesPrepared: boolean,
  ): Promise<{ trackChangesPrepared: boolean; activatedForBatch: boolean }> {
    if (trackChangesPrepared || suggestion.type !== "track-change") {
      return { trackChangesPrepared, activatedForBatch: false };
    }

    const activated = await this.deps.ensureTrackChangesActive();

    return {
      trackChangesPrepared: true,
      activatedForBatch: activated,
    };
  }

  /**
   * Executes one suggestion command and normalizes unexpected thrown errors.
   */
  private async executeSuggestionCommand(
    suggestion: Suggestion,
  ): Promise<CommandResult> {
    const command = new ApplySuggestionCommand(suggestion, textLocator);

    try {
      return await command.execute();
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return {
        success: false,
        commandId: suggestion.id,
        error: message,
      };
    }
  }

  /**
   * Infers a stable failure reason from a command error message.
   */
  private inferApplicationFailureReason(
    commandResult: CommandResult,
  ): SuggestionApplicationFailureReason {
    const message = commandResult.error?.toLowerCase() ?? "";

    if (
      message.includes("anchor no encontrado") ||
      message.includes("texto original no encontrado")
    ) {
      return "not-found";
    }

    if (
      message.includes("cc existente") ||
      message.includes("content control")
    ) {
      return "covered-by-existing-cc";
    }

    return "command-error";
  }

  /**
   * Updates counters and logs after one command execution.
   */
  private registerSuggestionOutcome(
    suggestion: Suggestion,
    commandResult: CommandResult,
    failedSuggestions: SuggestionApplicationFailure[],
    successCount: number,
  ): number {
    if (commandResult.success) {
      console.log(`✅ [BatchApplyOrchestrator] "${suggestion.id}" aplicada`);
      return successCount + 1;
    }

    failedSuggestions.push({
      suggestion,
      reason: this.inferApplicationFailureReason(commandResult),
      message: commandResult.error ?? "Error desconocido al aplicar sugerencia",
    });
    console.warn(
      `⚠️ [BatchApplyOrchestrator] "${suggestion.id}" falló: ${commandResult.error}`,
    );
    return successCount;
  }

  /**
   * Reports progress through the optional callback.
   */
  private reportApplyProgress(
    onProgress: ProgressCallback | undefined,
    completedCount: number,
    total: number,
  ): void {
    if (!onProgress) {
      return;
    }

    onProgress(
      "applying",
      completedCount,
      total,
      `Aplicando sugerencia ${completedCount} de ${total}...`,
    );
  }

  /** Applies a minimal delta rebase to later snapshot hints after a successful patch. */
  private rebasePendingSnapshotHints(
    suggestions: Suggestion[],
    appliedSuggestionId: string,
    commandResult: CommandResult,
  ): void {
    const patch = commandResult.mutationPatch;
    if (!commandResult.success || !patch || patch.deltaLength === 0) {
      return;
    }

    for (const suggestion of suggestions) {
      if (suggestion.id === appliedSuggestionId) {
        continue;
      }

      const hint = suggestion.positionHint;
      if (!hint || hint.source !== "snapshot") {
        continue;
      }

      if (hint.end > patch.affectedStart && hint.start < patch.affectedEnd) {
        suggestion.positionHint = {
          ...hint,
          requiresLocalReread: true,
        };
        continue;
      }

      if (hint.start >= patch.affectedEnd) {
        suggestion.positionHint = {
          ...hint,
          start: hint.start + patch.deltaLength,
          end: hint.end + patch.deltaLength,
        };
      }
    }
  }
}
