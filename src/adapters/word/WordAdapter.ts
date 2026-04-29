/* global Word */

/**
 * Word Adapter — implements `IDocumentPort` using Office.js.
 *
 * This facade now delegates text reading, applied-suggestion inspection,
 * navigation, and Track Changes review state concerns to focused modules.
 *
 * @module WordAdapter
 */

import type {
  ApplySuggestionsResult,
  CommandResult,
} from "../../domain/DocumentApplication.types";
import type { ProgressCallback } from "../../domain/pipeline/PipelineEvents.types";
import type {
  IDocumentPort,
  IResolutionObservabilityPort,
} from "../../domain/ports";
import type { DocumentReviewState } from "../../domain/review/DocumentReviewStateMachine.types";
import type {
  Suggestion,
  SuggestionNavigationResult,
} from "../../domain/suggestion/Suggestion.types";
import type { SuggestionActionResult } from "../../domain/suggestion/SuggestionResolutionWorkflow.types";
import type { TextSource } from "../../domain/TextSource.types";
import { NoopResolutionObservabilityAdapter } from "../observability/NoopResolutionObservabilityAdapter";
import { BatchApplyOrchestrator } from "./BatchApplyOrchestrator";
import {
  cleanupResolvedComments,
  getCleanupPreview,
} from "./cleanup/CommentCleanup";
import { ResolveSuggestionCommand } from "./resolve-suggestion/ResolveSuggestionCommand";
import { WordAppliedSuggestionInspector } from "./WordAppliedSuggestionInspector";
import { WordSuggestionNavigationAdapter } from "./WordSuggestionNavigationAdapter";
import {
  getDefaultTextLocator,
  type TextLocator,
  type WordSearchContainer,
} from "./WordTextLocatorContext";
import { WordTextSourceAdapter } from "./WordTextSourceAdapter";
import { WordTrackChangesAdapter } from "./WordTrackChangesAdapter";

export class WordAdapter implements IDocumentPort {
  private readonly textSourceAdapter = new WordTextSourceAdapter();

  private readonly appliedSuggestionInspector =
    new WordAppliedSuggestionInspector();

  private readonly trackChangesAdapter = new WordTrackChangesAdapter();

  private readonly suggestionNavigationAdapter: WordSuggestionNavigationAdapter;

  constructor(
    private readonly textLocator: TextLocator = getDefaultTextLocator(),
    private readonly observabilityPort: IResolutionObservabilityPort = new NoopResolutionObservabilityAdapter(),
  ) {
    this.suggestionNavigationAdapter = new WordSuggestionNavigationAdapter(
      this.textLocator,
    );
  }

  /** Resolves the text to analyze from selection or full document. */
  async getTextToAnalyze(): Promise<TextSource> {
    return this.textSourceAdapter.getTextToAnalyze();
  }

  /** Returns the set of original texts already applied as Stylistic suggestions. */
  async getAppliedOriginalTexts(): Promise<Set<string>> {
    return this.appliedSuggestionInspector.getAppliedOriginalTexts();
  }

  /** Applies suggestions as tracked changes via `BatchApplyOrchestrator`. */
  async applySuggestions(
    suggestions: Suggestion[],
    onProgress?: ProgressCallback,
  ): Promise<ApplySuggestionsResult> {
    const orchestrator = new BatchApplyOrchestrator({
      ensureTrackChangesActive: () =>
        Word.run((ctx) =>
          this.trackChangesAdapter.ensureTrackChangesActive(ctx),
        ),
      getDocumentReviewState: () =>
        this.trackChangesAdapter.getDocumentReviewState(),
      deriveDocumentState: (state) =>
        this.trackChangesAdapter.deriveDocumentState(state),
      rereadSuggestionPositionHint: (suggestion, patch) =>
        this.rereadSuggestionPositionHint(suggestion, patch),
    });
    return orchestrator.run(suggestions, onProgress);
  }

  /** Rebuilds one localized hint from real Word after local patch reseed stops being trustworthy. */
  private async rereadSuggestionPositionHint(
    suggestion: Suggestion,
    patch: NonNullable<CommandResult["mutationPatch"]>,
  ): Promise<Suggestion["positionHint"] | undefined> {
    return Word.run(async (context) => {
      const body = context.document.body as unknown as WordSearchContainer;
      const localizedRange =
        await this.suggestionNavigationAdapter.searchWithFallback(
          context,
          body,
          patch.updatedText,
        );
      if (!localizedRange) {
        return undefined;
      }

      localizedRange.load("text");
      const containingParagraph = localizedRange.paragraphs
        .getFirst()
        .getRange("Whole");
      containingParagraph.load("text");
      await context.sync();

      const anchorRange =
        (await this.suggestionNavigationAdapter.searchWithFallback(
          context,
          localizedRange as unknown as WordSearchContainer,
          suggestion.anchor,
        )) ??
        (await this.suggestionNavigationAdapter.searchWithFallback(
          context,
          containingParagraph as unknown as WordSearchContainer,
          suggestion.anchor,
        ));

      if (!anchorRange) {
        return undefined;
      }

      const localizedStart = localizedRange.text.indexOf(suggestion.anchor);
      const paragraphStart = containingParagraph.text.indexOf(
        suggestion.anchor,
      );
      const start = localizedStart >= 0 ? localizedStart : paragraphStart;
      if (start < 0) {
        return undefined;
      }

      return {
        start,
        end: start + suggestion.anchor.length,
        snapshotVersion: patch.snapshotVersion,
        ...(patch.paragraphId ? { paragraphId: patch.paragraphId } : {}),
        source: "localized-reread",
      };
    });
  }

  /** Returns the current document-derived Stylistic review state. */
  async getDocumentReviewState(): Promise<DocumentReviewState> {
    return this.trackChangesAdapter.getDocumentReviewState();
  }

  /** Returns a dry-run summary of comments that can be deleted right now. */
  async getCleanupPreview(): Promise<{ deletable: number; kept: number }> {
    return getCleanupPreview();
  }

  /** Deletes Stylistic comments whose tracked changes have been resolved. */
  async cleanupResolvedComments(): Promise<{ deleted: number; kept: number }> {
    return cleanupResolvedComments();
  }

  /** Accepts all Stylistic tracked changes associated with a suggestion. */
  async acceptSuggestion(
    suggestion: Suggestion,
  ): Promise<SuggestionActionResult> {
    return new ResolveSuggestionCommand(
      suggestion,
      "accept",
      this.textLocator,
      this.observabilityPort,
    ).execute();
  }

  /** Rejects all Stylistic tracked changes associated with a suggestion. */
  async rejectSuggestion(
    suggestion: Suggestion,
  ): Promise<SuggestionActionResult> {
    return new ResolveSuggestionCommand(
      suggestion,
      "reject",
      this.textLocator,
      this.observabilityPort,
    ).execute();
  }

  /** Disables Track Changes only when the user explicitly requests it. */
  async disableTrackChanges(): Promise<void> {
    await this.trackChangesAdapter.disableTrackChanges();
  }

  /** Navigates to the real suggestion artifact when available. */
  async navigateToText(
    target: Suggestion | string,
  ): Promise<SuggestionNavigationResult> {
    return this.suggestionNavigationAdapter.navigateToText(target);
  }
}
