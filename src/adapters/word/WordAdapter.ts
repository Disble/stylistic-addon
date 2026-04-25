/* global Word, console */

/**
 * Word Adapter — implements `IDocumentPort` using Office.js.
 *
 * This is the ONLY file in the codebase that references the `Word` global.
 * All other modules interact with Word documents exclusively through this
 * adapter via the `IDocumentPort` interface.
 *
 * Responsibilities (Facade — delegates heavy logic to specialized Commands):
 * - Read document text (full body or active selection).
 * - Query existing Stylistic tracked changes (Guard pattern).
 * - Apply suggestions as tracked changes via `ApplySuggestionCommand`.
 * - Resolve suggestions (accept/reject) via `ResolveSuggestionCommand`.
 * - Own workflow-level Track Changes lifecycle and document-derived pending state.
 * - Delegate comment cleanup to `CommentCleanup`.
 *
 * @module WordAdapter
 */

import type {
  IDocumentPort,
  IResolutionObservabilityPort,
} from "../../domain/ports";
import {
  DocumentReviewStateMachine,
  type DocumentReviewUiState,
} from "../../domain/review/DocumentReviewStateMachine";
import type {
  ApplySuggestionsResult,
  CommandResult,
  DocumentReviewState,
  ProgressCallback,
  Suggestion,
  SuggestionActionResult,
  TextSource,
} from "../../domain/types";
import {
  STYLISTIC_IDENTITY_TITLE_PREFIX,
  STYLISTIC_TAG_PREFIX,
} from "../../infrastructure/config";
import { NoopResolutionObservabilityAdapter } from "../observability/NoopResolutionObservabilityAdapter";
import { BatchApplyOrchestrator } from "./BatchApplyOrchestrator";
import {
  cleanupResolvedComments,
  getCleanupPreview,
} from "./cleanup/CommentCleanup";
import {
  isValidOperationalReplaceIdentity,
  parseReplaceIdentityTitle,
} from "./ReplaceIdentityParser";
import { ResolveSuggestionCommand } from "./ResolveSuggestionCommand";
import {
  getDefaultTextLocator,
  type TextLocator,
  type WordSearchContainer,
} from "./WordTextLocatorContext";

type ParagraphSnapshot = {
  text?: string;
  styleBuiltIn?: string;
  firstLineIndent?: number;
  leftIndent?: number;
};

const PARAGRAPH_LOAD_FIELDS =
  "items/text,items/styleBuiltIn,items/firstLineIndent,items/leftIndent";

function isHeadingStyle(styleBuiltIn?: string): boolean {
  return styleBuiltIn === "Title" || /^Heading\d+$/.test(styleBuiltIn ?? "");
}

function shouldPrefixIndent(paragraph: ParagraphSnapshot): boolean {
  if (isHeadingStyle(paragraph.styleBuiltIn)) {
    return false;
  }

  return (
    (paragraph.firstLineIndent ?? 0) > 0 || (paragraph.leftIndent ?? 0) > 0
  );
}

function buildStructuredParagraphText(paragraphs: ParagraphSnapshot[]): string {
  if (paragraphs.length === 0) {
    return "";
  }

  return paragraphs
    .map((paragraph) => {
      const text = paragraph.text ?? "";
      if (
        !shouldPrefixIndent(paragraph) ||
        text.length === 0 ||
        text.startsWith("\t")
      ) {
        return text;
      }

      return `\t${text}`;
    })
    .join("\n\n");
}

/** Returns the canonical Stylistic tag for one suggestion. */
function buildSuggestionTag(suggestion: Suggestion): string {
  return `${STYLISTIC_TAG_PREFIX}${suggestion.type}:${suggestion.id}`;
}

export class WordAdapter implements IDocumentPort {
  constructor(
    private readonly textLocator: TextLocator = getDefaultTextLocator(),
    private readonly observabilityPort: IResolutionObservabilityPort = new NoopResolutionObservabilityAdapter(),
  ) {}

  /**
   * Creates a normalized document-review snapshot.
   */
  private buildDocumentReviewState(
    pendingStylisticArtifacts: number,
    trackChangesActive: boolean,
  ): DocumentReviewState {
    return {
      pendingStylisticArtifacts,
      hasPendingStylisticArtifacts: pendingStylisticArtifacts > 0,
      trackChangesActive,
    };
  }

  /**
   * Derives the explicit document-review UI state from a document snapshot.
   */
  private deriveDocumentState(
    reviewState: DocumentReviewState,
  ): DocumentReviewUiState {
    return DocumentReviewStateMachine.deriveState(reviewState);
  }

  /** Chooses a strict operational-wrapper navigation CC without ranking fallback. */
  private selectNavigationContentControl(
    ccs: Word.ContentControl[],
    suggestion: Suggestion,
  ): Word.ContentControl | null {
    if (ccs.length === 0) {
      return null;
    }

    const validCandidates = ccs.filter((cc) => {
      const identity = parseReplaceIdentityTitle(cc.title);
      return isValidOperationalReplaceIdentity(identity, suggestion);
    });

    return validCandidates.length === 1 ? validCandidates[0] : null;
  }

  /** Searches a body or range through the shared Word text locator. */
  private async searchWithFallback(
    context: Word.RequestContext,
    container: WordSearchContainer,
    searchText: string,
  ): Promise<Word.Range | null> {
    return this.textLocator.locate({ context, container, searchText });
  }

  /** Re-locates a suggestion by its contextual anchor when direct artifact lookup fails. */
  private async resolveSuggestionFallbackRange(
    context: Word.RequestContext,
    suggestion: Suggestion,
  ): Promise<Word.Range | null> {
    const body = context.document.body as unknown as WordSearchContainer;
    const contextRange = await this.searchWithFallback(
      context,
      body,
      suggestion.context,
    );

    if (!contextRange) {
      return this.searchWithFallback(context, body, suggestion.anchor);
    }

    contextRange.load("text");
    const containingParagraph = contextRange.paragraphs
      .getFirst()
      .getRange("Whole");
    containingParagraph.load("text");
    await context.sync();

    const shouldExpandToParagraph =
      !contextRange.text.includes(suggestion.anchor) &&
      contextRange.text.length < suggestion.context.length - 20;

    const searchContainer = shouldExpandToParagraph
      ? (containingParagraph as unknown as WordSearchContainer)
      : (contextRange as unknown as WordSearchContainer);

    return this.searchWithFallback(context, searchContainer, suggestion.anchor);
  }

  /** Selects a range if present so Word scrolls the viewport to it. */
  private async selectRange(
    context: Word.RequestContext,
    range: Word.Range | null,
  ): Promise<void> {
    if (!range) {
      return;
    }

    range.select();
    await context.sync();
  }

  /**
   * Returns the current Track Changes activation state for the document.
   */
  private async loadTrackChangesActive(
    context: Word.RequestContext,
  ): Promise<boolean> {
    context.document.load("changeTrackingMode");
    await context.sync();
    return context.document.changeTrackingMode !== Word.ChangeTrackingMode.off;
  }

  /**
   * Reads the authoritative document-derived review state in the current batch.
   */
  private async inspectDocumentReviewState(
    context: Word.RequestContext,
  ): Promise<DocumentReviewState> {
    const allCCs = context.document.contentControls;
    allCCs.load("items/tag");
    context.document.load("changeTrackingMode");
    await context.sync();

    const pendingStylisticArtifacts = allCCs.items.filter((cc) =>
      cc.tag.startsWith(STYLISTIC_TAG_PREFIX),
    ).length;
    const trackChangesActive =
      context.document.changeTrackingMode !== Word.ChangeTrackingMode.off;

    return this.buildDocumentReviewState(
      pendingStylisticArtifacts,
      trackChangesActive,
    );
  }

  /**
   * Enables Track Changes once, lazily, before the first real track-change insertion.
   */
  private async ensureTrackChangesActive(
    context: Word.RequestContext,
  ): Promise<boolean> {
    const alreadyActive = await this.loadTrackChangesActive(context);
    if (alreadyActive) {
      return false;
    }

    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    await context.sync();
    return true;
  }

  /**
   * Resolves the text to analyze: returns the current selection if non-empty,
   * otherwise falls back to the full document body (Transparent Fallback pattern).
   */
  async getTextToAnalyze(): Promise<TextSource> {
    console.log("📖 [WordAdapter] Resolviendo texto a analizar...");
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      selection.paragraphs.load(PARAGRAPH_LOAD_FIELDS);
      await context.sync();

      const hasSelectedText = selection.text.trim().length > 0;
      const selText = hasSelectedText
        ? buildStructuredParagraphText(
            selection.paragraphs.items as ParagraphSnapshot[],
          ) || selection.text
        : "";

      if (selText && selText.trim().length > 0) {
        console.log(
          `📖 [WordAdapter] Selección activa — ${selText.length} chars`,
        );
        return { text: selText, isSelection: true };
      }

      const body = context.document.body;
      body.paragraphs.load(PARAGRAPH_LOAD_FIELDS);
      await context.sync();

      const bodyText = buildStructuredParagraphText(
        body.paragraphs.items as ParagraphSnapshot[],
      );
      if (bodyText.length > 0) {
        console.log(
          `📖 [WordAdapter] Documento completo — ${bodyText.length} chars`,
        );
        return { text: bodyText, isSelection: false };
      }

      body.load("text");
      await context.sync();
      console.log(
        `📖 [WordAdapter] Documento completo — ${body.text.length} chars`,
      );
      return { text: body.text, isSelection: false };
    });
  }

  /**
   * Returns the set of original texts already applied as Stylistic suggestions
   * (both track-change and comment-only types).
   *
   * Used as a guard to prevent duplicate tracked changes and duplicate
   * comment-only suggestions on re-run.
   *
   * Detects all active Stylistic CCs by tag prefix (`stylistic:`), which covers
   * both `stylistic:track-change:{id}` and `stylistic:comment-only:{id}` tags.
   * For each matching CC, prefers persisted anchor metadata from `cc.title`.
   * Falls back to the visible range text only for legacy CCs created before that
   * metadata existed.
   */
  async getAppliedOriginalTexts(): Promise<Set<string>> {
    console.log(
      "🛡️ [WordAdapter] Consultando CCs de Stylistic (track-change + comment-only)...",
    );
    return Word.run(async (context) => {
      const allCCs = context.document.contentControls;
      allCCs.load("items/tag,items/title");
      await context.sync();

      // JS-side prefix filter — Office.js getByTag() is exact-match only
      const stylisticCCs = allCCs.items.filter((cc) =>
        cc.tag.startsWith(STYLISTIC_TAG_PREFIX),
      );

      if (stylisticCCs.length === 0) {
        return new Set<string>();
      }

      const texts = new Set<string>();
      const legacyRanges: Word.Range[] = [];

      for (const cc of stylisticCCs) {
        const persistedIdentity = parseReplaceIdentityTitle(cc.title);
        if (persistedIdentity?.deletedSideRef?.value) {
          texts.add(persistedIdentity.deletedSideRef.value);
          continue;
        }

        if (cc.tag.startsWith("stylistic:track-change:")) {
          continue;
        }

        const persistedAnchor = cc.title?.trim();
        if (
          persistedAnchor &&
          !persistedAnchor.startsWith(STYLISTIC_IDENTITY_TITLE_PREFIX)
        ) {
          texts.add(persistedAnchor);
          continue;
        }

        const range = cc.getRange();
        range.load("text");
        legacyRanges.push(range);
      }

      if (legacyRanges.length > 0) {
        await context.sync();
        for (const range of legacyRanges) {
          texts.add(range.text);
        }
      }

      console.log(
        `🛡️ [WordAdapter] ${texts.size} texto(s) ya rastreado(s) (stylistic: CCs)`,
      );
      return texts;
    });
  }

  /**
   * Applies suggestions as tracked changes via `BatchApplyOrchestrator`.
   * Delegates sorting, Track Changes lifecycle, command execution, and progress
   * reporting to the orchestrator while providing Word-specific capabilities.
   */
  async applySuggestions(
    suggestions: Suggestion[],
    onProgress?: ProgressCallback,
  ): Promise<ApplySuggestionsResult> {
    const orchestrator = new BatchApplyOrchestrator({
      ensureTrackChangesActive: () =>
        Word.run((ctx) => this.ensureTrackChangesActive(ctx)),
      getDocumentReviewState: () => this.getDocumentReviewState(),
      deriveDocumentState: (state) => this.deriveDocumentState(state),
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
      const localizedRange = await this.searchWithFallback(
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
        (await this.searchWithFallback(
          context,
          localizedRange as unknown as WordSearchContainer,
          suggestion.anchor,
        )) ??
        (await this.searchWithFallback(
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

  /**
   * Returns the current document-derived Stylistic review state.
   */
  async getDocumentReviewState(): Promise<DocumentReviewState> {
    return Word.run((context) => this.inspectDocumentReviewState(context));
  }

  /**
   * Returns a dry-run summary of comments that can be deleted right now.
   */
  async getCleanupPreview(): Promise<{ deletable: number; kept: number }> {
    return getCleanupPreview();
  }

  /**
   * Deletes Stylistic comments whose tracked changes have been resolved.
   * Delegates to the `CommentCleanup` module (Range Colocation pattern).
   */
  async cleanupResolvedComments(): Promise<{ deleted: number; kept: number }> {
    return cleanupResolvedComments();
  }

  /**
   * Accepts all Stylistic tracked changes associated with a suggestion.
   * Delegates to `ResolveSuggestionCommand` (Command pattern).
   * Never throws — returns a `SuggestionActionResult`.
   */
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

  /**
   * Rejects all Stylistic tracked changes associated with a suggestion.
   * Delegates to `ResolveSuggestionCommand` (Command pattern).
   * Never throws — returns a `SuggestionActionResult`.
   */
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

  /**
   * Disables Track Changes only when the user explicitly requests it.
   */
  async disableTrackChanges(): Promise<void> {
    await Word.run(async (context) => {
      context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
      await context.sync();
    });
  }

  /**
   * Navigates to the real suggestion artifact when available.
   * Falls back to resilient text/context search only when direct artifact lookup
   * can no longer re-locate the suggestion.
   */
  async navigateToText(target: Suggestion | string): Promise<void> {
    try {
      await Word.run(async (context) => {
        if (typeof target !== "string") {
          const ccResult = context.document.contentControls.getByTag(
            buildSuggestionTag(target),
          );
          ccResult.load("items/tag,items/title");
          await context.sync();

          const selectedCc = this.selectNavigationContentControl(
            ccResult.items,
            target,
          );
          if (selectedCc) {
            await this.selectRange(context, selectedCc.getRange());
            return;
          }

          await this.selectRange(
            context,
            await this.resolveSuggestionFallbackRange(context, target),
          );
          return;
        }

        await this.selectRange(
          context,
          await this.searchWithFallback(
            context,
            context.document.body as unknown as WordSearchContainer,
            target,
          ),
        );
      });
    } catch {
      // Navigation is best-effort — silently ignore all failures
    }
  }
}
