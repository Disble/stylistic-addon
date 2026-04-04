/* global Word, console */

/**
 * Word Adapter — implements `IDocumentPort` using Office.js.
 *
 * This is the ONLY file in the codebase that references the `Word` global.
 * All other modules interact with Word documents exclusively through this
 * adapter via the `IDocumentPort` interface.
 *
 * Responsibilities:
 * - Read document text (full body or active selection).
 * - Query existing Stylistic tracked changes (Guard pattern).
 * - Apply suggestions as tracked changes via `ApplySuggestionCommand`.
 * - Own workflow-level Track Changes lifecycle and document-derived pending state.
 * - Delegate comment cleanup to `CommentCleanup`.
 *
 * @module WordAdapter
 */

import type { IDocumentPort } from "../../domain/ports";
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
  SuggestionApplicationFailure,
  SuggestionApplicationFailureReason,
  TextSource,
} from "../../domain/types";
import { ApplySuggestionCommand } from "./ApplySuggestionCommand";
import {
  cleanupResolvedComments,
  getCleanupPreview,
  OVERLAPPING_RELATIONS,
} from "./cleanup/CommentCleanup";
import { isStylisticComment } from "./StylisticCommentBuilder";

const STYLISTIC_TAG_PREFIX = "stylistic:";

type ResolutionStatus =
  | "accepted"
  | "rejected"
  | "already-resolved"
  | "unobservable";

type ParagraphSnapshot = {
  text?: string;
  styleBuiltIn?: string;
  firstLineIndent?: number;
  leftIndent?: number;
};

const PARAGRAPH_LOAD_FIELDS =
  "items/text,items/styleBuiltIn,items/firstLineIndent,items/leftIndent";

const RESOLUTION_RELATED_RELATIONS = new Set([
  ...OVERLAPPING_RELATIONS,
  "AdjacentBefore",
  "AdjacentAfter",
]);

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

export class WordAdapter implements IDocumentPort {
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
   * Resolves all tracked changes semantically tied to a suggestion CC.
   *
   * `cc.getTrackedChanges()` can miss one side of a replace operation in real
   * Word documents. To avoid false-success accept/reject flows, this method
   * unions the CC-scoped tracked changes with body-level tracked changes whose
   * ranges overlap the CC range.
   */
  private async collectTrackedChangesForContentControl(
    context: Word.RequestContext,
    cc: Word.ContentControl,
  ): Promise<Word.TrackedChange[]> {
    const ccRange = cc.getRange();

    const ccTrackedChanges = cc.getTrackedChanges();
    ccTrackedChanges.load({ select: "type,id" });

    const bodyTrackedChanges = context.document.body.getTrackedChanges();
    bodyTrackedChanges.load({ select: "type,id" });

    await context.sync();

    const trackedChangesById = new Map<string, Word.TrackedChange>();
    const trackedChangesWithoutId: Word.TrackedChange[] = [];

    const addTrackedChange = (tc: Word.TrackedChange) => {
      const id = String((tc as { id?: string | number }).id ?? "");
      if (id.length > 0) {
        trackedChangesById.set(id, tc);
      } else if (!trackedChangesWithoutId.includes(tc)) {
        trackedChangesWithoutId.push(tc);
      }
    };

    for (const tc of ccTrackedChanges.items) {
      addTrackedChange(tc);
    }

    const candidateBodyTrackedChanges = bodyTrackedChanges.items.filter(
      (tc) => {
        const id = String((tc as { id?: string | number }).id ?? "");
        return id.length === 0 || !trackedChangesById.has(id);
      },
    );

    const candidateRanges = candidateBodyTrackedChanges.map((tc) =>
      tc.getRange(),
    );
    const comparisons = candidateRanges.map((range) =>
      range.compareLocationWith(ccRange),
    );

    if (comparisons.length > 0) {
      await context.sync();
    }

    for (
      let index = 0;
      index < candidateBodyTrackedChanges.length;
      index += 1
    ) {
      const tc = candidateBodyTrackedChanges[index];
      if (
        RESOLUTION_RELATED_RELATIONS.has(comparisons[index].value as string)
      ) {
        addTrackedChange(tc);
      }
    }

    return [
      ...Array.from(trackedChangesById.values()),
      ...trackedChangesWithoutId,
    ];
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
        const persistedAnchor = cc.title?.trim();
        if (persistedAnchor) {
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
   * Sorts suggestions in reverse array order as a heuristic for applying
   * end-of-document suggestions first.
   *
   * The backend returns suggestions in reading order (start → end of document).
   * Reversing ensures each applied Content Control only affects text BEFORE the
   * next search target, preventing CC boundary interference when multiple
   * suggestions target the same paragraph.
   *
   * This is a heuristic: it assumes backend output order approximates document
   * order. A future improvement could sort by actual `Range` position using
   * `Range.compareLocationWith()` inside a single `Word.run`.
   */
  private sortByDocumentPosition(suggestions: Suggestion[]): Suggestion[] {
    if (suggestions.length <= 1) return suggestions;
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

    const activated = await Word.run(async (context) =>
      this.ensureTrackChangesActive(context),
    );

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
    const command = new ApplySuggestionCommand(suggestion);

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
      console.log(`✅ [WordAdapter] "${suggestion.id}" aplicada`);
      return successCount + 1;
    }

    failedSuggestions.push({
      suggestion,
      reason: this.inferApplicationFailureReason(commandResult),
      message: commandResult.error ?? "Error desconocido al aplicar sugerencia",
    });
    console.warn(
      `⚠️ [WordAdapter] "${suggestion.id}" falló: ${commandResult.error}`,
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

  /**
   * Applies suggestions as tracked changes using `ApplySuggestionCommand`
   * (Command pattern). Each suggestion runs in its own `Word.run` context
   * (per-suggestion isolation) to avoid stale ranges after OOXML insertions.
   * Suggestions are applied in reverse order (end-of-document first) to avoid
   * Content Control boundary interference when multiple suggestions share the
   * same paragraph. Command failures are aggregated as failed suggestions so
   * later suggestions can still run.
   */
  async applySuggestions(
    suggestions: Suggestion[],
    onProgress?: ProgressCallback,
  ): Promise<ApplySuggestionsResult> {
    console.log(
      `📝 [WordAdapter] applySuggestions: ${suggestions.length} sugerencias`,
    );

    if (suggestions.length === 0) {
      const pendingAfter = await this.getDocumentReviewState();
      return {
        successCount: 0,
        failedSuggestions: [],
        pendingAfter,
        documentState: this.deriveDocumentState(pendingAfter),
        trackChangesActivatedForBatch: false,
      };
    }

    // Apply end-of-document suggestions first so earlier searches are unaffected
    // by Content Controls created for later positions in the same paragraph.
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

      this.reportApplyProgress(
        onProgress,
        successCount + failedSuggestions.length,
        suggestions.length,
      );
    }

    console.log(
      `📝 [WordAdapter] Completado: ${successCount} éxitos, ${failedSuggestions.length} fallos`,
    );

    const pendingAfter = await this.getDocumentReviewState();

    return {
      successCount,
      failedSuggestions,
      pendingAfter,
      documentState: this.deriveDocumentState(pendingAfter),
      trackChangesActivatedForBatch,
    };
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
   * Finds the Content Control for a suggestion, supporting both the new tag
   * format (`stylistic:{type}:{id}`) and the legacy bare-ID format.
   *
   * Office.js `getByTag()` is exact-match only, so we try both formats and
   * return the first CC found. Returns `null` if not found under either format.
   */
  private findCCByTag(
    context: Word.RequestContext,
    suggestion: Suggestion,
  ): {
    newFormatResult: Word.ContentControlCollection;
    legacyResult: Word.ContentControlCollection;
  } {
    const newTag = `stylistic:${suggestion.type}:${suggestion.id}`;
    const legacyTag = suggestion.id;
    const newFormatResult = context.document.contentControls.getByTag(newTag);
    const legacyResult = context.document.contentControls.getByTag(legacyTag);
    newFormatResult.load("items");
    legacyResult.load("items");
    return { newFormatResult, legacyResult };
  }

  /** Maps a resolution action to its terminal success status. */
  private toResolutionStatus(action: "accept" | "reject"): ResolutionStatus {
    return action === "accept" ? ("accepted" as const) : ("rejected" as const);
  }

  /**
   * Builds a document-aware resolution result.
   */
  private buildResolutionResult(
    status: SuggestionActionResult["status"],
    trackedChangesAffected: number,
    commentDeleted: boolean,
    pendingBefore: DocumentReviewState,
    pendingAfter: DocumentReviewState,
    error?: string,
  ): SuggestionActionResult {
    const transition = DocumentReviewStateMachine.evaluateTransition(
      pendingBefore,
      pendingAfter,
    );

    return {
      status,
      trackedChangesAffected,
      commentDeleted,
      pendingAfter,
      documentState: transition.to,
      ...(error ? { error } : {}),
    };
  }

  /** Deletes the Stylistic comment colocated with the suggestion CC range. */
  private async deleteColocatedStylisticComment(
    context: Word.RequestContext,
    cc: Word.ContentControl,
  ): Promise<boolean> {
    const comments = context.document.body.getComments();
    comments.load({ select: "authorName,content" });
    await context.sync();

    const stylisticComments = comments.items.filter(isStylisticComment);
    const ccRange = cc.getRange();

    for (const comment of stylisticComments) {
      const commentRange = comment.getRange();
      const locationResult = commentRange.compareLocationWith(ccRange);
      await context.sync();
      if (OVERLAPPING_RELATIONS.includes(locationResult.value as string)) {
        comment.delete();
        return true;
      }
    }

    return false;
  }

  /** Resolves the comment-only branch by deleting the CC and returning terminal status. */
  private async resolveCommentOnlySuggestion(
    context: Word.RequestContext,
    cc: Word.ContentControl,
    suggestion: Suggestion,
    action: "accept" | "reject",
    commentDeleted: boolean,
    pendingBefore: DocumentReviewState,
  ): Promise<SuggestionActionResult> {
    cc.delete(true);
    await context.sync();
    const pendingAfter = await this.inspectDocumentReviewState(context);
    console.log(
      `🗨️ [WordAdapter] "${suggestion.id}": comment-only ${action}ed, comentario eliminado: ${commentDeleted}`,
    );

    return this.buildResolutionResult(
      this.toResolutionStatus(action),
      0,
      commentDeleted,
      pendingBefore,
      pendingAfter,
    );
  }

  /**
   * Deletes the CC anchor after tracked changes were already resolved.
   * Reject can legitimately lose the inserted-side CC as a side effect.
   */
  private async cleanupResolvedSuggestionAnchor(
    context: Word.RequestContext,
    cc: Word.ContentControl,
    suggestion: Suggestion,
    action: "accept" | "reject",
  ): Promise<void> {
    try {
      cc.delete(true);
      await context.sync();
    } catch (cleanupError) {
      if (action === "accept") {
        throw cleanupError;
      }

      console.warn(
        `⚠️ [WordAdapter] "${suggestion.id}": reject cleanup skipped after successful resolution: ${cleanupError instanceof Error ? cleanupError.message : String(cleanupError)}`,
      );
    }
  }

  /**
   * Shared implementation for accepting or rejecting a suggestion.
   *
   * Branches on `suggestion.type`:
   * - `"track-change"`: finds Stylistic TCs inside the CC, accepts/rejects them,
   *   deletes the colocated comment, then deletes the CC.
   * - `"comment-only"`: skips TC lookup entirely. Finds and deletes the Word
   *   comment anchored to the CC range, then deletes the CC.
   *
   * Also supports legacy CCs tagged with a bare suggestion ID (no `stylistic:`
   * prefix) — those are treated as `"track-change"` for backward compatibility.
   *
   * Never throws — catches all errors and returns a result object.
   */
  private async resolveSuggestion(
    suggestion: Suggestion,
    action: "accept" | "reject",
  ): Promise<SuggestionActionResult> {
    try {
      return await Word.run(async (context) => {
        const pendingBefore = await this.inspectDocumentReviewState(context);

        // 1. Find the Content Control — try new tag format first, fall back to legacy bare ID
        const { newFormatResult, legacyResult } = this.findCCByTag(
          context,
          suggestion,
        );
        await context.sync();

        let cc: Word.ContentControl | null = null;
        if (newFormatResult.items.length > 0) {
          cc = newFormatResult.items[0];
        } else if (legacyResult.items.length > 0) {
          cc = legacyResult.items[0];
          console.log(
            `🔁 [WordAdapter] "${suggestion.id}": usando tag legado (bare ID)`,
          );
        }

        if (!cc) {
          return this.buildResolutionResult(
            "cc-not-found",
            0,
            false,
            pendingBefore,
            pendingBefore,
          );
        }

        // 2. Find and delete the colocated Stylistic comment (shared by both branches)
        const commentDeleted = await this.deleteColocatedStylisticComment(
          context,
          cc,
        );

        // 3a. comment-only branch: no TCs to process — just delete the CC
        if (suggestion.type === "comment-only") {
          return this.resolveCommentOnlySuggestion(
            context,
            cc,
            suggestion,
            action,
            commentDeleted,
            pendingBefore,
          );
        }

        // 3b. track-change branch: resolve all tracked changes semantically tied
        // to this CC. Some replace operations expose only one side through
        // `cc.getTrackedChanges()`, so we must also inspect overlapping body TCs.
        const trackedChanges =
          await this.collectTrackedChangesForContentControl(context, cc);

        if (trackedChanges.length === 0) {
          const pendingAfter = await this.inspectDocumentReviewState(context);
          return this.buildResolutionResult(
            "unobservable",
            0,
            commentDeleted,
            pendingBefore,
            pendingAfter,
            "Word no expuso suficientes tracked changes para confirmar la resolución.",
          );
        }

        for (const tc of trackedChanges) {
          if (action === "accept") {
            tc.accept();
          } else {
            tc.reject();
          }
        }

        await this.cleanupResolvedSuggestionAnchor(
          context,
          cc,
          suggestion,
          action,
        );

        const pendingAfter = await this.inspectDocumentReviewState(context);

        return this.buildResolutionResult(
          this.toResolutionStatus(action),
          trackedChanges.length,
          commentDeleted,
          pendingBefore,
          pendingAfter,
        );
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      const pendingAfter = await this.getDocumentReviewState().catch(() =>
        this.buildDocumentReviewState(0, false),
      );
      return {
        status: "error" as const,
        trackedChangesAffected: 0,
        commentDeleted: false,
        pendingAfter,
        documentState: this.deriveDocumentState(pendingAfter),
        error: message,
      };
    }
  }

  /**
   * Accepts all Stylistic tracked changes associated with a suggestion.
   * Also deletes the associated Stylistic comment if present.
   * Never throws — returns a `SuggestionActionResult`.
   */
  async acceptSuggestion(
    suggestion: Suggestion,
  ): Promise<SuggestionActionResult> {
    return this.resolveSuggestion(suggestion, "accept");
  }

  /**
   * Rejects all Stylistic tracked changes associated with a suggestion.
   * Also deletes the associated Stylistic comment if present.
   * Never throws — returns a `SuggestionActionResult`.
   */
  async rejectSuggestion(
    suggestion: Suggestion,
  ): Promise<SuggestionActionResult> {
    return this.resolveSuggestion(suggestion, "reject");
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
   * Navigates the Word document to the first occurrence of `text` by selecting
   * the matching range (Word scrolls to the selection automatically).
   * Never throws — silently no-ops when text is not found or on any error.
   */
  async navigateToText(text: string): Promise<void> {
    try {
      await Word.run(async (context) => {
        const results = context.document.body.search(text, {
          matchCase: true,
          matchWholeWord: false,
        });
        results.load("items");
        await context.sync();
        if (results.items.length > 0) {
          results.items[0].select();
          await context.sync();
        }
      });
    } catch {
      // Navigation is best-effort — silently ignore all failures
    }
  }
}
