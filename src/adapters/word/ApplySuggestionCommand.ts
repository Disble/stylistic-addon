/* global Word, console */

/**
 * ApplySuggestionCommand — Command pattern for tracked-change insertion.
 *
 * Encapsulates the complete logic for applying a single `Suggestion` as a
 * native Word tracked change with an embedded justification comment.
 *
 * Each command:
 * 1. Searches the document for the original text (case-sensitive).
 * 2. Calls range.insertText(suggestedText, replace) — assuming the workflow
 *    layer already enabled Track Changes when appropriate.
 * 3. Inserts a comment on the inserted range with category + justification.
 * 4. Wraps the inserted range in a ContentControl tagged `stylistic:{type}:{id}`.
 *
 * Each suggestion runs in its own `Word.run` context (per-suggestion isolation)
 * to avoid stale ranges after insertions shift document positions.
 *
 * This command intentionally does NOT own Track Changes lifecycle policy.
 * Global document review state is coordinated by the workflow/document layer.
 *
 * @module ApplySuggestionCommand
 */

import type {
  ChangeType,
  CommandResult,
  ReplaceSuggestionIdentity,
  Suggestion,
  WordArtifactRef,
} from "../../domain/types";
import {
  STYLISTIC_IDENTITY_TITLE_PREFIX,
  STYLISTIC_TAG_PREFIX,
} from "../../infrastructure/config";
import { buildStylisticCommentContent } from "./StylisticCommentBuilder";
import {
  findFirstAlphanumericOffset,
  findUniqueLocatorSubstring,
  findWhitespaceInsensitiveSlice,
} from "./WordSearchAdapter";

/** Minimal search surface shared by `Word.Body` and `Word.Range`. */
type SearchContainer = {
  search(text: string, options: Record<string, boolean>): Word.RangeCollection;
  load(property: "text"): void;
  text: string;
};

/** Returns the canonical Stylistic tag for one suggestion. */
function buildSuggestionTag(suggestion: Suggestion): string {
  return `${STYLISTIC_TAG_PREFIX}${suggestion.type}:${suggestion.id}`;
}

/** Creates one Word artifact reference owned by the apply adapter. */
function createArtifactRef(
  kind: WordArtifactRef["kind"],
  role: WordArtifactRef["role"],
  value: string,
): WordArtifactRef {
  return { kind, role, value };
}

/**
 * Serializes versioned replace identity metadata into the Content Control title.
 */
function serializeReplaceIdentity(identity: ReplaceSuggestionIdentity): string {
  return `${STYLISTIC_IDENTITY_TITLE_PREFIX}${JSON.stringify(identity)}`;
}

/**
 * Builds compound v2 metadata for replace suggestions.
 *
 * The inserted-side Content Control remains an operational reference, not the
 * whole domain identity. Deleted/original-side and anchor references are stored
 * explicitly so later observation can distinguish legacy vs v2 behavior.
 */
function buildReplaceIdentity(
  suggestion: Suggestion,
): ReplaceSuggestionIdentity {
  return {
    suggestionId: suggestion.id,
    version: "compound-v2",
    insertedSideRef: createArtifactRef(
      "content-control",
      "inserted-side",
      buildSuggestionTag(suggestion),
    ),
    deletedSideRef: createArtifactRef(
      "anchor",
      "deleted-side",
      suggestion.anchor,
    ),
    anchorRef: createArtifactRef(
      "anchor",
      "operational-anchor",
      suggestion.context,
    ),
  };
}

/**
 * Chooses the persisted Content Control title payload for one suggestion.
 */
function buildContentControlTitle(
  suggestion: Suggestion,
  changeType: ChangeType,
): string {
  if (suggestion.type === "track-change" && changeType === "replace") {
    return serializeReplaceIdentity(buildReplaceIdentity(suggestion));
  }

  return suggestion.anchor;
}

/**
 * Determines the type of tracked change operation for a suggestion.
 * Strategy pattern: selects insert / delete / replace based on text content.
 *
 * Must only be called for `"track-change"` suggestions where `suggestedText`
 * is defined. Comment-only suggestions are handled by a separate branch in
 * `execute()` and never reach this function.
 */
function classifyChange(suggestion: Suggestion): ChangeType {
  const hasOriginal = suggestion.anchor.length > 0;
  const hasSuggested = (suggestion.suggestedText?.length ?? 0) > 0;
  if (hasOriginal && !hasSuggested) return "delete";
  if (!hasOriginal && hasSuggested) return "insert";
  return "replace";
}

/**
 * Converts unknown error values into a stable, readable log message.
 */
function stringifyUnknownError(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  if (typeof error === "string") {
    return error;
  }
  try {
    return JSON.stringify(error);
  } catch {
    return Object.prototype.toString.call(error);
  }
}

/**
 * A Command that applies one `Suggestion` as a tracked change in Word.
 *
 * Implements `DocumentCommand` from `domain/types.ts`.
 */
export class ApplySuggestionCommand {
  readonly id: string;
  readonly description: string;

  constructor(private readonly suggestion: Suggestion) {
    this.id = suggestion.id;
    this.description = `Apply suggestion [${suggestion.type}]: "${suggestion.anchor.substring(0, 40)}" → "${(suggestion.suggestedText ?? "").substring(0, 40)}"`;
  }

  /**
   * Returns true when the error message indicates a Word search API rejection
   * due to an invalid or too-long search string (e.g. accented characters in
   * ignorePunct/ignoreSpace mode, or certain special-character combinations).
   */
  private static isSearchInvalidError(error: unknown): boolean {
    const message = stringifyUnknownError(error);
    return message.includes("SearchStringInvalidOrTooLong");
  }

  /**
   * Runs one Word search attempt and normalizes SearchStringInvalidOrTooLong.
   */
  private async runSearchAttempt(
    context: Word.RequestContext,
    container: SearchContainer,
    searchText: string,
    options: Record<string, boolean>,
  ): Promise<{ invalid: boolean; results?: Word.RangeCollection }> {
    try {
      const results = container.search(searchText, options);
      results.load("items");
      await context.sync();
      return { invalid: false, results };
    } catch (error) {
      if (ApplySuggestionCommand.isSearchInvalidError(error)) {
        return { invalid: true };
      }
      throw error;
    }
  }

  /**
   * Resolves a range by scanning text with whitespace-insensitive matching.
   *
   * When `findUniqueLocatorSubstring` returns a candidate that starts with
   * special characters (em-dashes, inverted question marks, etc.), Word may
   * reject it with `SearchStringInvalidOrTooLong`. In that case, a second
   * attempt is made with a sub-slice of `rawSlice` that starts at the first
   * alphanumeric character, sidestepping the problematic leading chars while
   * still uniquely identifying the same range in the document.
   */
  private async resolveByWhitespaceScan(
    context: Word.RequestContext,
    container: SearchContainer,
    searchText: string,
    searchOptions: Record<string, boolean>,
  ): Promise<Word.Range | null> {
    container.load("text");
    await context.sync();

    const rawSlice = findWhitespaceInsensitiveSlice(searchText, container.text);
    if (!rawSlice) {
      return null;
    }

    const fallbackSearchText = findUniqueLocatorSubstring(
      rawSlice,
      container.text,
    );
    if (!fallbackSearchText) {
      return null;
    }

    const fallbackAttempt = await this.runSearchAttempt(
      context,
      container,
      fallbackSearchText,
      searchOptions,
    );
    if (!fallbackAttempt.invalid) {
      return fallbackAttempt.results?.items[0] ?? null;
    }

    // The candidate was rejected by Word (SearchStringInvalidOrTooLong).
    // This often happens when the rawSlice starts with special characters such
    // as em-dashes (—) or inverted question marks (¿). Retry from the first
    // alphanumeric character in rawSlice to produce a Word-safe search string.
    const alphanumericOffset = findFirstAlphanumericOffset(rawSlice);
    if (alphanumericOffset <= 0) {
      return null;
    }

    const alphanumericSlice = rawSlice.slice(alphanumericOffset);
    const alphanumericCandidate = findUniqueLocatorSubstring(
      alphanumericSlice,
      container.text,
    );
    if (!alphanumericCandidate) {
      return null;
    }

    const retryAttempt = await this.runSearchAttempt(
      context,
      container,
      alphanumericCandidate,
      searchOptions,
    );
    if (retryAttempt.invalid) {
      return null;
    }

    return retryAttempt.results?.items[0] ?? null;
  }

  /**
   * Searches a body or range using the standard three-step fallback strategy.
   *
   * Step 1 (exact, case-sensitive): direct match — skipped when text > 256 chars.
   * Step 2 (ignorePunct + ignoreSpace): tolerates minor punctuation/spacing differences.
   * Step 3 (whitespace-insensitive text scan): reads `.text`, strips whitespace and
   *         normalizes cross-source chars (e.g. smart quotes) from both the needle and
   *         the haystack, locates the match, and searches the resulting slice. When the
   *         slice is longer than Word's 256-char search limit, it is truncated to the
   *         first 256 characters so the `search()` call can succeed.
   *
   * If any `container.search()` call throws `SearchStringInvalidOrTooLong` (Word
   * rejects strings with certain accented or special characters in some search modes),
   * the method skips directly to step 3 so the suggestion is still applied.
   */
  private async searchWithFallback(
    context: Word.RequestContext,
    container: SearchContainer,
    searchText: string,
  ): Promise<Word.Range | null> {
    const searchOptions = { matchCase: true, matchWholeWord: false };
    const relaxedOptions = {
      matchCase: true,
      matchWholeWord: false,
      ignorePunct: true,
      ignoreSpace: true,
    };

    const exactMatchAllowed = searchText.length <= 256;
    let exactResults: Word.RangeCollection | undefined;

    if (exactMatchAllowed) {
      const exactAttempt = await this.runSearchAttempt(
        context,
        container,
        searchText,
        searchOptions,
      );
      if (exactAttempt.invalid) {
        return this.resolveByWhitespaceScan(
          context,
          container,
          searchText,
          searchOptions,
        );
      }

      exactResults = exactAttempt.results;
      if ((exactResults?.items.length ?? 0) > 0) {
        return exactResults?.items[0] ?? null;
      }
    }

    const shouldRunRelaxedSearch =
      !exactMatchAllowed || (exactResults?.items.length ?? 0) === 0;
    if (shouldRunRelaxedSearch) {
      const relaxedAttempt = await this.runSearchAttempt(
        context,
        container,
        searchText,
        relaxedOptions,
      );
      if (relaxedAttempt.invalid) {
        return this.resolveByWhitespaceScan(
          context,
          container,
          searchText,
          searchOptions,
        );
      }

      if ((relaxedAttempt.results?.items.length ?? 0) > 0) {
        return relaxedAttempt.results?.items[0] ?? null;
      }
    }

    // Step 3: whitespace-insensitive text scan — activates when all search()
    // attempts complete with zero results.
    return this.resolveByWhitespaceScan(
      context,
      container,
      searchText,
      searchOptions,
    );
  }

  /**
   * Resolves the exact anchor range by first locating the surrounding context,
   * then searching the anchor within that context range.
   */
  private async resolveAnchorRange(
    context: Word.RequestContext,
    body: Word.Body,
  ): Promise<Word.Range | null> {
    const contextRange = await this.searchWithFallback(
      context,
      body as SearchContainer,
      this.suggestion.context,
    );
    if (!contextRange) {
      return null;
    }

    contextRange.load("text");
    const containingParagraph = contextRange.paragraphs
      .getFirst()
      .getRange("Whole");
    containingParagraph.load("text");
    await context.sync();

    const matchText = contextRange.text;
    console.log(
      `🔬 [ApplySuggestionCommand] "${this.id}": contextMatchLen=${matchText.length}, paragraphLen=${containingParagraph.text.length}, anchorIndexInMatch=${matchText.indexOf(this.suggestion.anchor)}, anchorIndexInParagraph=${containingParagraph.text.indexOf(this.suggestion.anchor)}`,
    );

    const shouldExpandToParagraph =
      !matchText.includes(this.suggestion.anchor) &&
      matchText.length < this.suggestion.context.length - 20;

    const searchContainer = shouldExpandToParagraph
      ? (containingParagraph as unknown as SearchContainer)
      : (contextRange as unknown as SearchContainer);

    if (shouldExpandToParagraph) {
      console.log(
        `🔬 [ApplySuggestionCommand] "${this.id}": context match (${matchText.length} chars) does not contain anchor — expanding to paragraph (${containingParagraph.text.length} chars)`,
      );
    }

    return this.searchWithFallback(
      context,
      searchContainer,
      this.suggestion.anchor,
    );
  }

  /**
   * Executes the command: searches for the anchor within its context and
   * replaces it with a native Word tracked change.
   *
   * Returns `{ success: false }` for recoverable application failures such as
   * missing anchor text, missing search matches, or Office insertion errors.
   */
  async execute(): Promise<CommandResult> {
    console.log(
      `🔍 [ApplySuggestionCommand] "${this.id}": "${this.suggestion.anchor.substring(0, 40)}" → "${(this.suggestion.suggestedText ?? "").substring(0, 40)}"`,
    );

    if (this.suggestion.type === "comment-only") {
      return this.executeCommentOnly();
    }

    const changeType = classifyChange(this.suggestion);

    if (changeType === "insert") {
      console.warn(
        `⚠️ [ApplySuggestionCommand] "${this.id}": inserción sin texto ancla`,
      );
      return {
        success: false,
        commandId: this.id,
        error: "Insert-only suggestions require anchor text",
      };
    }

    try {
      return await Word.run(async (context) => {
        const body = context.document.body;
        let range = await this.resolveAnchorRange(context, body);
        if (!range) {
          console.warn(
            `🔍 [ApplySuggestionCommand] "${this.id}": anchor no encontrado`,
          );
          return {
            success: false,
            commandId: this.id,
            error: "Anchor no encontrado en el contexto",
          };
        }

        const parentCC = range.parentContentControlOrNullObject;
        parentCC.load("tag");
        await context.sync();

        const existingTag = parentCC.isNullObject ? "" : parentCC.tag;
        const isAlreadyCovered =
          existingTag.startsWith("stylistic:") ||
          /^chunk\d+-\d+$/.test(existingTag);

        if (isAlreadyCovered) {
          console.log(
            `♻️ [ApplySuggestionCommand] "${this.id}": CC existente detectado — eliminando wrapper y reinsertando`,
          );
          // keepContent: true = remove CC wrapper only, text stays in document
          // keepContent: false = remove CC wrapper AND DELETE its content ← DANGER
          parentCC.delete(true);
          await context.sync();

          range = await this.resolveAnchorRange(context, body);
          if (!range) {
            console.warn(
              `🔍 [ApplySuggestionCommand] "${this.id}": anchor no encontrado tras eliminar CC existente`,
            );
            return {
              success: false,
              commandId: this.id,
              error: "Anchor no encontrado en el contexto",
            };
          }
        }

        console.log(
          `📄 [ApplySuggestionCommand] "${this.id}": insertando TC nativo (tipo: ${changeType})`,
        );

        const insertedRange = range.insertText(
          this.suggestion.suggestedText ?? "",
          Word.InsertLocation.replace,
        );

        insertedRange.insertComment(
          buildStylisticCommentContent(
            this.suggestion.category,
            this.suggestion.justification,
          ),
        );

        const cc = insertedRange.insertContentControl();
        cc.tag = buildSuggestionTag(this.suggestion);
        cc.title = buildContentControlTitle(this.suggestion, changeType);
        cc.appearance = "Hidden";
        cc.cannotDelete = false;
        await context.sync();

        console.log(
          `✅ [ApplySuggestionCommand] "${this.id}": insertado exitosamente`,
        );

        return { success: true, commandId: this.id };
      });
    } catch (error) {
      const message = stringifyUnknownError(error);
      console.warn(`⚠️ [ApplySuggestionCommand] "${this.id}": ${message}`);
      return { success: false, commandId: this.id, error: message };
    }
  }

  /**
   * Executes the comment-only path: locates the anchor within its context and
   * inserts a Word comment at that range with NO tracked change markup.
   *
   * changeTrackingMode is NOT read or modified in this path.
   *
   * The search + fallback logic mirrors the track-change path so both paths
   * share the same whitespace-insensitive anchor resolution behaviour.
   */
  private async executeCommentOnly(): Promise<CommandResult> {
    try {
      return await Word.run(async (context) => {
        const body = context.document.body;
        let range = await this.resolveAnchorRange(context, body);
        if (!range) {
          console.warn(
            `🔍 [ApplySuggestionCommand] "${this.id}": anchor no encontrado (comment-only)`,
          );
          return {
            success: false,
            commandId: this.id,
            error: "Anchor no encontrado en el contexto",
          };
        }

        const parentCC = range.parentContentControlOrNullObject;
        parentCC.load("tag");
        await context.sync();

        const existingTag = parentCC.isNullObject ? "" : parentCC.tag;
        const isAlreadyCovered =
          existingTag.startsWith("stylistic:") ||
          /^chunk\d+-\d+$/.test(existingTag);

        if (isAlreadyCovered) {
          console.log(
            `♻️ [ApplySuggestionCommand] "${this.id}": CC existente detectado (comment-only) — eliminando wrapper y reinsertando`,
          );
          parentCC.delete(true);
          await context.sync();

          range = await this.resolveAnchorRange(context, body);
          if (!range) {
            console.warn(
              `🔍 [ApplySuggestionCommand] "${this.id}": anchor no encontrado tras eliminar CC existente (comment-only)`,
            );
            return {
              success: false,
              commandId: this.id,
              error: "Anchor no encontrado en el contexto",
            };
          }
        }

        console.log(
          `📄 [ApplySuggestionCommand] "${this.id}": insertando comment-only nativo`,
        );

        range.insertComment(
          buildStylisticCommentContent(
            this.suggestion.category,
            this.suggestion.justification,
          ),
        );

        const cc = range.insertContentControl();
        cc.tag = buildSuggestionTag(this.suggestion);
        cc.title = this.suggestion.anchor;
        cc.appearance = "Hidden";
        cc.cannotDelete = false;
        await context.sync();

        console.log(
          `✅ [ApplySuggestionCommand] "${this.id}": comment-only insertado exitosamente`,
        );

        return { success: true, commandId: this.id };
      });
    } catch (error) {
      const message = stringifyUnknownError(error);
      console.warn(`⚠️ [ApplySuggestionCommand] "${this.id}": ${message}`);
      return { success: false, commandId: this.id, error: message };
    }
  }
}
