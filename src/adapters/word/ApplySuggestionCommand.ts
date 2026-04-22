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
  ApplyMutationPatch,
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
import type {
  TextLocator,
  WordSearchContainer,
} from "./WordTextLocatorContext";

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

/** Detects Word's invalid/too-long search failures. */
function isSearchInvalidError(error: unknown): boolean {
  return stringifyUnknownError(error).includes("SearchStringInvalidOrTooLong");
}

/** Normalizes one character for local paragraph-context scoring. */
function normalizeContextScoreChar(char: string): string {
  if (char === "\u201C" || char === "\u201D") {
    return '"';
  }

  if (char === "\u2018" || char === "\u2019") {
    return "'";
  }

  if (char >= "\u0013" && char <= "\u0015") {
    return "";
  }

  return char.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

/** Normalizes text for paragraph-context similarity scoring. */
function normalizeForContextScore(text: string): string {
  return Array.from(text.toLowerCase())
    .map((char) => normalizeContextScoreChar(char))
    .join("")
    .replace(/[^\p{L}\p{N}\s]/gu, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/** Scores how well one paragraph aligns with the original backend context. */
function scoreParagraphAgainstContext(
  paragraphText: string,
  contextText: string,
): number {
  const paragraphTokens = new Set(
    normalizeForContextScore(paragraphText)
      .split(" ")
      .filter((token) => token.length >= 3),
  );
  const contextTokens = normalizeForContextScore(contextText)
    .split(" ")
    .filter((token) => token.length >= 3);

  return contextTokens.reduce(
    (score, token) => score + (paragraphTokens.has(token) ? 1 : 0),
    0,
  );
}

/** Builds a localized mutation patch from one successful anchor replacement. */
function buildApplyMutationPatch(
  suggestion: Suggestion,
  containerText: string,
): ApplyMutationPatch | undefined {
  const replacement = suggestion.suggestedText ?? "";
  const affectedStart = containerText.indexOf(suggestion.anchor);

  if (affectedStart < 0) {
    return undefined;
  }

  const affectedEnd = affectedStart + suggestion.anchor.length;

  return {
    suggestionId: suggestion.id,
    snapshotVersion: (suggestion.positionHint?.snapshotVersion ?? 0) + 1,
    paragraphId: suggestion.positionHint?.paragraphId,
    originalText: containerText,
    updatedText:
      containerText.slice(0, affectedStart) +
      replacement +
      containerText.slice(affectedEnd),
    deltaLength: replacement.length - suggestion.anchor.length,
    affectedStart,
    affectedEnd,
  };
}

/**
 * A Command that applies one `Suggestion` as a tracked change in Word.
 *
 * Implements `DocumentCommand` from `domain/types.ts`.
 */
export class ApplySuggestionCommand {
  readonly id: string;
  readonly description: string;

  constructor(
    private readonly suggestion: Suggestion,
    private readonly textLocator: TextLocator,
  ) {
    this.id = suggestion.id;
    this.description = `Apply suggestion [${suggestion.type}]: "${suggestion.anchor.substring(0, 40)}" → "${(suggestion.suggestedText ?? "").substring(0, 40)}"`;
  }

  /**
   * Resolves the exact anchor range by first locating the surrounding context,
   * then searching the anchor within that context range.
   */
  private async resolveAnchorRange(
    context: Word.RequestContext,
    body: Word.Body,
  ): Promise<Word.Range | null> {
    const contextRange = await this.textLocator.locate({
      context,
      container: body as WordSearchContainer,
      searchText: this.suggestion.context,
    });
    if (!contextRange) {
      console.log(
        `🔬 [ApplySuggestionCommand] "${this.id}": context not found — retrying anchor search in full body`,
      );
      return this.resolveBestBodyAnchorFallback(context, body);
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
    const shouldRetryInParagraphAfterMiss =
      !shouldExpandToParagraph &&
      matchText.length < this.suggestion.context.length - 20 &&
      containingParagraph.text.length > matchText.length;

    const searchContainer = shouldExpandToParagraph
      ? (containingParagraph as unknown as WordSearchContainer)
      : (contextRange as unknown as WordSearchContainer);

    if (shouldExpandToParagraph) {
      console.log(
        `🔬 [ApplySuggestionCommand] "${this.id}": context match (${matchText.length} chars) does not contain anchor — expanding to paragraph (${containingParagraph.text.length} chars)`,
      );
    }

    const anchorRange = await this.textLocator.locate({
      context,
      container: searchContainer,
      searchText: this.suggestion.anchor,
    });

    if (anchorRange || !shouldRetryInParagraphAfterMiss) {
      return anchorRange;
    }

    console.log(
      `🔬 [ApplySuggestionCommand] "${this.id}": anchor not found inside partial context match (${matchText.length} chars) — retrying in paragraph (${containingParagraph.text.length} chars)`,
    );

    return this.textLocator.locate({
      context,
      container: containingParagraph as unknown as WordSearchContainer,
      searchText: this.suggestion.anchor,
    });
  }

  /** Selects the best full-body anchor candidate when context re-location fails. */
  private async resolveBestBodyAnchorFallback(
    context: Word.RequestContext,
    body: Word.Body,
  ): Promise<Word.Range | null> {
    const searchOptions = { matchCase: true, matchWholeWord: false };
    const relaxedOptions = {
      matchCase: true,
      matchWholeWord: false,
      ignorePunct: true,
      ignoreSpace: true,
    };

    try {
      const exactResults = body.search(this.suggestion.anchor, searchOptions);
      exactResults.load("items");
      await context.sync();

      if (exactResults.items.length === 1) {
        return exactResults.items[0] ?? null;
      }

      if (exactResults.items.length > 1) {
        return this.selectBestAnchorCandidateByContext(
          context,
          exactResults.items,
        );
      }

      const relaxedResults = body.search(
        this.suggestion.anchor,
        relaxedOptions,
      );
      relaxedResults.load("items");
      await context.sync();

      if (relaxedResults.items.length === 1) {
        return relaxedResults.items[0] ?? null;
      }

      if (relaxedResults.items.length > 1) {
        return this.selectBestAnchorCandidateByContext(
          context,
          relaxedResults.items,
        );
      }
    } catch (error) {
      if (!isSearchInvalidError(error)) {
        throw error;
      }
    }

    return this.textLocator.locate({
      context,
      container: body as WordSearchContainer,
      searchText: this.suggestion.anchor,
    });
  }

  /** Chooses the anchor whose paragraph best matches the stale backend context. */
  private async selectBestAnchorCandidateByContext(
    context: Word.RequestContext,
    candidates: Word.Range[],
  ): Promise<Word.Range | null> {
    const paragraphs = candidates.map((candidate) => ({
      candidate,
      paragraph: candidate.paragraphs.getFirst().getRange("Whole"),
    }));

    for (const entry of paragraphs) {
      entry.paragraph.load("text");
    }
    await context.sync();

    const scored = paragraphs
      .map(({ candidate, paragraph }) => ({
        candidate,
        score: scoreParagraphAgainstContext(
          paragraph.text,
          this.suggestion.context,
        ),
      }))
      .sort((left, right) => right.score - left.score);

    const best = scored[0];
    const secondBest = scored[1];

    if (!best || best.score <= 0) {
      console.log(
        `🔬 [ApplySuggestionCommand] "${this.id}": ambiguous full-body anchor fallback without contextual winner`,
      );
      return null;
    }

    if (secondBest && secondBest.score === best.score) {
      console.log(
        `🔬 [ApplySuggestionCommand] "${this.id}": ambiguous full-body anchor fallback tie (${best.score})`,
      );
      return null;
    }

    console.log(
      `🔬 [ApplySuggestionCommand] "${this.id}": disambiguated full-body fallback with contextual score ${best.score}`,
    );

    return best.candidate;
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

        const containingParagraph = range.paragraphs
          .getFirst()
          .getRange("Whole");
        containingParagraph.load("text");
        await context.sync();

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

        return {
          success: true,
          commandId: this.id,
          mutationPatch: buildApplyMutationPatch(
            this.suggestion,
            containingParagraph.text,
          ),
        };
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
