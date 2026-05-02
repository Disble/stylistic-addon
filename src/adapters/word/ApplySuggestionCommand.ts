/* global Word, console */

/**
 * ApplySuggestionCommand — Command pattern for tracked-change insertion.
 *
 * Encapsulates the complete logic for applying a single `Suggestion` as a
 * native Word tracked change with an embedded justification comment.
 *
 * Each command:
 * 1. Searches the document for the original text (case-sensitive).
 * 2. For replace suggestions, creates or reuses an operational wrapper without
 *    changing Track Changes state.
 * 3. Calls range.insertText(suggestedText, replace) — assuming the workflow
 *    layer already enabled Track Changes when appropriate.
 * 4. Re-locates a clean current-side range for replace suggestions when Word
 *    returns a hybrid tracked-change span.
 * 5. Inserts a comment on the current inserted side with category + justification.
 * 6. Wraps that current inserted side in a ContentControl tagged
 *    `stylistic:{type}:{id}`.
 *
 * Each suggestion runs in its own `Word.run` context (per-suggestion isolation)
 * to avoid stale ranges after insertions shift document positions.
 *
 * This command intentionally does NOT own Track Changes lifecycle policy.
 * Global document review state is coordinated by the workflow/document layer;
 * apply-time wrapper creation must not temporarily disable Track Changes.
 *
 * @module ApplySuggestionCommand
 */

import type { CommandResult } from "../../domain/DocumentApplication.types";
import type { Suggestion } from "../../domain/suggestion/Suggestion.types";
import { STYLISTIC_TAG_PREFIX } from "../../infrastructure/config";
import { applySuggestionObservability } from "../observability/ConsoleApplySuggestionObservabilityAdapter";
import { ApplySuggestionAnchorResolver } from "./apply-suggestion/ApplySuggestionAnchorResolver";
import { ApplySuggestionIdentityBuilder } from "./apply-suggestion/ApplySuggestionIdentityBuilder";
import { ApplySuggestionMutationPatchBuilder } from "./apply-suggestion/ApplySuggestionMutationPatchBuilder";
import { ApplySuggestionOperationalWrapperResolver } from "./apply-suggestion/ApplySuggestionOperationalWrapperResolver";
import { ApplySuggestionReplaceRangeResolver } from "./apply-suggestion/ApplySuggestionReplaceRangeResolver";
import {
  type TrackChangeSubtypeResolution,
  TrackChangeSubtypeResolver,
} from "./apply-suggestion/TrackChangeSubtypeResolver";
import { buildStylisticCommentContent } from "./StylisticCommentBuilder";
import type { TextLocator } from "./WordTextLocatorContext";

/**
 * A Command that applies one `Suggestion` as a tracked change in Word.
 *
 * Implements `DocumentCommand` from `domain/types.ts`.
 */
export class ApplySuggestionCommand {
  readonly id: string;
  readonly description: string;
  private readonly identityBuilder: ApplySuggestionIdentityBuilder;
  private readonly mutationPatchBuilder: ApplySuggestionMutationPatchBuilder;
  private readonly anchorResolver: ApplySuggestionAnchorResolver;
  private readonly replaceRangeResolver: ApplySuggestionReplaceRangeResolver;
  private readonly operationalWrapperResolver: ApplySuggestionOperationalWrapperResolver;
  private readonly subtypeResolver: TrackChangeSubtypeResolver;

  constructor(
    private readonly suggestion: Suggestion,
    private readonly textLocator: TextLocator,
  ) {
    this.id = suggestion.id;
    this.description = `Apply suggestion [${suggestion.type}]: "${suggestion.anchor.substring(0, 40)}" → "${(suggestion.suggestedText ?? "").substring(0, 40)}"`;
    this.identityBuilder = new ApplySuggestionIdentityBuilder();
    this.mutationPatchBuilder = new ApplySuggestionMutationPatchBuilder();
    this.subtypeResolver = new TrackChangeSubtypeResolver();
    this.anchorResolver = new ApplySuggestionAnchorResolver(
      suggestion,
      this.textLocator,
      this.id,
    );
    this.replaceRangeResolver = new ApplySuggestionReplaceRangeResolver(
      suggestion,
      this.textLocator,
      this.id,
    );
    this.operationalWrapperResolver =
      new ApplySuggestionOperationalWrapperResolver(
        suggestion,
        this.textLocator,
        this.identityBuilder,
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

    const changeType = this.mutationPatchBuilder.classifyChange(
      this.suggestion,
    );
    const subtypeResolution = this.subtypeResolver.resolve(this.suggestion);

    if (subtypeResolution.subtype === "insert" || changeType === "insert") {
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
        const range = await this.anchorResolver.resolveAnchorRange(
          context,
          body,
        );
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

        const operationalWrapperResult =
          await this.operationalWrapperResolver.resolveOperationalWrapper(
            context,
            range,
          );

        if (operationalWrapperResult?.error) {
          console.log(
            `♻️ [ApplySuggestionCommand] "${this.id}": CC existente detectado — covered-by-existing-cc abort before mutation`,
          );
          return {
            success: false,
            commandId: this.id,
            error: operationalWrapperResult.error,
          };
        }

        const containingParagraph = range.paragraphs
          .getFirst()
          .getRange("Whole");
        containingParagraph.load("text");
        const shouldReadTrackingMode = subtypeResolution.subtype === "replace";
        if (shouldReadTrackingMode) {
          context.document.load("changeTrackingMode");
        }
        await context.sync();
        const preMutationTrackingMode = shouldReadTrackingMode
          ? context.document.changeTrackingMode
          : "not-read";

        console.log(
          `📄 [ApplySuggestionCommand] "${this.id}": insertando TC nativo (tipo: ${changeType})`,
        );
        applySuggestionObservability.logPreMutationScope(this.id, {
          changeType: subtypeResolution.subtype,
          changeTrackingMode: preMutationTrackingMode,
          anchor: this.suggestion.anchor,
          suggestedText: this.suggestion.suggestedText ?? "",
          paragraphLength: containingParagraph.text.length,
          paragraphPreview: containingParagraph.text.substring(0, 180),
        });

        const operationalWrapper = operationalWrapperResult.wrapper ?? null;
        const mutationRange = operationalWrapper
          ? await this.operationalWrapperResolver.resolveAnchorInsideWrapper(
              context,
              operationalWrapper,
            )
          : range;

        if (!mutationRange) {
          return {
            success: false,
            commandId: this.id,
            error:
              "No se pudo re-localizar el anchor dentro del wrapper operacional",
          };
        }

        applySuggestionObservability.logMutationTargetResolved(this.id, {
          usesOperationalWrapper: operationalWrapper !== null,
          insertLocation: "replace",
        });

        const annotationRange = await this.applyNativeTrackChangeMutation(
          context,
          mutationRange,
          operationalWrapper!.getRange(),
          subtypeResolution,
        );

        if (!annotationRange) {
          applySuggestionObservability.warnPostMutationIsolationFailure(
            this.id,
          );
          return {
            success: false,
            commandId: this.id,
            error: "No se pudo aislar el texto insertado de la sugerencia",
          };
        }

        annotationRange.insertComment(
          buildStylisticCommentContent(
            this.suggestion.category,
            this.suggestion.justification,
          ),
        );

        const cc = annotationRange.insertContentControl();
        cc.tag = this.identityBuilder.buildSuggestionTag(this.suggestion);
        cc.title = this.identityBuilder.buildContentControlTitle(
          this.suggestion,
        );
        cc.appearance = "Hidden";
        cc.cannotDelete = false;
        await context.sync();

        console.log(
          `✅ [ApplySuggestionCommand] "${this.id}": insertado exitosamente`,
        );

        const mutationPatch =
          subtypeResolution.subtype === "formatting"
            ? undefined
            : this.mutationPatchBuilder.buildApplyMutationPatch(
                this.suggestion,
                containingParagraph.text,
              );

        return {
          success: true,
          commandId: this.id,
          mutationPatch,
        };
      });
    } catch (error) {
      const message = this.mutationPatchBuilder.stringifyUnknownError(error);
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
        let range = await this.anchorResolver.resolveAnchorRange(context, body);
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

          range = await this.anchorResolver.resolveAnchorRange(context, body);
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
        cc.tag = this.identityBuilder.buildSuggestionTag(this.suggestion);
        cc.title = this.identityBuilder.buildContentControlTitle(
          this.suggestion,
        );
        cc.appearance = "Hidden";
        cc.cannotDelete = false;
        await context.sync();

        console.log(
          `✅ [ApplySuggestionCommand] "${this.id}": comment-only insertado exitosamente`,
        );

        return { success: true, commandId: this.id };
      });
    } catch (error) {
      const message = this.mutationPatchBuilder.stringifyUnknownError(error);
      console.warn(`⚠️ [ApplySuggestionCommand] "${this.id}": ${message}`);
      return { success: false, commandId: this.id, error: message };
    }
  }

  /**
   * Applies the concrete native Word mutation for the resolved subtype and
   * returns the range that must receive comment + metadata annotation.
   */
  private async applyNativeTrackChangeMutation(
    context: Word.RequestContext,
    mutationRange: Word.Range,
    wrapperRange: Word.Range,
    subtypeResolution: TrackChangeSubtypeResolution,
  ): Promise<Word.Range | null> {
    if (subtypeResolution.subtype === "formatting") {
      if (!subtypeResolution.formatting) {
        return null;
      }

      if (subtypeResolution.formatting.kind === "italic") {
        mutationRange.font.italic = true;
      } else {
        mutationRange.font.bold = true;
      }
      await context.sync();
      return mutationRange;
    }

    const insertedRange = mutationRange.insertText(
      this.suggestion.suggestedText ?? "",
      Word.InsertLocation.replace,
    );
    applySuggestionObservability.logInsertTextIssued(this.id);

    if (subtypeResolution.subtype === "delete-only") {
      return wrapperRange;
    }

    return this.replaceRangeResolver.resolveReplaceAnnotationRange(
      context,
      insertedRange,
      wrapperRange,
    );
  }
}
