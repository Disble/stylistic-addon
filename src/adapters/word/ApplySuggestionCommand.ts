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
 * 3. Re-locates a clean current-side range for replace suggestions when Word
 *    returns a hybrid tracked-change span.
 * 4. Inserts a comment on the current inserted side with category + justification.
 * 5. Wraps that current inserted side in a ContentControl tagged
 *    `stylistic:{type}:{id}`.
 *
 * Each suggestion runs in its own `Word.run` context (per-suggestion isolation)
 * to avoid stale ranges after insertions shift document positions.
 *
 * This command intentionally does NOT own Track Changes lifecycle policy.
 * Global document review state is coordinated by the workflow/document layer.
 *
 * @module ApplySuggestionCommand
 */

import type { CommandResult } from "../../domain/DocumentApplication.types";
import type { Suggestion } from "../../domain/suggestion/Suggestion.types";
import { STYLISTIC_TAG_PREFIX } from "../../infrastructure/config";
import { ApplySuggestionAnchorResolver } from "./apply-suggestion/ApplySuggestionAnchorResolver";
import { ApplySuggestionIdentityBuilder } from "./apply-suggestion/ApplySuggestionIdentityBuilder";
import { ApplySuggestionMutationPatchBuilder } from "./apply-suggestion/ApplySuggestionMutationPatchBuilder";
import { ApplySuggestionOperationalWrapperResolver } from "./apply-suggestion/ApplySuggestionOperationalWrapperResolver";
import { ApplySuggestionReplaceRangeResolver } from "./apply-suggestion/ApplySuggestionReplaceRangeResolver";
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

  constructor(
    private readonly suggestion: Suggestion,
    private readonly textLocator: TextLocator,
  ) {
    this.id = suggestion.id;
    this.description = `Apply suggestion [${suggestion.type}]: "${suggestion.anchor.substring(0, 40)}" → "${(suggestion.suggestedText ?? "").substring(0, 40)}"`;
    this.identityBuilder = new ApplySuggestionIdentityBuilder();
    this.mutationPatchBuilder = new ApplySuggestionMutationPatchBuilder();
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
          changeType === "replace"
            ? await this.operationalWrapperResolver.resolveOperationalWrapper(
                context,
                range,
              )
            : null;

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

        if (changeType !== "replace") {
          const parentCC = range.parentContentControlOrNullObject;
          parentCC.load("tag");
          await context.sync();

          const existingTag = parentCC.isNullObject ? "" : parentCC.tag;
          const isAlreadyCovered =
            existingTag.startsWith(STYLISTIC_TAG_PREFIX) ||
            /^chunk\d+-\d+$/.test(existingTag);

          if (isAlreadyCovered) {
            console.log(
              `♻️ [ApplySuggestionCommand] "${this.id}": CC existente detectado — covered-by-existing-cc abort before mutation`,
            );
            return {
              success: false,
              commandId: this.id,
              error: "Anchor cubierto por un Content Control existente",
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

        const operationalWrapper =
          changeType === "replace"
            ? (operationalWrapperResult?.wrapper ?? null)
            : null;
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

        const insertedRange = mutationRange.insertText(
          this.suggestion.suggestedText ?? "",
          Word.InsertLocation.replace,
        );

        const annotationRange =
          changeType === "replace"
            ? await this.replaceRangeResolver.resolveReplaceAnnotationRange(
                context,
                insertedRange,
                operationalWrapper!.getRange(),
              )
            : insertedRange;

        if (!annotationRange) {
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

        return {
          success: true,
          commandId: this.id,
          mutationPatch: this.mutationPatchBuilder.buildApplyMutationPatch(
            this.suggestion,
            containingParagraph.text,
          ),
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
}
