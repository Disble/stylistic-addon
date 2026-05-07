/* global Word */

import type { Suggestion } from "../../../domain/suggestion/Suggestion.types";
import {
  STYLISTIC_OPERATIONAL_WRAPPER_TAG_PREFIX,
  STYLISTIC_TAG_PREFIX,
} from "../../../infrastructure/config";
import { applySuggestionObservability } from "../../observability/ConsoleApplySuggestionObservabilityAdapter";
import {
  isValidOperationalReplaceIdentity,
  parseReplaceIdentityTitle,
} from "../ReplaceIdentityParser";
import type { TextLocator, WordSearchContainer } from "../WordTextLocatorContext.types";
import type { ApplySuggestionIdentityBuilder } from "./ApplySuggestionIdentityBuilder";
import type {
  ApplySuggestionOperationalWrapperResolution,
  ParentOperationalContentControl,
} from "./ApplySuggestionOperationalWrapperResolver.types";

/**
 * Owns replace-suggestion operational-wrapper creation, reuse, and in-wrapper
 * re-location.
 *
 * The wrapper defines the mutation and later resolution scope for native replace
 * suggestions. It is not a Track Changes lifecycle boundary: enabling Track
 * Changes belongs to the batch apply workflow, and cleanup-time temporary
 * disabling belongs to the resolution cleanup workflow.
 */
export class ApplySuggestionOperationalWrapperResolver {
  constructor(
    private readonly suggestion: Suggestion,
    private readonly textLocator: TextLocator,
    private readonly identityBuilder: ApplySuggestionIdentityBuilder
  ) {}

  /**
   * Creates the external operational wrapper without owning Track Changes state.
   *
   * The batch-level apply workflow is responsible for enabling Track Changes before
   * replacement mutations. Real Word validation showed wrapper creation works with
   * Track Changes active; temporarily disabling it here can leave later replacement
   * mutations untracked when Word invalidates the path.
   */
  async createOperationalWrapper(
    context: Word.RequestContext,
    anchorRange: Word.Range
  ): Promise<Word.ContentControl> {
    applySuggestionObservability.logCreatingOperationalWrapper(this.suggestion.id, {
      trackChangesOwnership: "batch-apply-workflow",
    });

    const wrapper = anchorRange.insertContentControl();
    wrapper.tag = this.identityBuilder.buildOperationalWrapperTag(this.suggestion);
    wrapper.title = this.identityBuilder.serializeReplaceIdentity(
      this.identityBuilder.buildReplaceIdentity(this.suggestion)
    );
    wrapper.appearance = "Hidden";
    wrapper.cannotDelete = false;
    await context.sync();

    return wrapper;
  }

  /** Reuses a valid parent operational wrapper or returns a fail-closed error. */
  async resolveOperationalWrapper(
    context: Word.RequestContext,
    anchorRange: Word.Range
  ): Promise<ApplySuggestionOperationalWrapperResolution> {
    const parentCC =
      anchorRange.parentContentControlOrNullObject as ParentOperationalContentControl;
    parentCC.load("tag,title");
    await context.sync();

    applySuggestionObservability.logParentContentControl(this.suggestion.id, {
      hasParentContentControl: !parentCC.isNullObject,
      tag: parentCC.isNullObject ? "" : (parentCC.tag ?? ""),
      title: parentCC.isNullObject ? "" : (parentCC.title ?? ""),
    });

    if (parentCC.isNullObject) {
      return {
        kind: "success",
        wrapper: await this.createOperationalWrapper(context, anchorRange),
      };
    }

    const existingTag = parentCC.tag ?? "";
    const isStylisticArtifact =
      existingTag.startsWith(STYLISTIC_TAG_PREFIX) || /^chunk\d+-\d+$/.test(existingTag);

    if (!existingTag.startsWith(STYLISTIC_OPERATIONAL_WRAPPER_TAG_PREFIX)) {
      if (isStylisticArtifact) {
        applySuggestionObservability.warnNonOperationalStylisticContentControl(this.suggestion.id, {
          existingTag,
        });
        return { kind: "error", error: "Anchor cubierto por un Content Control existente" };
      }

      return {
        kind: "success",
        wrapper: await this.createOperationalWrapper(context, anchorRange),
      };
    }

    const existingIdentity = parseReplaceIdentityTitle(parentCC.title);
    const canReuseWrapper =
      isValidOperationalReplaceIdentity(existingIdentity, this.suggestion) &&
      typeof parentCC.getRange === "function";

    if (!canReuseWrapper) {
      applySuggestionObservability.warnOperationalWrapperIdentityMismatch(this.suggestion.id, {
        existingTag,
        title: parentCC.title ?? "",
      });
      return { kind: "error", error: "Anchor cubierto por un Content Control existente" };
    }

    applySuggestionObservability.logReusingOperationalWrapper(this.suggestion.id, {
      existingTag,
      title: parentCC.title ?? "",
    });

    return { kind: "success", wrapper: parentCC };
  }

  /** Re-finds the anchor inside the operational wrapper so mutation scope matches the wrapper. */
  async resolveAnchorInsideWrapper(
    context: Word.RequestContext,
    wrapper: Word.ContentControl
  ): Promise<Word.Range | null> {
    return this.textLocator.locate({
      context,
      container: wrapper.getRange() as unknown as WordSearchContainer,
      searchText: this.suggestion.anchor,
    });
  }
}
