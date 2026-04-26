/* global Word */

import type { Suggestion } from "../../../domain/suggestion/Suggestion.types";
import {
  STYLISTIC_OPERATIONAL_WRAPPER_TAG_PREFIX,
  STYLISTIC_TAG_PREFIX,
} from "../../../infrastructure/config";
import {
  isValidOperationalReplaceIdentity,
  parseReplaceIdentityTitle,
} from "../ReplaceIdentityParser";
import type {
  TextLocator,
  WordSearchContainer,
} from "../WordTextLocatorContext";
import type { ApplySuggestionIdentityBuilder } from "./ApplySuggestionIdentityBuilder";

type ParentOperationalContentControl = Word.ContentControl & {
  tag?: string;
  title?: string;
  getRange?: () => Word.Range;
};

/** Result of resolving or creating an operational wrapper. */
type ApplySuggestionOperationalWrapperResolution =
  | {
      /** Reusable or newly created wrapper. */
      wrapper: Word.ContentControl;

      /** No error when wrapper resolution succeeded. */
      error?: undefined;
    }
  | {
      /** No wrapper is available when resolution fails closed. */
      wrapper?: undefined;

      /** Stable fail-closed error message. */
      error: string;
    };

/**
 * Owns operational-wrapper creation, reuse, and in-wrapper re-location.
 */
export class ApplySuggestionOperationalWrapperResolver {
  constructor(
    private readonly suggestion: Suggestion,
    private readonly textLocator: TextLocator,
    private readonly identityBuilder: ApplySuggestionIdentityBuilder,
  ) {}

  /** Creates the external operational wrapper before Track Changes mutates the anchor. */
  async createOperationalWrapper(
    context: Word.RequestContext,
    anchorRange: Word.Range,
  ): Promise<Word.ContentControl> {
    context.document.load("changeTrackingMode");
    await context.sync();

    const previousTrackingMode = context.document.changeTrackingMode;
    if (previousTrackingMode !== Word.ChangeTrackingMode.off) {
      context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
      await context.sync();
    }

    const wrapper = anchorRange.insertContentControl();
    wrapper.tag = this.identityBuilder.buildOperationalWrapperTag(
      this.suggestion,
    );
    wrapper.title = this.identityBuilder.serializeReplaceIdentity(
      this.identityBuilder.buildReplaceIdentity(this.suggestion),
    );
    wrapper.appearance = "Hidden";
    wrapper.cannotDelete = false;
    await context.sync();

    if (previousTrackingMode !== Word.ChangeTrackingMode.off) {
      context.document.changeTrackingMode = previousTrackingMode;
      await context.sync();
    }

    return wrapper;
  }

  /** Reuses a valid parent operational wrapper or returns a fail-closed error. */
  async resolveOperationalWrapper(
    context: Word.RequestContext,
    anchorRange: Word.Range,
  ): Promise<ApplySuggestionOperationalWrapperResolution> {
    const parentCC =
      anchorRange.parentContentControlOrNullObject as ParentOperationalContentControl;
    parentCC.load("tag,title");
    await context.sync();

    if (parentCC.isNullObject) {
      return {
        wrapper: await this.createOperationalWrapper(context, anchorRange),
      };
    }

    const existingTag = parentCC.tag ?? "";
    const isStylisticArtifact =
      existingTag.startsWith(STYLISTIC_TAG_PREFIX) ||
      /^chunk\d+-\d+$/.test(existingTag);

    if (!existingTag.startsWith(STYLISTIC_OPERATIONAL_WRAPPER_TAG_PREFIX)) {
      if (isStylisticArtifact) {
        return { error: "Anchor cubierto por un Content Control existente" };
      }

      return {
        wrapper: await this.createOperationalWrapper(context, anchorRange),
      };
    }

    const existingIdentity = parseReplaceIdentityTitle(parentCC.title);
    const canReuseWrapper =
      isValidOperationalReplaceIdentity(existingIdentity, this.suggestion) &&
      typeof parentCC.getRange === "function";

    if (!canReuseWrapper) {
      return { error: "Anchor cubierto por un Content Control existente" };
    }

    return { wrapper: parentCC };
  }

  /** Re-finds the anchor inside the operational wrapper so mutation scope matches the wrapper. */
  async resolveAnchorInsideWrapper(
    context: Word.RequestContext,
    wrapper: Word.ContentControl,
  ): Promise<Word.Range | null> {
    return this.textLocator.locate({
      context,
      container: wrapper.getRange() as unknown as WordSearchContainer,
      searchText: this.suggestion.anchor,
    });
  }
}
