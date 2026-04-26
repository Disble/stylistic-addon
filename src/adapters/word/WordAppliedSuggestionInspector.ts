/* global Word, console */

import {
  STYLISTIC_IDENTITY_TITLE_PREFIX,
  STYLISTIC_TAG_PREFIX,
} from "../../infrastructure/config";
import { parseReplaceIdentityTitle } from "./ReplaceIdentityParser";

/** Reads already-applied Stylistic originals from the document. */
export class WordAppliedSuggestionInspector {
  /** Collects all persisted original texts for active Stylistic suggestions. */
  async getAppliedOriginalTexts(): Promise<Set<string>> {
    console.log(
      "🛡️ [WordAdapter] Consultando CCs de Stylistic (track-change + comment-only)...",
    );
    return Word.run(async (context) => {
      const allCCs = context.document.contentControls;
      allCCs.load("items/tag,items/title");
      await context.sync();

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
}
