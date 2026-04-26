import { describe, expect, it } from "vitest";
import { makeSuggestion } from "../ApplySuggestionCommandTestHelper";
import { ApplySuggestionIdentityBuilder } from "./ApplySuggestionIdentityBuilder";

describe("ApplySuggestionIdentityBuilder", () => {
  it("builds canonical operational-wrapper replace identity metadata", () => {
    const builder = new ApplySuggestionIdentityBuilder();
    const suggestion = makeSuggestion({
      id: "replace-identity-1",
      anchor: "texto original",
      context: "Contexto con texto original.",
      suggestedText: "texto sugerido",
    });

    expect(builder.buildSuggestionTag(suggestion)).toBe(
      "stylistic:track-change:replace-identity-1",
    );
    expect(builder.buildOperationalWrapperTag(suggestion)).toBe(
      "stylistic-operational-wrapper:replace-identity-1",
    );
    expect(builder.buildReplaceIdentity(suggestion)).toEqual({
      suggestionId: "replace-identity-1",
      version: "operational-wrapper-v1",
      insertedSideRef: {
        kind: "content-control",
        role: "inserted-side",
        value: "stylistic:track-change:replace-identity-1",
      },
      deletedSideRef: {
        kind: "anchor",
        role: "deleted-side",
        value: "texto original",
      },
      anchorRef: {
        kind: "anchor",
        role: "operational-anchor",
        value: "Contexto con texto original.",
      },
      groupId: "replace-identity-1",
      groupIndex: 0,
      groupSize: 1,
    });
  });

  it("serializes persisted title payloads with the configured prefix", () => {
    const builder = new ApplySuggestionIdentityBuilder();
    const identity = builder.buildReplaceIdentity(
      makeSuggestion({ id: "replace-identity-2" }),
    );

    expect(builder.serializeReplaceIdentity(identity)).toBe(
      `stylistic-meta-v2:${JSON.stringify(identity)}`,
    );
  });
});
