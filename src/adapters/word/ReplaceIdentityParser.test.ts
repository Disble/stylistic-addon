import { describe, expect, it } from "vitest";
import {
  isValidCompoundReplaceIdentity,
  parseReplaceIdentityTitle,
} from "./ReplaceIdentityParser";
import {
  makeCompoundV2Title,
  makeSuggestion,
} from "./WordAdapterActionTestHelper";

describe("ReplaceIdentityParser", () => {
  it("parses a persisted compound-v2 title payload", () => {
    const title = makeCompoundV2Title({
      suggestionId: "s-42",
      insertedTag: "stylistic:track-change:s-42",
      deletedValue: "texto original",
      anchorValue: "Contexto con texto original.",
    });

    const parsed = parseReplaceIdentityTitle(title);

    expect(parsed).toEqual({
      suggestionId: "s-42",
      version: "compound-v2",
      insertedSideRef: {
        kind: "content-control",
        role: "inserted-side",
        value: "stylistic:track-change:s-42",
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
    });
  });

  it("returns null when the title has no Stylistic identity prefix", () => {
    expect(parseReplaceIdentityTitle("texto original")).toBeNull();
  });

  it("returns null when the title payload is malformed JSON", () => {
    expect(
      parseReplaceIdentityTitle('stylistic-meta-v2:{"broken":'),
    ).toBeNull();
  });

  it("accepts a valid compound-v2 identity for the matching suggestion", () => {
    const suggestion = makeSuggestion({
      id: "s-42",
      anchor: "texto original",
      context: "Contexto con texto original.",
    });
    const identity = parseReplaceIdentityTitle(
      makeCompoundV2Title({
        suggestionId: suggestion.id,
        insertedTag: `stylistic:${suggestion.type}:${suggestion.id}`,
        deletedValue: suggestion.anchor,
        anchorValue: suggestion.context,
      }),
    );

    expect(isValidCompoundReplaceIdentity(identity, suggestion)).toBe(true);
  });

  it("rejects identities whose inserted-side tag does not match the suggestion", () => {
    const suggestion = makeSuggestion({ id: "s-42" });
    const identity = parseReplaceIdentityTitle(
      makeCompoundV2Title({
        suggestionId: suggestion.id,
        insertedTag: "stylistic:track-change:other",
      }),
    );

    expect(isValidCompoundReplaceIdentity(identity, suggestion)).toBe(false);
  });

  it("rejects identities whose deleted-side or anchor values drift from the suggestion", () => {
    const suggestion = makeSuggestion({
      id: "s-42",
      anchor: "texto original",
      context: "Contexto con texto original.",
    });
    const identity = parseReplaceIdentityTitle(
      makeCompoundV2Title({
        suggestionId: suggestion.id,
        deletedValue: "texto distinto",
        anchorValue: "Contexto distinto.",
      }),
    );

    expect(isValidCompoundReplaceIdentity(identity, suggestion)).toBe(false);
  });
});
