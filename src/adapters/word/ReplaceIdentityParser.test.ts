import { describe, expect, it } from "vitest";
import {
  getDeletedSideLocator,
  getOperationalAnchorLocator,
  isValidCompoundReplaceIdentity,
  parseReplaceIdentityTitle,
  scoreCompoundReplaceIdentityMatch,
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

  it("accepts identities whose deleted-side and anchor values drift while the inserted-side tag stays valid", () => {
    const suggestion = makeSuggestion({
      id: "s-42",
      anchor: "texto original",
      context: "Contexto con texto original.",
    });
    const identity = parseReplaceIdentityTitle(
      makeCompoundV2Title({
        suggestionId: suggestion.id,
        insertedTag: `stylistic:${suggestion.type}:${suggestion.id}`,
        deletedValue: "texto original con drift del host",
        anchorValue: "Contexto más largo rehidratado por Word.",
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

  it("rejects identities whose deleted-side or anchor metadata is structurally empty", () => {
    const suggestion = makeSuggestion({
      id: "s-42",
      anchor: "texto original",
      context: "Contexto con texto original.",
    });
    const identity = parseReplaceIdentityTitle(
      makeCompoundV2Title({
        suggestionId: suggestion.id,
        deletedValue: "",
        anchorValue: "",
      }),
    );

    expect(isValidCompoundReplaceIdentity(identity, suggestion)).toBe(false);
  });

  it("scores exact deleted-side and anchor matches above structurally valid drifted identities", () => {
    const suggestion = makeSuggestion({
      id: "s-42",
      anchor: "texto original",
      context: "Contexto con texto original.",
    });
    const exactIdentity = parseReplaceIdentityTitle(
      makeCompoundV2Title({
        suggestionId: suggestion.id,
        insertedTag: `stylistic:${suggestion.type}:${suggestion.id}`,
        deletedValue: suggestion.anchor,
        anchorValue: suggestion.context,
      }),
    );
    const driftedIdentity = parseReplaceIdentityTitle(
      makeCompoundV2Title({
        suggestionId: suggestion.id,
        insertedTag: `stylistic:${suggestion.type}:${suggestion.id}`,
        deletedValue: "texto original con drift del host",
        anchorValue: "Contexto más largo rehidratado por Word.",
      }),
    );

    expect(scoreCompoundReplaceIdentityMatch(exactIdentity, suggestion)).toBe(3);
    expect(scoreCompoundReplaceIdentityMatch(driftedIdentity, suggestion)).toBe(1);
  });

  it("returns explicit deleted-side and operational-anchor locators for valid identities", () => {
    const suggestion = makeSuggestion({
      id: "s-42",
      anchor: "texto original",
      context: "Contexto con texto original.",
    });
    const identity = parseReplaceIdentityTitle(
      makeCompoundV2Title({
        suggestionId: suggestion.id,
        insertedTag: `stylistic:${suggestion.type}:${suggestion.id}`,
        deletedValue: "texto borrado persistido",
        anchorValue: "Contexto operativo persistido.",
      }),
    );

    expect(getDeletedSideLocator(identity, suggestion)).toBe(
      "texto borrado persistido",
    );
    expect(getOperationalAnchorLocator(identity, suggestion)).toBe(
      "Contexto operativo persistido.",
    );
  });
});
