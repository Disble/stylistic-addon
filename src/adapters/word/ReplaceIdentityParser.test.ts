import { describe, expect, it } from "vitest";
import {
  getDeletedSideLocator,
  getOperationalAnchorLocator,
  isValidOperationalReplaceIdentity,
  parseReplaceIdentityTitle,
} from "./ReplaceIdentityParser";
import {
  makeOperationalWrapperTitle,
  makeSuggestion,
} from "./WordAdapterActionTestHelper";

describe("ReplaceIdentityParser", () => {
  it("parses a persisted operational-wrapper title payload", () => {
    const title = makeOperationalWrapperTitle({
      suggestionId: "s-42",
      insertedTag: "stylistic:track-change:s-42",
      deletedValue: "texto original",
      anchorValue: "Contexto con texto original.",
    });

    const parsed = parseReplaceIdentityTitle(title);

    expect(parsed).toEqual({
      suggestionId: "s-42",
      version: "operational-wrapper-v1",
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
      groupId: "s-42",
      groupIndex: 0,
      groupSize: 1,
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

  it("accepts a valid operational-wrapper identity for the matching suggestion", () => {
    const suggestion = makeSuggestion({
      id: "s-42",
      anchor: "texto original",
      context: "Contexto con texto original.",
    });
    const identity = parseReplaceIdentityTitle(
      makeOperationalWrapperTitle({
        suggestionId: suggestion.id,
        insertedTag: `stylistic:${suggestion.type}:${suggestion.id}`,
        deletedValue: suggestion.anchor,
        anchorValue: suggestion.context,
      }),
    );

    expect(isValidOperationalReplaceIdentity(identity, suggestion)).toBe(true);
  });

  it("rejects identities whose deleted-side or anchor values drift", () => {
    const suggestion = makeSuggestion({
      id: "s-42",
      anchor: "texto original",
      context: "Contexto con texto original.",
    });
    const identity = parseReplaceIdentityTitle(
      makeOperationalWrapperTitle({
        suggestionId: suggestion.id,
        insertedTag: `stylistic:${suggestion.type}:${suggestion.id}`,
        deletedValue: "texto original con drift del host",
        anchorValue: "Contexto más largo rehidratado por Word.",
      }),
    );

    expect(isValidOperationalReplaceIdentity(identity, suggestion)).toBe(false);
  });

  it("rejects identities whose inserted-side tag does not match the suggestion", () => {
    const suggestion = makeSuggestion({ id: "s-42" });
    const identity = parseReplaceIdentityTitle(
      makeOperationalWrapperTitle({
        suggestionId: suggestion.id,
        insertedTag: "stylistic:track-change:other",
      }),
    );

    expect(isValidOperationalReplaceIdentity(identity, suggestion)).toBe(false);
  });

  it("rejects identities whose deleted-side or anchor metadata is structurally empty", () => {
    const suggestion = makeSuggestion({
      id: "s-42",
      anchor: "texto original",
      context: "Contexto con texto original.",
    });
    const identity = parseReplaceIdentityTitle(
      makeOperationalWrapperTitle({
        suggestionId: suggestion.id,
        deletedValue: "",
        anchorValue: "",
      }),
    );

    expect(isValidOperationalReplaceIdentity(identity, suggestion)).toBe(false);
  });

  it("returns explicit deleted-side and operational-anchor locators for valid identities", () => {
    const suggestion = makeSuggestion({
      id: "s-42",
      anchor: "texto original",
      context: "Contexto con texto original.",
    });
    const identity = parseReplaceIdentityTitle(
      makeOperationalWrapperTitle({
        suggestionId: suggestion.id,
        insertedTag: `stylistic:${suggestion.type}:${suggestion.id}`,
        deletedValue: suggestion.anchor,
        anchorValue: suggestion.context,
      }),
    );

    expect(getDeletedSideLocator(identity, suggestion)).toBe(
      suggestion.anchor,
    );
    expect(getOperationalAnchorLocator(identity, suggestion)).toBe(
      suggestion.context,
    );
  });
});
