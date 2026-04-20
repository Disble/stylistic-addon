import { describe, expect, it } from "vitest";
import {
  findFirstAlphanumericOffset,
  findUniqueLocatorSubstring,
  findWhitespaceInsensitiveSlice,
  normalizeChar,
  removeWhitespaceWithIndices,
} from "./TextSearchCore";

describe("TextSearchCore", () => {
  describe("normalizeChar", () => {
    it("normalizes smart quotes into straight quotes", () => {
      expect(normalizeChar("\u201C")).toBe('"');
      expect(normalizeChar("\u201D")).toBe('"');
      expect(normalizeChar("\u2018")).toBe("'");
      expect(normalizeChar("\u2019")).toBe("'");
    });

    it("strips diacritics for comparison", () => {
      expect(normalizeChar("á")).toBe("a");
      expect(normalizeChar("ñ")).toBe("n");
    });
  });

  describe("removeWhitespaceWithIndices", () => {
    it("removes whitespace while preserving original indices", () => {
      expect(removeWhitespaceWithIndices("a  b\nc")).toEqual({
        text: "abc",
        indices: [0, 3, 5],
      });
    });

    it("skips Word field-code content entirely", () => {
      expect(removeWhitespaceWithIndices("texto\u0013campo\u0015original")).toEqual({
        text: "textooriginal",
        indices: [0, 1, 2, 3, 4, 12, 13, 14, 15, 16, 17, 18, 19],
      });
    });
  });

  describe("findWhitespaceInsensitiveSlice", () => {
    it("finds a match when document spacing differs", () => {
      expect(
        findWhitespaceInsensitiveSlice(
          "texto original",
          "Contexto con texto\n\noriginal.",
        ),
      ).toBe("texto\n\noriginal");
    });

    it("normalizes smart quotes in the returned match", () => {
      expect(
        findWhitespaceInsensitiveSlice(
          '"Ninguna chica habla así"',
          "\u201CNinguna chica habla así\u201D dijo ella.",
        ),
      ).toBe("\u201CNinguna chica habla así\u201D");
    });

    it("matches accented and unaccented variants", () => {
      expect(
        findWhitespaceInsensitiveSlice(
          "empezo bien",
          "aunque empezó bien, luego no.",
        ),
      ).toBe("empezó bien");
    });

    it("returns null when the search text becomes empty after normalization", () => {
      expect(findWhitespaceInsensitiveSlice("   ", "Cualquier texto")).toBeNull();
    });
  });

  describe("findUniqueLocatorSubstring", () => {
    it("returns the full slice when it already fits the default limit", () => {
      const slice = "texto original";
      const containerText = "Contexto con texto original aquí.";

      expect(findUniqueLocatorSubstring(slice, containerText)).toBe(slice);
    });

    it("returns the shortest unique prefix when the slice exceeds the limit", () => {
      const slice = "A".repeat(50) + "UNIQUE_MARKER" + "B".repeat(237);
      const containerText = slice + " seguido de más texto";
      const result = findUniqueLocatorSubstring(slice, containerText);

      expect(result).not.toBeNull();
      if (result === null) {
        throw new Error("Expected a unique locator substring");
      }

      expect(result.length).toBeLessThanOrEqual(256);
      expect(slice.startsWith(result)).toBe(true);
      expect(containerText.split(result).length - 1).toBe(1);
    });

    it("returns null when no prefix within the limit is unique", () => {
      const slice = "AB".repeat(150);
      const containerText = slice + " separador " + slice;

      expect(findUniqueLocatorSubstring(slice, containerText)).toBeNull();
    });
  });

  describe("findFirstAlphanumericOffset", () => {
    it("returns the index of the first alphanumeric character", () => {
      expect(findFirstAlphanumericOffset('---¿“hola”?')).toBe(5);
    });

    it("returns -1 when the text has no alphanumeric characters", () => {
      expect(findFirstAlphanumericOffset("---¿“”?!")).toBe(-1);
    });
  });
});
