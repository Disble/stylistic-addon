import { describe, expect, it } from "vitest";
import {
  findUniqueLocatorSubstring,
  findWhitespaceInsensitiveSlice,
} from "./WordSearchAdapter";

describe("WordSearchAdapter", () => {
  describe("findWhitespaceInsensitiveSlice", () => {
    it("finds an exact match with no whitespace differences", () => {
      expect(
        findWhitespaceInsensitiveSlice(
          "texto original",
          "Contexto con texto original.",
        ),
      ).toBe("texto original");
    });

    it("finds a match when the document has extra whitespace", () => {
      expect(
        findWhitespaceInsensitiveSlice(
          "texto original",
          "Contexto con texto\n\noriginal.",
        ),
      ).toBe("texto\n\noriginal");
    });

    it("returns null when the text is not found", () => {
      expect(
        findWhitespaceInsensitiveSlice(
          "no existe",
          "Contexto con texto original.",
        ),
      ).toBeNull();
    });

    it("returns null when searchText is empty after stripping whitespace", () => {
      expect(
        findWhitespaceInsensitiveSlice("   ", "Cualquier texto."),
      ).toBeNull();
    });

    it("normalizes smart double quotes in documentText when searching", () => {
      const documentText =
        "\u201CNinguna chica habla as\u00ED\u201D dijo ella.";
      expect(
        findWhitespaceInsensitiveSlice(
          '"Ninguna chica habla así"',
          documentText,
        ),
      ).toBe("\u201CNinguna chica habla as\u00ED\u201D");
    });

    it("normalizes smart single quotes in documentText when searching", () => {
      const documentText = "It\u2019s a test.";
      expect(findWhitespaceInsensitiveSlice("It's a test", documentText)).toBe(
        "It\u2019s a test",
      );
    });

    it(
      "ignores Word field-code control characters (" +
        String.raw`\u0013-\u0015` +
        ") in documentText",
      () => {
        const documentText = "texto\u0013campo\u0015original";
        expect(
          findWhitespaceInsensitiveSlice("textooriginal", documentText),
        ).toBe("texto\u0013campo\u0015original");
      },
    );

    it("finds anchor with accent when document has unaccented version (backend sends corrected anchor)", () => {
      // The backend may send the anchor with the diacritic correction already applied.
      // e.g. anchor="lo qué me dijo" but the document still has "lo que me dijo".
      const documentText =
        "Y ¿sabes lo que me dijo cuando le pregunté?: qué era algo del equipo.";
      expect(
        findWhitespaceInsensitiveSlice("lo qué me dijo", documentText),
      ).toBe("lo que me dijo");
    });

    it("finds anchor with no accent when document has accented version (reverse mismatch)", () => {
      // Also handle the reverse: backend sends unaccented, document has accent.
      const documentText = "aunque empezó bien, luego no.";
      expect(findWhitespaceInsensitiveSlice("empezo bien", documentText)).toBe(
        "empezó bien",
      );
    });

    it("preserves original spacing in the returned slice", () => {
      const documentText = "El   texto   tiene   espacios.";
      expect(
        findWhitespaceInsensitiveSlice("texto tiene espacios", documentText),
      ).toBe("texto   tiene   espacios");
    });
  });

  describe("findUniqueLocatorSubstring", () => {
    it("returns the full slice when it is ≤256 chars and unique", () => {
      const slice = "texto original";
      const containerText = "Contexto con texto original aquí.";
      expect(findUniqueLocatorSubstring(slice, containerText)).toBe(slice);
    });

    it("returns a shortened prefix when the slice is >256 chars but a prefix is unique", () => {
      // Build a slice of 300 chars that appears once in containerText
      const slice = "A".repeat(50) + "UNIQUE_MARKER" + "B".repeat(237);
      const containerText = slice + " seguido de más texto";
      const result = findUniqueLocatorSubstring(slice, containerText);

      expect(result).not.toBeNull();
      if (result === null) {
        throw new Error("Expected a unique locator substring");
      }
      expect(result.length).toBeLessThanOrEqual(256);
      // Must be a prefix of the original slice
      expect(slice.startsWith(result)).toBe(true);
      // Must appear exactly once in containerText
      const occurrences = containerText.split(result).length - 1;
      expect(occurrences).toBe(1);
    });

    it("returns null when no ≤256-char prefix of the slice is unique in containerText", () => {
      // Repeated pattern — no unique prefix exists within 256 chars
      const repeatingUnit = "AB";
      const slice = repeatingUnit.repeat(150); // 300 chars, all repetitions
      // containerText contains two copies so every prefix repeats
      const containerText = slice + " separador " + slice;
      expect(findUniqueLocatorSubstring(slice, containerText)).toBeNull();
    });

    it("returns the full slice unchanged when it is exactly 256 chars", () => {
      const slice = "X".repeat(256);
      const containerText = slice + " y algo más";
      expect(findUniqueLocatorSubstring(slice, containerText)).toBe(slice);
    });

    it("finds the shortest unique prefix, not just any prefix", () => {
      // The first 10 chars ("AAAAAAAAAA") are repeated in containerText.
      // The first non-repeated char is at position 11 (the first "B"),
      // so the shortest unique prefix has exactly 11 chars.
      const prefix10 = "AAAAAAAAAA"; // repeated in containerText
      const uniquePart = prefix10 + "BBBBBBBBBB"; // only slice has the B suffix
      const slice = uniquePart + "C".repeat(280); // slice is >256 chars
      const containerText =
        prefix10 + " other content " + slice + " more content";

      const result = findUniqueLocatorSubstring(slice, containerText);

      expect(result).not.toBeNull();
      if (result === null) {
        throw new Error("Expected a unique locator substring");
      }
      expect(result.length).toBeLessThanOrEqual(256);
      // Must be a prefix of the original slice
      expect(slice.startsWith(result)).toBe(true);
      // Must appear exactly once in containerText
      const occurrences = containerText.split(result).length - 1;
      expect(occurrences).toBe(1);
      // The 10-char prefix is NOT unique (it repeats), so the result must be longer
      expect(result.length).toBeGreaterThan(10);
    });

    it("handles a real-world long Spanish context paragraph", () => {
      const longContext =
        "Eso también significa que no había otra forma de saber a ciencia cierta " +
        "quién era la tercera. Si lo que escuchó Mei era correcto, no verían a Anning " +
        "dentro de la preparatoria hasta el viernes, y si se guiaba por el mensaje, " +
        "la verían hasta el siguiente día, probablemente en las instalaciones de WEPO.";
      // longContext.length > 256 ✓

      const result = findUniqueLocatorSubstring(longContext, longContext);

      expect(result).not.toBeNull();
      if (result === null) {
        throw new Error("Expected a unique locator substring");
      }
      expect(result.length).toBeLessThanOrEqual(256);
      expect(longContext.startsWith(result)).toBe(true);
    });
  });
});
