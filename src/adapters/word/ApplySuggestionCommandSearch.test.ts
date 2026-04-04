import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ApplySuggestionCommand } from "./ApplySuggestionCommand";
import {
  createRange,
  installWordContext,
  makeSuggestion,
} from "./ApplySuggestionCommandTestHelper";

describe("ApplySuggestionCommand search behavior", () => {
  beforeEach(() => {
    vi.spyOn(console, "log").mockImplementation(() => undefined);
    vi.spyOn(console, "warn").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("returns a distinct error when context is not found", async () => {
    installWordContext({ contextSearchSequence: [[], [], []] });

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Anchor no encontrado en el contexto",
    });
  });

  it("returns a distinct error when anchor is not found inside the located context", async () => {
    installWordContext({
      contextSearchSequence: [
        [createRange({ text: "Contexto con texto original." })],
      ],
      anchorSearchSequence: [[], [], []],
    });

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Anchor no encontrado en el contexto",
    });
  });

  it("uses two-step search: body finds context, then context range finds anchor", async () => {
    const env = installWordContext();

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.context.document.body.search).toHaveBeenCalledWith(
      "Contexto con texto original.",
      { matchCase: true, matchWholeWord: false },
    );
    expect(env.bodyRange.search).toHaveBeenCalledWith("texto original", {
      matchCase: true,
      matchWholeWord: false,
    });
    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "texto sugerido",
      "Replace",
    );
  });

  it("tries ignorePunct+ignoreSpace when exact anchor search fails", async () => {
    const env = installWordContext();
    env.bodyRange.search = vi
      .fn()
      .mockReturnValueOnce({ items: [], load: vi.fn() })
      .mockReturnValueOnce({ items: [env.anchorRange], load: vi.fn() });

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.bodyRange.search).toHaveBeenNthCalledWith(1, "texto original", {
      matchCase: true,
      matchWholeWord: false,
    });
    expect(env.bodyRange.search).toHaveBeenNthCalledWith(2, "texto original", {
      matchCase: true,
      matchWholeWord: false,
      ignorePunct: true,
      ignoreSpace: true,
    });
  });

  it("skips the exact search when context text exceeds Word's 256-char search limit", async () => {
    const longContext = `Prefijo ${"x".repeat(270)}`;
    const env = installWordContext({
      documentText: longContext,
      contextText: longContext,
    });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({ context: longContext }),
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.context.document.body.search).toHaveBeenCalledTimes(1);
    expect(env.context.document.body.search).toHaveBeenCalledWith(longContext, {
      matchCase: true,
      matchWholeWord: false,
      ignorePunct: true,
      ignoreSpace: true,
    });
  });

  it("falls back to a whitespace-insensitive slice when backend and document spacing differ", async () => {
    const fallbackAnchor = createRange({ text: "texto\n\noriginal" });
    const env = installWordContext({
      contextText: "Contexto con texto\n\noriginal.",
      anchorSearchSequence: [[], [], [fallbackAnchor]],
    });

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.bodyRange.load).toHaveBeenCalledWith("text");
    expect(env.bodyRange.search).toHaveBeenNthCalledWith(
      3,
      "texto\n\noriginal",
      {
        matchCase: true,
        matchWholeWord: false,
      },
    );
    expect(fallbackAnchor.insertText).toHaveBeenCalledWith(
      "texto sugerido",
      "Replace",
    );
  });

  it("recovers when body.search() throws SearchStringInvalidOrTooLong — falls back to text scan for context", async () => {
    const contextWithAccents =
      "También entendió que lo qué me dijo era correcto.";
    const anchorWithAccents = "lo qué me dijo";
    const wordError = new Error(
      "The search string is invalid or too long. SearchStringInvalidOrTooLong",
    );

    const env = installWordContext({
      contextText: contextWithAccents,
      anchorText: anchorWithAccents,
      documentText: contextWithAccents,
    });
    // Steps 1 throws; code skips step 2 (same invalid string) and tries step 3 (slice) which succeeds
    env.context.document.body.search = vi
      .fn()
      .mockImplementationOnce(() => {
        throw wordError;
      }) // step 1: exact — fails
      .mockReturnValueOnce({ items: [env.bodyRange], load: vi.fn() }); // step 3: slice — succeeds

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        context: contextWithAccents,
        anchor: anchorWithAccents,
        suggestedText: "lo que me dijo",
      }),
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "lo que me dijo",
      "Replace",
    );
  });

  it("recovers when contextRange.search() throws SearchStringInvalidOrTooLong — falls back to text scan for anchor", async () => {
    const contextText =
      "Pronto entendió que como Mei no tiene un reloj inteligente, entonces a ella le llegan al móvil.";
    const anchorText = "no tiene";
    const wordError = new Error(
      "The search string is invalid or too long. SearchStringInvalidOrTooLong",
    );

    const env = installWordContext({
      contextText,
      anchorText,
    });
    // Step 1 throws; code skips step 2 and tries step 3 (slice) which succeeds
    env.bodyRange.search = vi
      .fn()
      .mockImplementationOnce(() => {
        throw wordError;
      }) // step 1: exact — fails
      .mockReturnValueOnce({ items: [env.anchorRange], load: vi.fn() }); // step 3: slice — succeeds

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        context: contextText,
        anchor: anchorText,
        suggestedText: "no tenía",
      }),
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "no tenía",
      "Replace",
    );
  });

  it("resolves context when the context text itself exceeds 256 chars — step 3 must use a short locator sub-slice, not the full long slice", async () => {
    // Reproduces the real failure: context is ~305 chars. Steps 1 and 2 throw
    // SearchStringInvalidOrTooLong. Step 3 finds the slice but that slice is still
    // >256 chars — so calling search() with it would also fail. The fix must
    // extract a short (≤256) prefix from the slice to use as the actual search string.
    const longContext =
      "Eso también significa que no había otra forma de saber a ciencia cierta " +
      "quién era la tercera. Si lo que escuchó Mei era correcto, no verían a Anning " +
      "dentro de la preparatoria hasta el viernes, y si se guiaba por el mensaje, " +
      "la verían hasta el siguiente día, probablemente en las instalaciones de WEPO.";
    // longContext.length > 256 ✓

    const anchorText = "Eso también significa";
    const wordError = new Error(
      "The search string is invalid or too long. SearchStringInvalidOrTooLong",
    );

    const env = installWordContext({
      documentText: longContext,
      contextText: longContext,
      anchorText,
    });

    // Body search: step 1 is SKIPPED (length > 256); step 2 throws (long + accented);
    // step 3 calls search() with a ≤256-char sub-slice — must succeed.
    env.context.document.body.search = vi
      .fn()
      .mockImplementationOnce(() => {
        throw wordError;
      }) // step 2: ignorePunct/ignoreSpace — too long
      .mockReturnValueOnce({ items: [env.bodyRange], load: vi.fn() }); // step 3: short slice ✓

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        context: longContext,
        anchor: anchorText,
        suggestedText: "Eso también significaba",
      }),
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });

    // The second body.search() call (index 1) is step 3 — must use a string ≤256 chars
    // (step 1 is skipped because length > 256; step 2 is call index 0)
    const searchCalls = vi.mocked(env.context.document.body.search).mock.calls;
    const step3CallArg = searchCalls[1]?.[0];
    expect(typeof step3CallArg).toBe("string");
    if (typeof step3CallArg !== "string") {
      throw new TypeError("Expected step-3 search argument to be a string");
    }
    expect(step3CallArg.length).toBeLessThanOrEqual(256);
    // And it must be a true prefix of the original long context (starts the same way)
    expect(
      longContext.startsWith(step3CallArg) ||
        longContext.includes(step3CallArg),
    ).toBe(true);

    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "Eso también significaba",
      "Replace",
    );
  });

  it("resolves context after a previous tracked change altered a diacritic in the same paragraph — simulates two-suggestion same-paragraph scenario", async () => {
    // Reproduces the chunk0-2/chunk0-3 real-world failure:
    //
    // chunk0-3 applies FIRST (reversed order): "qué era algo" → "que era algo"
    // After that, body.text has "que era algo" (no accent) where it had "qué era algo".
    //
    // chunk0-2 applies SECOND: its context includes "qué era algo" (accented, as originally
    // sent by the backend). body.text now has the corrected form "que era algo".
    //
    // The fix must: step 3 normalizes diacritics from BOTH sides so "qué" == "que"
    // and the context is still found.
    const backendContext =
      "—¿Sabes quién es la tercera? —dijo, pero no la dejó responder—. " +
      "Yo tengo mis especulaciones, pero no estoy segura. Desde la primera vez " +
      'que vi a Shu siempre pensé "Ninguna chica de preparatoria habla así", ' +
      "aunque el otro día vi a Jing conversando con Ann. Y ¿sabes lo qué me dijo " +
      "cuando le pregunté?: qué era algo del equipo de atletismo…";
    // Simulates body.text AFTER chunk0-3 already applied "qué era algo" → "que era algo"
    const bodyTextAfterPriorTC = backendContext.replace(
      "qué era algo",
      "que era algo",
    );

    const anchorText = "lo qué me dijo";
    const wordError = new Error(
      "The search string is invalid or too long. SearchStringInvalidOrTooLong",
    );

    const env = installWordContext({
      // body.text reflects the post-TC state (diacritic removed by prior suggestion)
      documentText: bodyTextAfterPriorTC,
      // contextText is what contextRange.text returns — same post-TC state
      contextText: bodyTextAfterPriorTC,
      anchorText,
    });

    // Steps 1 is SKIPPED (length > 256); step 2 throws (accents + long); step 3 must
    // find the context despite "qué" (backend) vs "que" (document) diacritic mismatch.
    env.context.document.body.search = vi
      .fn()
      .mockImplementationOnce(() => {
        throw wordError;
      }) // step 2: throws
      .mockReturnValueOnce({ items: [env.bodyRange], load: vi.fn() }); // step 3: succeeds

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        context: backendContext,
        anchor: anchorText,
        suggestedText: "lo que me dijo",
      }),
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "lo que me dijo",
      "Replace",
    );
  });

  it("resolves context with smart-quote chars in document text — step 3 normalizes non-standard quotes before comparing", async () => {
    // Reproduces the failure where backend sends context with ASCII quotes `"..."` but
    // Word's body.text contains smart/typographic quotes `\u201C...\u201D`.
    // The whitespace-insensitive scan must still find the match.
    const backendContext =
      "—¿Sabes quién es la tercera? —dijo, pero no la dejó responder—. " +
      "Yo tengo mis especulaciones, pero no estoy segura. Desde la primera vez " +
      'que vi a Shu siempre pensé "Ninguna chica de preparatoria habla así", ' +
      "aunque el otro día vi a Jing conversando con Ann. Y ¿sabes lo qué me dijo " +
      "cuando le pregunté?: qué era algo del equipo de atletismo…";
    // Word may store smart quotes \u201C / \u201D where the backend has straight " / "
    const wordDocumentText = backendContext
      .replace(/"/g, "\u201C")
      .replace(/"/g, "\u201D");

    const anchorText = "lo qué me dijo";
    const wordError = new Error(
      "The search string is invalid or too long. SearchStringInvalidOrTooLong",
    );

    const env = installWordContext({
      documentText: wordDocumentText,
      contextText: wordDocumentText,
      anchorText,
    });

    // Steps 1 is SKIPPED (length > 256); step 2 throws; step 3 must find the anchor despite quote normalization
    env.context.document.body.search = vi
      .fn()
      .mockImplementationOnce(() => {
        throw wordError;
      })
      .mockReturnValueOnce({ items: [env.bodyRange], load: vi.fn() });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        context: backendContext,
        anchor: anchorText,
        suggestedText: "lo que me dijo",
      }),
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "lo que me dijo",
      "Replace",
    );
  });
});
