import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ApplySuggestionCommand } from "./ApplySuggestionCommand";
import { WordTextLocatorAdapter } from "./WordTextLocatorAdapter";
import {
  createRange,
  installWordContext,
  makeSuggestion,
  type MockRange,
} from "./ApplySuggestionCommandTestHelper";

const textLocator = new WordTextLocatorAdapter();

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

    const result = await new ApplySuggestionCommand(
      makeSuggestion(),
      textLocator,
    ).execute();

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

    const result = await new ApplySuggestionCommand(
      makeSuggestion(),
      textLocator,
    ).execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Anchor no encontrado en el contexto",
    });
  });

  it("uses two-step search: body finds context, then context range finds anchor", async () => {
    const env = installWordContext();

    const result = await new ApplySuggestionCommand(
      makeSuggestion(),
      textLocator,
    ).execute();

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

    const result = await new ApplySuggestionCommand(
      makeSuggestion(),
      textLocator,
    ).execute();

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
      textLocator,
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

    const result = await new ApplySuggestionCommand(
      makeSuggestion(),
      textLocator,
    ).execute();

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
    env.context.document.body.search = vi
      .fn()
      .mockImplementationOnce(() => {
        throw wordError;
      })
      .mockReturnValueOnce({ items: [env.bodyRange], load: vi.fn() });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        context: contextWithAccents,
        anchor: anchorWithAccents,
        suggestedText: "lo que me dijo",
      }),
      textLocator,
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
    env.bodyRange.search = vi
      .fn()
      .mockImplementationOnce(() => {
        throw wordError;
      })
      .mockReturnValueOnce({ items: [env.anchorRange], load: vi.fn() });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        context: contextText,
        anchor: anchorText,
        suggestedText: "no tenía",
      }),
      textLocator,
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "no tenía",
      "Replace",
    );
  });

  it("resolves context when the context text itself exceeds 256 chars — step 3 must use a short locator sub-slice, not the full long slice", async () => {
    const longContext =
      "Eso también significa que no había otra forma de saber a ciencia cierta " +
      "quién era la tercera. Si lo que escuchó Mei era correcto, no verían a Anning " +
      "dentro de la preparatoria hasta el viernes, y si se guiaba por el mensaje, " +
      "la verían hasta el siguiente día, probablemente en las instalaciones de WEPO.";
    const anchorText = "Eso también significa";
    const wordError = new Error(
      "The search string is invalid or too long. SearchStringInvalidOrTooLong",
    );

    const env = installWordContext({
      documentText: longContext,
      contextText: longContext,
      anchorText,
    });

    env.context.document.body.search = vi
      .fn()
      .mockImplementationOnce(() => {
        throw wordError;
      })
      .mockReturnValueOnce({ items: [env.bodyRange], load: vi.fn() });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        context: longContext,
        anchor: anchorText,
        suggestedText: "Eso también significaba",
      }),
      textLocator,
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });

    const searchCalls = vi.mocked(env.context.document.body.search).mock.calls;
    const step3CallArg = searchCalls[1]?.[0];
    expect(typeof step3CallArg).toBe("string");
    if (typeof step3CallArg !== "string") {
      throw new TypeError("Expected step-3 search argument to be a string");
    }
    expect(step3CallArg.length).toBeLessThanOrEqual(256);
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
    const backendContext =
      "—¿Sabes quién es la tercera? —dijo, pero no la dejó responder—. " +
      "Yo tengo mis especulaciones, pero no estoy segura. Desde la primera vez " +
      'que vi a Shu siempre pensé "Ninguna chica de preparatoria habla así", ' +
      "aunque el otro día vi a Jing conversando con Ann. Y ¿sabes lo qué me dijo " +
      "cuando le pregunté?: qué era algo del equipo de atletismo…";
    const bodyTextAfterPriorTC = backendContext.replace(
      "qué era algo",
      "que era algo",
    );

    const anchorText = "lo qué me dijo";
    const wordError = new Error(
      "The search string is invalid or too long. SearchStringInvalidOrTooLong",
    );

    const env = installWordContext({
      documentText: bodyTextAfterPriorTC,
      contextText: bodyTextAfterPriorTC,
      anchorText,
    });

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
      textLocator,
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "lo que me dijo",
      "Replace",
    );
  });

  it("recovers when the step-3 candidate starts with special chars Word rejects — retries with first alphanumeric offset", async () => {
    const contextWithSpecialPrefix =
      "—¿Sabes quién es la tercera? —dijo, pero no la dejó responder—. " +
      "Yo tengo mis especulaciones, pero no estoy segura. Desde la primera vez " +
      'que vi a Shu siempre pensé "Ninguna chica de preparatoria habla así", ' +
      "aunque el otro día vi a Jing conversando con Ann. Y ¿sabes lo qué me dijo " +
      "cuando le pregunté?: qué era algo del equipo de atletismo…";
    const anchorText = "lo qué me dijo";
    const wordError = new Error(
      "The search string is invalid or too long. SearchStringInvalidOrTooLong",
    );

    const env = installWordContext({
      documentText: contextWithSpecialPrefix,
      contextText: contextWithSpecialPrefix,
      anchorText,
    });

    env.context.document.body.search = vi
      .fn()
      .mockImplementationOnce(() => {
        throw wordError;
      })
      .mockImplementationOnce(() => {
        throw wordError;
      })
      .mockReturnValueOnce({ items: [env.bodyRange], load: vi.fn() });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        context: contextWithSpecialPrefix,
        anchor: anchorText,
        suggestedText: "lo que me dijo",
      }),
      textLocator,
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "lo que me dijo",
      "Replace",
    );
  });

  it("resolves context with smart-quote chars in document text — step 3 normalizes non-standard quotes before comparing", async () => {
    const backendContext =
      "—¿Sabes quién es la tercera? —dijo, pero no la dejó responder—. " +
      "Yo tengo mis especulaciones, pero no estoy segura. Desde la primera vez " +
      'que vi a Shu siempre pensé "Ninguna chica de preparatoria habla así", ' +
      "aunque el otro día vi a Jing conversando con Ann. Y ¿sabes lo qué me dijo " +
      "cuando le pregunté?: qué era algo del equipo de atletismo…";
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
      textLocator,
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "lo que me dijo",
      "Replace",
    );
  });

  it("logs when the located context match is shorter than its containing paragraph and cannot contain the anchor", async () => {
    const fullParagraph =
      "—¿Sabes quién es la tercera? —dijo, pero no la dejó responder—. " +
      "Yo tengo mis especulaciones, pero no estoy segura. Desde la primera vez " +
      'que vi a Shu siempre pensé "Ninguna chica de preparatoria habla así", ' +
      "aunque el otro día vi a Jing conversando con Ann. Y ¿sabes lo qué me dijo " +
      "cuando le pregunté?: qué era algo del equipo de atletismo…";
    const shortMatch = createRange({
      text: "—¿",
      paragraphText: fullParagraph,
      searchSequence: [[], [], []],
    });

    const env = installWordContext({
      documentText: fullParagraph,
      contextText: fullParagraph,
      anchorText: "lo qué me dijo",
      contextSearchSequence: [[shortMatch]],
      anchorSearchSequence: [[], [], []],
      setupParagraphSearch: (_ctx, contextRange) => {
        const pr = (
          contextRange as MockRange & { paragraphs: { getFirst: () => { getRange: () => MockRange } } }
        ).paragraphs.getFirst().getRange("Whole") as MockRange;
        pr.search = vi
          .fn()
          .mockReturnValueOnce({ items: [], load: vi.fn() })
          .mockReturnValueOnce({ items: [], load: vi.fn() })
          .mockReturnValueOnce({ items: [], load: vi.fn() });
      },
    });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        context: fullParagraph,
        anchor: "lo qué me dijo",
        suggestedText: "lo que me dijo",
      }),
      textLocator,
    ).execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Anchor no encontrado en el contexto",
    });
    expect(console.log).toHaveBeenCalledWith(
      expect.stringContaining(
        'contextMatchLen=2, paragraphLen=338, anchorIndexInMatch=-1, anchorIndexInParagraph=265',
      ),
    );
    expect(env.anchorRange.insertText).not.toHaveBeenCalled();
  });

  it("expands to paragraph and finds anchor when the context match is shorter than its paragraph", async () => {
    const fullParagraph =
      "—¿Sabes quién es la tercera? —dijo, pero no la dejó responder—. " +
      "Yo tengo mis especulaciones, pero no estoy segura. Desde la primera vez " +
      'que vi a Shu siempre pensé "Ninguna chica de preparatoria habla así", ' +
      "aunque el otro día vi a Jing conversando con Ann. Y ¿sabes lo qué me dijo " +
      "cuando le pregunté?: qué era algo del equipo de atletismo…";
    const shortMatch = createRange({
      text: "—¿",
      paragraphText: fullParagraph,
      searchSequence: [[], [], []],
    });

    const anchorRangeRef = { current: null as MockRange | null };

    const env = installWordContext({
      documentText: fullParagraph,
      contextText: fullParagraph,
      anchorText: "lo qué me dijo",
      contextSearchSequence: [[shortMatch]],
      anchorSearchSequence: [[], [], []],
      anchorRangeRef,
    });

    anchorRangeRef.current = env.anchorRange;
    const shortMatchParagraphRange = (
      shortMatch as MockRange & { paragraphs: { getFirst: () => { getRange: () => MockRange } } }
    ).paragraphs.getFirst().getRange("Whole") as MockRange;
    shortMatchParagraphRange.search = vi
      .fn()
      .mockReturnValueOnce({ items: [], load: vi.fn() })
      .mockReturnValueOnce({ items: [], load: vi.fn() })
      .mockReturnValueOnce({ items: [env.anchorRange], load: vi.fn() });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        context: fullParagraph,
        anchor: "lo qué me dijo",
        suggestedText: "lo que me dijo",
      }),
      textLocator,
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(console.log).toHaveBeenCalledWith(
      expect.stringContaining(
        'contextMatchLen=2, paragraphLen=338, anchorIndexInMatch=-1, anchorIndexInParagraph=265',
      ),
    );
    expect(console.log).toHaveBeenCalledWith(
      expect.stringContaining("does not contain anchor — expanding to paragraph"),
    );
    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "lo que me dijo",
      "Replace",
    );
  });

  it("recovers when a previous nearby replacement mutates the backend context but the anchor text still exists in the paragraph", async () => {
    const backendContext =
      "pero el flujo era muy bajo para considerarse una multitud, así queAsí que Xia y Shu";
    const documentParagraph =
      "pero la afluencia era muy baja para considerarse una multitud, así queAsí que Xia y Shu";
    const anchorText = "así queAsí que";

    const shortMutatedMatch = createRange({
      text: "así queAsí que Xia y Shu",
      paragraphText: documentParagraph,
      searchSequence: [[], [], []],
    });

    const env = installWordContext({
      documentText: documentParagraph,
      contextText: documentParagraph,
      anchorText,
      contextSearchSequence: [[shortMutatedMatch]],
      anchorSearchSequence: [[], [], []],
    });

    const paragraphRange = (
      shortMutatedMatch as MockRange & {
        paragraphs: { getFirst: () => { getRange: () => MockRange } };
      }
    ).paragraphs.getFirst().getRange("Whole") as MockRange;
    paragraphRange.search = vi
      .fn()
      .mockReturnValueOnce({ items: [env.anchorRange], load: vi.fn() });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        context: backendContext,
        anchor: anchorText,
        suggestedText: "así que",
      }),
      textLocator,
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(console.log).toHaveBeenCalledWith(
      expect.stringContaining(
        'contextMatchLen=24, paragraphLen=87, anchorIndexInMatch=0, anchorIndexInParagraph=63',
      ),
    );
    expect(console.log).toHaveBeenCalledWith(
      expect.stringContaining(
        'anchor not found inside partial context match (24 chars) — retrying in paragraph (87 chars)',
      ),
    );
    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "así que",
      "Replace",
    );
  });

  it("recovers 'el flujo era muy bajo' when a later nearby replacement mutates the backend context", async () => {
    const backendContext =
      "pero el flujo era muy bajo para considerarse una multitud, así queAsí que Xia y Shu";
    const documentParagraph =
      "pero el flujo era muy bajo para considerarse una multitud, así que Xia y Shu";
    const anchorText = "el flujo era muy bajo";

    const shortMutatedMatch = createRange({
      text: "así que Xia y Shu",
      paragraphText: documentParagraph,
      searchSequence: [[], [], []],
    });

    const env = installWordContext({
      documentText: documentParagraph,
      contextText: documentParagraph,
      anchorText,
      contextSearchSequence: [[shortMutatedMatch]],
      anchorSearchSequence: [[], [], []],
    });

    const paragraphRange = (
      shortMutatedMatch as MockRange & {
        paragraphs: { getFirst: () => { getRange: () => MockRange } };
      }
    ).paragraphs.getFirst().getRange("Whole") as MockRange;
    paragraphRange.search = vi
      .fn()
      .mockReturnValueOnce({ items: [env.anchorRange], load: vi.fn() });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        context: backendContext,
        anchor: anchorText,
        suggestedText: "la afluencia era muy baja",
      }),
      textLocator,
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(console.log).toHaveBeenCalledWith(
      expect.stringContaining("does not contain anchor — expanding to paragraph"),
    );
    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "la afluencia era muy baja",
      "Replace",
    );
  });

  it("recovers 'opciones. Desde' after a later same-paragraph replacement mutates the backend context", async () => {
    const backendContext =
      "Al mismo tiempo, Xia estaba considerando varias opciones. Desde que Anning les hubiera mentido hasta que WEPO no haya querido dejar registro de Mei.";
    const documentParagraph =
      "Al mismo tiempo, Xia estaba considerando varias opciones. Desde que Anning les hubiera mentido hasta que WEPO no hubiera querido dejar registro de Mei.";
    const anchorText = "opciones. Desde";

    const shortMutatedMatch = createRange({
      text: "opciones. Desde que Anning les hubiera mentido hasta que WEPO no ",
      paragraphText: documentParagraph,
      searchSequence: [[], [], []],
    });

    const env = installWordContext({
      documentText: documentParagraph,
      contextText: documentParagraph,
      anchorText,
      contextSearchSequence: [[shortMutatedMatch]],
      anchorSearchSequence: [[], [], []],
    });

    const paragraphRange = (
      shortMutatedMatch as MockRange & {
        paragraphs: { getFirst: () => { getRange: () => MockRange } };
      }
    ).paragraphs.getFirst().getRange("Whole") as MockRange;
    paragraphRange.search = vi
      .fn()
      .mockReturnValueOnce({ items: [env.anchorRange], load: vi.fn() });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        context: backendContext,
        anchor: anchorText,
        suggestedText: "opciones: desde",
      }),
      textLocator,
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(console.log).toHaveBeenCalledWith(
      expect.stringContaining("anchor not found inside partial context match"),
    );
    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "opciones: desde",
      "Replace",
    );
  });

  it("recovers 'engranajes de su cabeza' after a later same-sentence replacement mutates the backend context", async () => {
    const backendContext =
      "Intentó forzar los engranajes de su cabeza, pero la falta de sueño y cansancio físico había hecho suficiente mella para bloquearlos.";
    const documentParagraph =
      "Intentó forzar los engranajes de su cabeza, pero la falta de sueño y cansancio físico habían hecho suficiente mella para bloquearlos.";
    const anchorText = "engranajes de su cabeza";

    const shortMutatedMatch = createRange({
      text: "engranajes de su cabeza, pero la falta de sueño y cansancio físico ",
      paragraphText: documentParagraph,
      searchSequence: [[], [], []],
    });

    const env = installWordContext({
      documentText: documentParagraph,
      contextText: documentParagraph,
      anchorText,
      contextSearchSequence: [[shortMutatedMatch]],
      anchorSearchSequence: [[], [], []],
    });

    const paragraphRange = (
      shortMutatedMatch as MockRange & {
        paragraphs: { getFirst: () => { getRange: () => MockRange } };
      }
    ).paragraphs.getFirst().getRange("Whole") as MockRange;
    paragraphRange.search = vi
      .fn()
      .mockReturnValueOnce({ items: [env.anchorRange], load: vi.fn() });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        context: backendContext,
        anchor: anchorText,
        suggestedText: "engranajes de su mente",
      }),
      textLocator,
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(console.log).toHaveBeenCalledWith(
      expect.stringContaining("anchor not found inside partial context match"),
    );
    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "engranajes de su mente",
      "Replace",
    );
  });
});
