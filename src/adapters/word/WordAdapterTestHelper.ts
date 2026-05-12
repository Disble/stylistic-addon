import type { CommandResult } from "../../domain/DocumentApplication.types";
import type { Suggestion } from "../../domain/suggestion/Suggestion.types";
import type { WordRunCallback } from "./WordAdapterTestHelper.types";

const hoistedCommandMocks = vi.hoisted(() => ({
  constructor: vi.fn<(suggestion: Suggestion) => void>(),
  execute: vi.fn<(suggestion: Suggestion) => Promise<CommandResult>>(),
}));

const hoistedCleanupMocks = vi.hoisted(() => ({
  getCleanupPreview: vi.fn(),
  cleanupResolvedComments: vi.fn(),
}));

/**
 * Returns the shared command mock registry used by WordAdapter tests.
 */
export function getCommandMocks() {
  return hoistedCommandMocks;
}

/**
 * Returns the shared cleanup mock registry used by WordAdapter tests.
 */
export function getCleanupMocks() {
  return hoistedCleanupMocks;
}

vi.mock("./ApplySuggestionCommand", () => ({
  ApplySuggestionCommand: class {
    private readonly suggestion: Suggestion;

    constructor(suggestion: Suggestion) {
      this.suggestion = suggestion;
      hoistedCommandMocks.constructor(suggestion);
    }

    execute() {
      return hoistedCommandMocks.execute(this.suggestion);
    }
  },
}));

vi.mock("./cleanup/CommentCleanup", () => ({
  getCleanupPreview: hoistedCleanupMocks.getCleanupPreview,
  cleanupResolvedComments: hoistedCleanupMocks.cleanupResolvedComments,
  OVERLAPPING_RELATIONS: [
    "Equal",
    "Contains",
    "ContainsStart",
    "ContainsEnd",
    "Inside",
    "InsideStart",
    "InsideEnd",
    "OverlapsBefore",
    "OverlapsAfter",
  ],
  STYLISTIC_TAG_PREFIX: "stylistic:",
}));

/**
 * Builds a canonical suggestion fixture for `WordAdapter` tests.
 */
export function makeSuggestion(overrides: Partial<Suggestion> = {}): Suggestion {
  const anchor = overrides.anchor ?? "texto original";
  return {
    id: "s-1",
    context: overrides.context ?? `Contexto con ${anchor}.`,
    anchor,
    suggestedText: "texto sugerido",
    justification: "Mas claro",
    category: "Claridad",
    severity: "medium",
    type: "track-change",
    ...overrides,
  };
}

/**
 * Installs a `Word.run` mock that resolves against the provided context.
 */
export function installWordWithContext(context: any) {
  const run = vi.fn(async <T>(callback: WordRunCallback<T>) => callback(context));
  const wordGlobal = globalThis as unknown as {
    Word?: {
      run: typeof run;
      ChangeTrackingMode: Record<string, string>;
    };
  };
  wordGlobal.Word = {
    run,
    ChangeTrackingMode: {
      off: "off",
      trackAll: "trackAll",
      trackMine: "trackMine",
    },
  };
  return run;
}

/**
 * Installs a rejecting `Word.run` mock for adapter sad paths.
 */
export function installRejectingWord(error: Error) {
  const run = vi.fn().mockRejectedValue(error);
  const wordGlobal = globalThis as unknown as {
    Word?: { run: typeof run };
  };
  wordGlobal.Word = { run };
  return run;
}

/**
 * Installs an `Office.context.document.settings` mock for document identity tests.
 */
export function installOfficeDocumentSettings(
  options: {
    existingValue?: unknown;
    saveErrorMessage?: string;
  } = {}
) {
  const get = vi.fn((_key: string) => options.existingValue);
  const set = vi.fn();
  const saveAsync = vi.fn((callback: (result: Office.AsyncResult<void>) => void) => {
    if (options.saveErrorMessage) {
      callback({
        status: "failed" as unknown as Office.AsyncResultStatus,
        error: { message: options.saveErrorMessage } as Office.Error,
      } as Office.AsyncResult<void>);
      return;
    }

    callback({
      status: "succeeded" as unknown as Office.AsyncResultStatus,
      value: undefined,
    } as Office.AsyncResult<void>);
  });

  const settings = { get, set, saveAsync } as unknown as Office.Settings;
  const officeGlobal = globalThis as unknown as {
    Office?: {
      AsyncResultStatus?: Record<string, string>;
      context?: {
        document?: {
          settings?: Office.Settings;
        };
      };
    };
  };

  officeGlobal.Office = {
    AsyncResultStatus: {
      Succeeded: "succeeded",
      Failed: "failed",
      ...(officeGlobal.Office?.AsyncResultStatus ?? {}),
    },
    context: {
      document: {
        settings,
      },
    },
  };

  return { settings, get, set, saveAsync };
}

/**
 * Builds a paragraph snapshot compatible with `getTextToAnalyze()` tests.
 */
export function makeParagraph(
  text: string,
  overrides: Partial<{
    styleBuiltIn: string;
    firstLineIndent: number;
    leftIndent: number;
  }> = {}
) {
  return {
    text,
    styleBuiltIn: "Normal",
    firstLineIndent: 0,
    leftIndent: 0,
    ...overrides,
  };
}
