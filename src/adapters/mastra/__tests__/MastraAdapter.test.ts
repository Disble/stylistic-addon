import type { TextChunk } from "../../../domain/chunking/TextChunk.types";
import type { WorkflowSuggestion } from "../../../domain/mastra/MastraWorkflow.types";
import { MASTRA_BASE_URL, WORKFLOW_ID } from "../../../infrastructure/config";
import { createMockMastraPollOutput } from "../MockMastraPollOutputFactory";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const mastraMocks = vi.hoisted(() => {
  const details = vi.fn();
  const createRun = vi.fn();
  const start = vi.fn();
  const runById = vi.fn();
  const getWorkflow = vi.fn();
  const constructor = vi.fn();

  return {
    details,
    createRun,
    start,
    runById,
    getWorkflow,
    constructor,
    workflow: {
      details,
      createRun,
      runById,
    },
    run: {
      runId: "run-from-create-run",
      start,
    },
  };
});

vi.mock("@mastra/client-js", () => ({
  MastraClient: class {
    constructor(options: unknown) {
      mastraMocks.constructor(options);
    }

    getWorkflow(workflowId: string) {
      return mastraMocks.getWorkflow(workflowId);
    }
  },
}));

/** Builds a canonical analysis chunk fixture. */
function makeChunk(overrides: Partial<TextChunk> = {}): TextChunk {
  return {
    text: "Texto de ejemplo para analizar.",
    index: 2,
    total: 5,
    startOffset: 120,
    ...overrides,
  };
}

/** Re-imports the adapter module after resetting module state. */
async function importAdapterModule() {
  vi.resetModules();
  return import("../MastraAdapter");
}

describe("MastraAdapter", () => {
  let logSpy: ReturnType<typeof vi.spyOn>;
  let errorSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    vi.doMock("../../../infrastructure/config", async () => {
      const actual = await vi.importActual<typeof import("../../../infrastructure/config")>(
        "../../../infrastructure/config"
      );

      return {
        ...actual,
        MASTRA_POLL_BYPASS_ENABLED: false,
      };
    });
    vi.clearAllMocks();
    mastraMocks.getWorkflow.mockReturnValue(mastraMocks.workflow);
    mastraMocks.details.mockResolvedValue({ id: WORKFLOW_ID });
    mastraMocks.createRun.mockResolvedValue(mastraMocks.run);
    mastraMocks.start.mockResolvedValue({ message: "Workflow started" });
    mastraMocks.runById.mockResolvedValue({ status: "running" });
    logSpy = vi.spyOn(console, "log").mockImplementation(() => {});
    errorSpy = vi.spyOn(console, "error").mockImplementation(() => {});
  });

  afterEach(() => {
    vi.doUnmock("../../../infrastructure/config");
    logSpy.mockRestore();
    errorSpy.mockRestore();
  });

  it("defers Mastra client creation until a backend operation needs it", async () => {
    const { MastraAdapter } = await importAdapterModule();

    const firstAdapter = new MastraAdapter();
    const secondAdapter = new MastraAdapter();

    expect(firstAdapter).toBeInstanceOf(MastraAdapter);
    expect(secondAdapter).toBeInstanceOf(MastraAdapter);
    expect(mastraMocks.constructor).not.toHaveBeenCalled();
  });

  it("creates a Mastra client with the configured base URL when checking connectivity", async () => {
    const { MastraAdapter } = await importAdapterModule();
    const adapter = new MastraAdapter();

    await adapter.checkConnection();

    expect(mastraMocks.constructor).toHaveBeenCalledOnce();
    expect(mastraMocks.constructor).toHaveBeenCalledWith({
      baseUrl: MASTRA_BASE_URL,
      headers: undefined,
    });
  });

  it("returns true from checkConnection when workflow details resolve", async () => {
    const { MastraAdapter } = await importAdapterModule();
    const adapter = new MastraAdapter();

    await expect(adapter.checkConnection()).resolves.toBe(true);
    expect(mastraMocks.getWorkflow).toHaveBeenCalledWith(WORKFLOW_ID);
    expect(mastraMocks.details).toHaveBeenCalledOnce();
  });

  it("returns true from checkConnection without touching the backend when bypass is enabled", async () => {
    vi.doUnmock("../../../infrastructure/config");
    vi.doMock("../../../infrastructure/config", async () => {
      const actual = await vi.importActual<typeof import("../../../infrastructure/config")>(
        "../../../infrastructure/config"
      );

      return {
        ...actual,
        MASTRA_POLL_BYPASS_ENABLED: true,
      };
    });

    const { MastraAdapter } = await importAdapterModule();
    const adapter = new MastraAdapter();

    await expect(adapter.checkConnection()).resolves.toBe(true);
    expect(mastraMocks.getWorkflow).not.toHaveBeenCalled();
  });

  it("returns the bypass mock only for the first chunk", async () => {
    vi.doUnmock("../../../infrastructure/config");
    vi.doMock("../../../infrastructure/config", async () => {
      const actual = await vi.importActual<typeof import("../../../infrastructure/config")>(
        "../../../infrastructure/config"
      );

      return {
        ...actual,
        MASTRA_POLL_BYPASS_ENABLED: true,
      };
    });

    const { MastraAdapter } = await importAdapterModule();
    const adapter = new MastraAdapter();

    const firstChunk = await adapter.pollChunkAnalysis(0, "bypass-run-0");
    const laterChunk = await adapter.pollChunkAnalysis(1, "bypass-run-1");

    expect(firstChunk.status).toBe("success");
    expect(firstChunk.origin).toBe("backend");
    expect(firstChunk.suggestions).toHaveLength(createMockMastraPollOutput().suggestions.length);
    expect(firstChunk.suggestions[0].id).toBe("chunk0-0");
    expect(laterChunk).toEqual({
      chunkIndex: 1,
      runId: "bypass-run-1",
      status: "success",
      origin: "backend",
      suggestions: [],
    });
    expect(mastraMocks.getWorkflow).not.toHaveBeenCalled();
  });

  it("submits chunk analysis with workflow input and returns the created runId", async () => {
    const { MastraAdapter } = await importAdapterModule();
    const adapter = new MastraAdapter();

    const result = await adapter.submitChunkAnalysis(makeChunk({ text: "Hola mundo", index: 4 }), {
      documentUuid: "11111111-1111-4111-8111-111111111111",
      genero: "narrativa-literaria",
    });

    expect(mastraMocks.getWorkflow).toHaveBeenCalledWith(WORKFLOW_ID);
    expect(mastraMocks.createRun).toHaveBeenCalledOnce();
    expect(mastraMocks.start).toHaveBeenCalledWith({
      inputData: {
        text: "Hola mundo",
        documentUuid: "11111111-1111-4111-8111-111111111111",
        genero: "narrativa-literaria",
      },
    });
    expect(result).toEqual({ chunkIndex: 4, runId: "run-from-create-run" });
  });

  it("returns an error when createRun does not expose a usable runId", async () => {
    mastraMocks.createRun.mockResolvedValueOnce({ start: mastraMocks.start });
    const { MastraAdapter } = await importAdapterModule();
    const adapter = new MastraAdapter();

    const result = await adapter.submitChunkAnalysis(makeChunk(), {
      documentUuid: "11111111-1111-4111-8111-111111111111",
      genero: "general",
    });

    expect(result).toEqual({
      chunkIndex: 2,
      error: "Workflow createRun did not return a valid runId",
    });
    expect(mastraMocks.start).not.toHaveBeenCalled();
  });

  it("polls workflow state with explicit fields only and maps valid suggestions", async () => {
    const suggestions: WorkflowSuggestion[] = [
      {
        context: "Se encontró el texto observado dentro del párrafo.",
        anchor: "texto observado",
        justification: "Revisar tono",
        category: "Tono",
        severity: "low",
        type: "comment-only",
      },
      {
        context: "Xia no tenía idea de lo que estaba pasando por la mente de Mei ni Shu.",
        anchor: "ni Shu",
        suggestedText: "ni de Shu",
        justification: "Omisión de preposición en estructuras paralelas.",
        category: "gramática",
        severity: "high",
        type: "track-change",
      },
    ];
    mastraMocks.runById.mockResolvedValueOnce({
      status: "success",
      result: { suggestions },
    });

    const { MastraAdapter } = await importAdapterModule();
    const adapter = new MastraAdapter();

    const result = await adapter.pollChunkAnalysis(3, "run-3");

    expect(mastraMocks.runById).toHaveBeenCalledWith("run-3", {
      fields: ["result", "error"],
      withNestedWorkflows: false,
    });
    expect(result).toEqual({
      chunkIndex: 3,
      runId: "run-3",
      status: "success",
      origin: "backend",
      suggestions: [
        {
          id: "chunk3-0",
          context: "Se encontró el texto observado dentro del párrafo.",
          anchor: "texto observado",
          justification: "Revisar tono",
          category: "Tono",
          severity: "low",
          type: "comment-only",
        },
        {
          id: "chunk3-1",
          context: "Xia no tenía idea de lo que estaba pasando por la mente de Mei ni Shu.",
          anchor: "ni Shu",
          suggestedText: "ni de Shu",
          justification: "Omisión de preposición en estructuras paralelas.",
          category: "gramática",
          severity: "high",
          type: "track-change",
        },
      ],
    });
  });

  it("accepts delete-only track-change suggestions with empty suggestedText", async () => {
    const suggestions: WorkflowSuggestion[] = [
      {
        context: "No obstante, siguió sosteniéndola del brazo a pesar de eso.",
        anchor: " a pesar de eso",
        suggestedText: "",
        justification: "Eliminar redundancia final.",
        category: "claridad",
        severity: "medium",
        type: "track-change",
      },
    ];
    mastraMocks.runById.mockResolvedValueOnce({
      status: "success",
      result: { suggestions },
    });

    const { MastraAdapter } = await importAdapterModule();
    const adapter = new MastraAdapter();

    const result = await adapter.pollChunkAnalysis(5, "run-5");

    expect(result).toEqual({
      chunkIndex: 5,
      runId: "run-5",
      status: "success",
      origin: "backend",
      suggestions: [
        {
          id: "chunk5-0",
          context: "No obstante, siguió sosteniéndola del brazo a pesar de eso.",
          anchor: " a pesar de eso",
          suggestedText: "",
          justification: "Eliminar redundancia final.",
          category: "claridad",
          severity: "medium",
          type: "track-change",
        },
      ],
    });
  });

  it("accepts formatting track-change suggestions encoded as markdown", async () => {
    const suggestions: WorkflowSuggestion[] = [
      {
        context: "Ese era el inicio del post mortem reportado por PRIME.",
        anchor: "post mortem",
        suggestedText: "*post mortem*",
        justification: "Marcar latinismo en cursiva.",
        category: "estilo",
        severity: "low",
        type: "track-change",
      },
    ];
    mastraMocks.runById.mockResolvedValueOnce({
      status: "success",
      result: { suggestions },
    });

    const { MastraAdapter } = await importAdapterModule();
    const adapter = new MastraAdapter();

    const result = await adapter.pollChunkAnalysis(6, "run-6");

    expect(result).toEqual({
      chunkIndex: 6,
      runId: "run-6",
      status: "success",
      origin: "backend",
      suggestions: [
        {
          id: "chunk6-0",
          context: "Ese era el inicio del post mortem reportado por PRIME.",
          anchor: "post mortem",
          suggestedText: "*post mortem*",
          justification: "Marcar latinismo en cursiva.",
          category: "estilo",
          severity: "low",
          type: "track-change",
        },
      ],
    });
  });

  it("fails closed when workflow enters a resume-only state", async () => {
    mastraMocks.runById.mockResolvedValueOnce({ status: "suspended" });
    const { MastraAdapter } = await importAdapterModule();
    const adapter = new MastraAdapter();

    const result = await adapter.pollChunkAnalysis(2, "run-2");

    expect(result).toEqual({
      chunkIndex: 2,
      runId: "run-2",
      status: "failed",
      origin: "backend",
      suggestions: [],
      error:
        'Workflow entered "suspended" state and requires resume(), which this frontend does not support',
    });
  });

  it("fails closed when the workflow success payload is structurally invalid", async () => {
    mastraMocks.runById.mockResolvedValueOnce({
      status: "success",
      result: {
        suggestions: [
          {
            context: "Texto observado dentro del párrafo.",
            justification: "Revisar tono",
            category: "Tono",
            severity: "low",
            type: "comment-only",
          },
        ],
      },
    });
    const { MastraAdapter } = await importAdapterModule();
    const adapter = new MastraAdapter();

    const result = await adapter.pollChunkAnalysis(4, "run-4");

    expect(result).toEqual({
      chunkIndex: 4,
      runId: "run-4",
      status: "failed",
      origin: "frontend-terminal",
      suggestions: [],
      error: "Invalid workflow success payload: expected suggestions[]",
    });
  });

  it("cancels an existing backend run by rehydrating the run from runId", async () => {
    const cancel = vi.fn().mockResolvedValue({ message: "Workflow run canceled" });
    mastraMocks.createRun.mockResolvedValueOnce({
      runId: "run-cancel",
      cancel,
      start: mastraMocks.start,
    });

    const { MastraAdapter } = await importAdapterModule();
    const adapter = new MastraAdapter();

    const result = await adapter.cancelChunkAnalysis(7, "run-cancel");

    expect(mastraMocks.createRun).toHaveBeenCalledWith({ runId: "run-cancel" });
    expect(cancel).toHaveBeenCalledOnce();
    expect(result).toEqual({ chunkIndex: 7, runId: "run-cancel", canceled: true });
  });
});
