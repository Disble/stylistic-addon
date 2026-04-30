import type { TextChunk } from "../../domain/chunking/TextChunk.types";
import type { WorkflowSuggestion } from "../../domain/mastra/MastraWorkflow.types";
import { MASTRA_BASE_URL, WORKFLOW_ID } from "../../infrastructure/config";
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
  return import("./MastraAdapter");
}

describe("MastraAdapter", () => {
  let logSpy: ReturnType<typeof vi.spyOn>;
  let errorSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    vi.doMock("../../infrastructure/config", async () => {
      const actual = await vi.importActual<typeof import("../../infrastructure/config")>(
        "../../infrastructure/config",
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
    vi.doUnmock("../../infrastructure/config");
    logSpy.mockRestore();
    errorSpy.mockRestore();
  });

  it("creates one singleton Mastra client at module load", async () => {
    const { MastraAdapter } = await importAdapterModule();

    new MastraAdapter();
    new MastraAdapter();

    expect(mastraMocks.constructor).toHaveBeenCalledOnce();
    expect(mastraMocks.constructor).toHaveBeenCalledWith({
      baseUrl: MASTRA_BASE_URL,
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
    vi.doUnmock("../../infrastructure/config");
    vi.doMock("../../infrastructure/config", async () => {
      const actual = await vi.importActual<typeof import("../../infrastructure/config")>(
        "../../infrastructure/config",
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
    vi.doUnmock("../../infrastructure/config");
    vi.doMock("../../infrastructure/config", async () => {
      const actual = await vi.importActual<typeof import("../../infrastructure/config")>(
        "../../infrastructure/config",
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
    expect(firstChunk.suggestions).toHaveLength(23);
    expect(firstChunk.suggestions[0].id).toBe("chunk0-0");
    expect(laterChunk).toEqual({
      chunkIndex: 1,
      runId: "bypass-run-1",
      status: "success",
      suggestions: [],
    });
    expect(mastraMocks.getWorkflow).not.toHaveBeenCalled();
  });

  it("submits chunk analysis with workflow input and returns the created runId", async () => {
    const { MastraAdapter } = await importAdapterModule();
    const adapter = new MastraAdapter();

    const result = await adapter.submitChunkAnalysis(
      makeChunk({ text: "Hola mundo", index: 4 }),
      "narrativa-literaria",
      "maria-garcia",
    );

    expect(mastraMocks.getWorkflow).toHaveBeenCalledWith(WORKFLOW_ID);
    expect(mastraMocks.createRun).toHaveBeenCalledOnce();
    expect(mastraMocks.start).toHaveBeenCalledWith({
      inputData: {
        text: "Hola mundo",
        genero: "narrativa-literaria",
        autorSlug: "maria-garcia",
      },
    });
    expect(result).toEqual({ chunkIndex: 4, runId: "run-from-create-run" });
  });

  it("returns an error when createRun does not expose a usable runId", async () => {
    mastraMocks.createRun.mockResolvedValueOnce({ start: mastraMocks.start });
    const { MastraAdapter } = await importAdapterModule();
    const adapter = new MastraAdapter();

    const result = await adapter.submitChunkAnalysis(
      makeChunk(),
      "general",
      "disble",
    );

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

  it("fails closed when workflow enters a resume-only state", async () => {
    mastraMocks.runById.mockResolvedValueOnce({ status: "suspended" });
    const { MastraAdapter } = await importAdapterModule();
    const adapter = new MastraAdapter();

    const result = await adapter.pollChunkAnalysis(2, "run-2");

    expect(result).toEqual({
      chunkIndex: 2,
      runId: "run-2",
      status: "failed",
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
      suggestions: [],
      error: "Invalid workflow success payload: expected suggestions[]",
    });
  });
});
