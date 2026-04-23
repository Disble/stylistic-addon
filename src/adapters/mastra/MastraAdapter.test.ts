import type { TextChunk, WorkflowSuggestion } from "../../domain/types";
import { MASTRA_BASE_URL, WORKFLOW_ID } from "../../infrastructure/config";

const mastraMocks = vi.hoisted(() => {
  const details = vi.fn();
  const createRun = vi.fn();
  const start = vi.fn();
  const runById = vi.fn();
  const getWorkflow = vi.fn();
  const mockConstructor = vi.fn();

  return {
    details,
    createRun,
    start,
    runById,
    getWorkflow,
    constructor: mockConstructor,
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

function makeChunk(overrides: Partial<TextChunk> = {}): TextChunk {
  return {
    text: "Texto de ejemplo para analizar.",
    index: 2,
    total: 5,
    startOffset: 120,
    ...overrides,
  };
}

async function importAdapterModule() {
  vi.resetModules();
  return import("./MastraAdapter");
}

describe("MastraAdapter", () => {
  let logSpy: ReturnType<typeof vi.spyOn>;
  let errorSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    vi.doUnmock("../../infrastructure/config");
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

  describe("module initialization", () => {
    it("creates one Mastra client at module load with the configured baseUrl", async () => {
      const { MastraAdapter } = await importAdapterModule();

      new MastraAdapter();
      new MastraAdapter();

      expect(mastraMocks.constructor).toHaveBeenCalledOnce();
      expect(mastraMocks.constructor).toHaveBeenCalledWith({
        baseUrl: MASTRA_BASE_URL,
      });
    });
  });

  describe("checkConnection", () => {
    it("returns true when workflow details resolve", async () => {
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      await expect(adapter.checkConnection()).resolves.toBe(true);

      expect(mastraMocks.getWorkflow).toHaveBeenCalledWith(WORKFLOW_ID);
      expect(mastraMocks.details).toHaveBeenCalledOnce();
    });

    it("returns false when fetching workflow details rejects", async () => {
      mastraMocks.details.mockRejectedValueOnce(new Error("Backend down"));
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      await expect(adapter.checkConnection()).resolves.toBe(false);

      expect(mastraMocks.getWorkflow).toHaveBeenCalledWith(WORKFLOW_ID);
      expect(errorSpy).toHaveBeenCalled();
    });

    it("returns false when getWorkflow throws synchronously", async () => {
      mastraMocks.getWorkflow.mockImplementationOnce(() => {
        throw new Error("Workflow registry unavailable");
      });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      await expect(adapter.checkConnection()).resolves.toBe(false);

      expect(mastraMocks.details).not.toHaveBeenCalled();
    });

    it("returns true without touching the backend when the bypass is enabled", async () => {
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
      expect(mastraMocks.details).not.toHaveBeenCalled();
    });
  });

  describe("submitChunkAnalysis", () => {
    it("forwards text, genero, and autorSlug to start", async () => {
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();
      const chunk = makeChunk({
        text: "Hola mundo",
        index: 4,
        total: 9,
        startOffset: 999,
      });

      await adapter.submitChunkAnalysis(
        chunk,
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
    });

    it("returns the runId created by createRun when submit is acknowledged", async () => {
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.submitChunkAnalysis(
        makeChunk({ index: 7 }),
        "general",
        "Disble",
      );

      expect(result).toEqual({
        chunkIndex: 7,
        runId: "run-from-create-run",
      });
    });

    it("returns an error when createRun does not provide a usable runId", async () => {
      mastraMocks.createRun.mockResolvedValueOnce({ start: mastraMocks.start });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.submitChunkAnalysis(
        makeChunk(),
        "general",
        "Disble",
      );

      expect(result).toEqual({
        chunkIndex: 2,
        error: "Workflow createRun did not return a valid runId",
      });
      expect(mastraMocks.start).not.toHaveBeenCalled();
    });

    it("accepts submit acknowledgement payloads without validating message contents", async () => {
      mastraMocks.start.mockResolvedValueOnce(undefined);
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.submitChunkAnalysis(
        makeChunk(),
        "general",
        "Disble",
      );

      expect(result).toEqual({
        chunkIndex: 2,
        runId: "run-from-create-run",
      });
    });

    it("returns the thrown message when submit fails", async () => {
      mastraMocks.start.mockRejectedValueOnce(
        new Error("connect ECONNREFUSED"),
      );
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.submitChunkAnalysis(
        makeChunk({ index: 3 }),
        "general",
        "Disble",
      );

      expect(result).toEqual({
        chunkIndex: 3,
        error: "connect ECONNREFUSED",
      });
    });

    it("returns a bypass runId without calling createRun when the bypass is enabled", async () => {
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

      const result = await adapter.submitChunkAnalysis(
        makeChunk({ index: 4 }),
        "general",
        "Disble",
      );

      expect(mastraMocks.getWorkflow).not.toHaveBeenCalled();
      expect(mastraMocks.createRun).not.toHaveBeenCalled();
      expect(result).toEqual({
        chunkIndex: 4,
        runId: "bypass-run-4",
      });
    });
  });

  describe("pollChunkAnalysis", () => {
    it("calls workflow.runById with only explicit payload fields because status is implicit", async () => {
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      await adapter.pollChunkAnalysis(2, "run-123");

      expect(mastraMocks.getWorkflow).toHaveBeenCalledWith(WORKFLOW_ID);
      expect(mastraMocks.runById).toHaveBeenCalledWith("run-123", {
        fields: ["result", "error"],
        withNestedWorkflows: false,
      });
    });

    it("returns mocked suggestions without calling runById when the poll bypass is enabled", async () => {
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

      const result = await adapter.pollChunkAnalysis(2, "run-bypass");

      expect(mastraMocks.runById).not.toHaveBeenCalled();
      expect(result).toEqual({
        chunkIndex: 2,
        runId: "run-bypass",
        status: "success",
        suggestions: [
          {
            id: "chunk2-0",
            context:
              "Xia no tenía idea de lo que estaba pasando por la mente de Mei ni Shu.",
            anchor: "ni Shu",
            suggestedText: "ni de Shu",
            justification:
              "Omisión de preposición en estructuras paralelas.",
            category: "gramática",
            severity: "high",
            type: "track-change",
          },
          {
            id: "chunk2-1",
            context:
              "Cualquier cosa que pudiera decir eran meras suposiciones.",
            anchor: "eran meras suposiciones",
            suggestedText: "era una mera suposición",
            justification: "Concordancia verbal con sujeto singular.",
            category: "gramática",
            severity: "high",
            type: "track-change",
          },
          {
            id: "chunk2-2",
            context: "Todas se quedaron en silencio.",
            anchor: "se quedaron en silencio",
            suggestedText: "guardaron silencio",
            justification:
              "Eco léxico por repetición de 'silencio' en párrafos cercanos.",
            category: "estilo",
            severity: "medium",
            type: "track-change",
          },
          {
            id: "chunk2-3",
            context:
              "desperté en unas instalaciones de WEPO lejos de las garras de Jack y desde allí no volví a saber de él.",
            anchor: "desde allí",
            justification:
              "Evaluar si se refiere a la ubicación (correcto) o al momento temporal (debería ser 'desde entonces').",
            category: "estilo",
            severity: "low",
            type: "comment-only",
          },
        ],
      });
    });

    it("returns failed when workflow enters suspended state", async () => {
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

    it("returns failed when workflow enters paused state", async () => {
      mastraMocks.runById.mockResolvedValueOnce({ status: "paused" });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.pollChunkAnalysis(2, "run-2");

      expect(result).toEqual({
        chunkIndex: 2,
        runId: "run-2",
        status: "failed",
        suggestions: [],
        error:
          'Workflow entered "paused" state and requires resume(), which this frontend does not support',
      });
    });

    it("rejects a workflow suggestion missing anchor", async () => {
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

      const result = await adapter.pollChunkAnalysis(3, "run-3");

      expect(result).toEqual({
        chunkIndex: 3,
        runId: "run-3",
        status: "failed",
        suggestions: [],
        error: "Invalid workflow success payload: expected suggestions[]",
      });
    });

    it("rejects a workflow suggestion whose anchor is not part of context", async () => {
      mastraMocks.runById.mockResolvedValueOnce({
        status: "success",
        result: {
          suggestions: [
            {
              context: "Texto observado dentro del párrafo.",
              anchor: "otro fragmento",
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

    it("maps a comment-only workflow suggestion with no suggestedText", async () => {
      const suggestions: WorkflowSuggestion[] = [
        {
          context: "Se encontró el texto observado dentro del párrafo.",
          anchor: "texto observado",
          justification: "Revisar tono",
          category: "Tono",
          severity: "low",
          type: "comment-only",
        },
      ];
      mastraMocks.runById.mockResolvedValueOnce({
        status: "success",
        result: { suggestions },
      });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.pollChunkAnalysis(3, "run-3");

      expect(result.status).toBe("success");
      expect(result.suggestions).toHaveLength(1);
      expect(result.suggestions[0]).toEqual({
        id: "chunk3-0",
        context: "Se encontró el texto observado dentro del párrafo.",
        anchor: "texto observado",
        justification: "Revisar tono",
        category: "Tono",
        severity: "low",
        type: "comment-only",
      });
      // suggestedText must NOT be present for comment-only suggestions
      expect(result.suggestions[0]).not.toHaveProperty("suggestedText");
    });

    it("maps successful workflow suggestions into domain suggestions with generated ids", async () => {
      const suggestions: WorkflowSuggestion[] = [
        {
          context: "La frase muy muy aparece duplicada en este contexto.",
          anchor: "muy muy",
          suggestedText: "muy",
          justification: "Redundancia",
          category: "Estilo",
          severity: "medium",
        },
        {
          context: "Podés reemplazar en este momento por algo más directo.",
          anchor: "en este momento",
          suggestedText: "ahora",
          justification: "Mas directo",
          category: "Claridad",
          severity: "low",
        },
      ];
      mastraMocks.runById.mockResolvedValueOnce({
        status: "success",
        result: { suggestions, warnings: ["ignored"] },
      });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.pollChunkAnalysis(7, "run-7");

      expect(result).toEqual({
        chunkIndex: 7,
        runId: "run-7",
        status: "success",
        suggestions: [
          {
            id: "chunk7-0",
            context: "La frase muy muy aparece duplicada en este contexto.",
            anchor: "muy muy",
            suggestedText: "muy",
            justification: "Redundancia",
            category: "Estilo",
            severity: "medium",
            type: "track-change",
          },
          {
            id: "chunk7-1",
            context: "Podés reemplazar en este momento por algo más directo.",
            anchor: "en este momento",
            suggestedText: "ahora",
            justification: "Mas directo",
            category: "Claridad",
            severity: "low",
            type: "track-change",
          },
        ],
      });
    });

    it("returns running without treating it as an error", async () => {
      mastraMocks.runById.mockResolvedValueOnce({ status: "running" });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.pollChunkAnalysis(2, "run-2");

      expect(result).toEqual({
        chunkIndex: 2,
        runId: "run-2",
        status: "running",
        suggestions: [],
      });
    });

    it("returns failed when a success payload is malformed", async () => {
      mastraMocks.runById.mockResolvedValueOnce({
        status: "success",
        result: "bad payload",
      });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.pollChunkAnalysis(2, "run-2");

      expect(result).toEqual({
        chunkIndex: 2,
        runId: "run-2",
        status: "failed",
        suggestions: [],
        error: "Invalid workflow success payload: expected suggestions[]",
      });
    });

    it("returns failed when poll payload has no status", async () => {
      mastraMocks.runById.mockResolvedValueOnce({
        result: { suggestions: [] },
      });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.pollChunkAnalysis(2, "run-2");

      expect(result).toEqual({
        chunkIndex: 2,
        runId: "run-2",
        status: "failed",
        suggestions: [],
        error: "Invalid workflow poll payload: missing status",
      });
    });

    it("extracts serialized workflow errors for terminal failures", async () => {
      mastraMocks.runById.mockResolvedValueOnce({
        status: "failed",
        error: { message: "workflow exploded" },
      });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.pollChunkAnalysis(0, "run-0");

      expect(result).toEqual({
        chunkIndex: 0,
        runId: "run-0",
        status: "failed",
        suggestions: [],
        error: "workflow exploded",
      });
    });

    it("returns failed with explicit message for unknown statuses", async () => {
      mastraMocks.runById.mockResolvedValueOnce({ status: "completed" });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.pollChunkAnalysis(0, "run-0");

      expect(result).toEqual({
        chunkIndex: 0,
        runId: "run-0",
        status: "failed",
        suggestions: [],
        error: "Unknown workflow status: completed",
      });
    });

    it("returns fallback error when terminal status has no error payload", async () => {
      mastraMocks.runById.mockResolvedValueOnce({ status: "canceled" });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.pollChunkAnalysis(1, "run-1");

      expect(result).toEqual({
        chunkIndex: 1,
        runId: "run-1",
        status: "canceled",
        suggestions: [],
        error:
          'Workflow terminated with status "canceled" without an error payload',
      });
    });

    it("serializes non-standard error payloads for terminal failures", async () => {
      mastraMocks.runById.mockResolvedValueOnce({
        status: "failed",
        error: { code: "E_TIMEOUT", detail: "upstream timeout" },
      });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.pollChunkAnalysis(5, "run-5");

      expect(result).toEqual({
        chunkIndex: 5,
        runId: "run-5",
        status: "failed",
        suggestions: [],
        error: '{"code":"E_TIMEOUT","detail":"upstream timeout"}',
      });
    });
  });
});
