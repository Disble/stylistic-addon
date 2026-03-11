import { MASTRA_BASE_URL, WORKFLOW_ID } from "../../infrastructure/config";
import type { TextChunk, WorkflowSuggestion } from "../../domain/types";

const mastraMocks = vi.hoisted(() => {
  const details = vi.fn();
  const createRun = vi.fn();
  const startAsync = vi.fn();
  const getWorkflow = vi.fn();
  const constructor = vi.fn();

  return {
    details,
    createRun,
    startAsync,
    getWorkflow,
    constructor,
    workflow: {
      details,
      createRun,
    },
    run: {
      startAsync,
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
    vi.clearAllMocks();

    mastraMocks.getWorkflow.mockReturnValue(mastraMocks.workflow);
    mastraMocks.details.mockResolvedValue({ id: WORKFLOW_ID });
    mastraMocks.createRun.mockResolvedValue(mastraMocks.run);
    mastraMocks.startAsync.mockResolvedValue({
      status: "success",
      result: { suggestions: [] },
    });

    logSpy = vi.spyOn(console, "log").mockImplementation(() => {});
    errorSpy = vi.spyOn(console, "error").mockImplementation(() => {});
  });

  afterEach(() => {
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
  });

  describe("analyzeChunk", () => {
    it("forwards text, profile, and language to startAsync", async () => {
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();
      const chunk = makeChunk({ text: "Hola mundo", index: 4, total: 9, startOffset: 999 });

      await adapter.analyzeChunk(chunk, "formal", "es");

      expect(mastraMocks.getWorkflow).toHaveBeenCalledWith(WORKFLOW_ID);
      expect(mastraMocks.createRun).toHaveBeenCalledOnce();
      expect(mastraMocks.startAsync).toHaveBeenCalledWith({
        inputData: {
          text: "Hola mundo",
          profile: "formal",
          language: "es",
        },
      });
    });

    it("maps successful workflow suggestions into domain suggestions with generated ids", async () => {
      const suggestions: WorkflowSuggestion[] = [
        {
          originalText: "muy muy",
          suggestedText: "muy",
          justification: "Redundancia",
          category: "Estilo",
          severity: "medium",
        },
        {
          originalText: "en este momento",
          suggestedText: "ahora",
          justification: "Mas directo",
          category: "Claridad",
          severity: "low",
        },
      ];
      mastraMocks.startAsync.mockResolvedValueOnce({
        status: "success",
        result: { suggestions, warnings: ["ignored"] },
      });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.analyzeChunk(makeChunk({ index: 7 }), "general", "es");

      expect(result).toEqual({
        chunkIndex: 7,
        suggestions: [
          {
            id: "chunk7-0",
            originalText: "muy muy",
            suggestedText: "muy",
            justification: "Redundancia",
            category: "Estilo",
            severity: "medium",
          },
          {
            id: "chunk7-1",
            originalText: "en este momento",
            suggestedText: "ahora",
            justification: "Mas directo",
            category: "Claridad",
            severity: "low",
          },
        ],
      });
    });

    it("returns empty suggestions on a successful response with an empty suggestions array", async () => {
      mastraMocks.startAsync.mockResolvedValueOnce({
        status: "success",
        result: { suggestions: [] },
      });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.analyzeChunk(makeChunk(), "general", "es");

      expect(result).toEqual({ chunkIndex: 2, suggestions: [] });
    });

    it("intentionally normalizes a non-array suggestions payload to an empty suggestions list", async () => {
      mastraMocks.startAsync.mockResolvedValueOnce({
        status: "success",
        result: { suggestions: "bad payload" },
      });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.analyzeChunk(makeChunk(), "general", "es");

      expect(result).toEqual({ chunkIndex: 2, suggestions: [] });
    });

    it("returns a normalized error when a success payload has a non-object result", async () => {
      mastraMocks.startAsync.mockResolvedValueOnce({
        status: "success",
        result: "bad payload",
      });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.analyzeChunk(makeChunk(), "general", "es");

      expect(result).toEqual({
        chunkIndex: 2,
        suggestions: [],
        error: "Invalid workflow success payload",
      });
    });

    it("returns a normalized error when a success payload omits suggestions", async () => {
      mastraMocks.startAsync.mockResolvedValueOnce({
        status: "success",
        result: { warnings: ["missing suggestions"] },
      });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.analyzeChunk(makeChunk(), "general", "es");

      expect(result).toEqual({
        chunkIndex: 2,
        suggestions: [],
        error: "Invalid workflow success payload",
      });
    });

    it("returns a workflow status error for non-success results", async () => {
      mastraMocks.startAsync.mockResolvedValueOnce({
        status: "failed",
        result: { suggestions: [] },
      });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.analyzeChunk(makeChunk({ index: 0 }), "general", "es");

      expect(result).toEqual({
        chunkIndex: 0,
        suggestions: [],
        error: "Workflow status: failed",
      });
    });

    it("returns the thrown Error message when createRun or startAsync rejects", async () => {
      mastraMocks.createRun.mockRejectedValueOnce(new Error("connect ECONNREFUSED"));
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.analyzeChunk(makeChunk({ index: 3 }), "general", "es");

      expect(result).toEqual({
        chunkIndex: 3,
        suggestions: [],
        error: "connect ECONNREFUSED",
      });
    });

    it("stringifies non-Error thrown values", async () => {
      mastraMocks.startAsync.mockRejectedValueOnce("service unavailable");
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.analyzeChunk(makeChunk({ index: 1 }), "general", "es");

      expect(result).toEqual({
        chunkIndex: 1,
        suggestions: [],
        error: "service unavailable",
      });
    });

    it("returns a normalized error when a success payload omits result entirely", async () => {
      mastraMocks.startAsync.mockResolvedValueOnce({ status: "success" });
      const { MastraAdapter } = await importAdapterModule();
      const adapter = new MastraAdapter();

      const result = await adapter.analyzeChunk(makeChunk(), "general", "es");

      expect(result).toEqual({
        chunkIndex: 2,
        suggestions: [],
        error: "Invalid workflow success payload",
      });
    });
  });
});
