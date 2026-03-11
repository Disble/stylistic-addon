import { DEFAULT_MAX_CHUNK_SIZE, MAX_RETRIES, RETRY_BASE_DELAY_MS } from "../infrastructure/config";
import type { InsertionResult, Suggestion } from "../domain/types";

const taskpaneMocks = vi.hoisted(() => ({
  orchestratorHandlers: [] as unknown[],
  run: vi.fn<(ctx: any) => Promise<void>>(),
  wordAdapterConstructor: vi.fn(),
  cleanupResolvedComments: vi.fn<() => Promise<{ deleted: number; kept: number }>>(),
  mastraAdapterConstructor: vi.fn(),
  retryDecoratorConstructor: vi.fn(),
}));

vi.mock("../adapters/word/WordAdapter", () => ({
  WordAdapter: class {
    constructor() {
      taskpaneMocks.wordAdapterConstructor();
    }

    cleanupResolvedComments() {
      return taskpaneMocks.cleanupResolvedComments();
    }
  },
}));

vi.mock("../adapters/mastra/MastraAdapter", () => ({
  MastraAdapter: class {
    constructor() {
      taskpaneMocks.mastraAdapterConstructor();
    }
  },
}));

vi.mock("../adapters/RetryAnalysisDecorator", () => ({
  RetryAnalysisDecorator: class {
    constructor(...args: unknown[]) {
      taskpaneMocks.retryDecoratorConstructor(...args);
    }
  },
}));

vi.mock("../domain/pipeline/PipelineOrchestrator", () => ({
  PipelineOrchestrator: class {
    constructor(handlers: unknown[]) {
      taskpaneMocks.orchestratorHandlers = handlers;
    }

    run(ctx: any) {
      return taskpaneMocks.run(ctx);
    }
  },
}));

function makeSuggestion(overrides: Partial<Suggestion> = {}): Suggestion {
  return {
    id: "s-1",
    originalText: "texto original",
    suggestedText: "texto sugerido",
    justification: "Mas claro",
    category: "Claridad",
    severity: "medium",
    ...overrides,
  };
}

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

class FakeElement {
  style = { display: "", width: "" };
  className = "";
  disabled = false;
  value = "";
  onclick: ((ev: MouseEvent) => any) | null = null;
  children: FakeElement[] = [];

  private text = "";
  private html = "";

  get textContent(): string {
    return this.text;
  }

  set textContent(value: string) {
    this.text = value ?? "";
    this.html = escapeHtml(this.text);
  }

  get innerHTML(): string {
    return this.html;
  }

  set innerHTML(value: string) {
    this.html = value;
  }

  appendChild(child: FakeElement): FakeElement {
    this.children.push(child);
    return child;
  }
}

class FakeDocument {
  private readonly elements = new Map<string, FakeElement>();

  constructor(ids: string[]) {
    for (const id of ids) {
      this.elements.set(id, new FakeElement());
    }
  }

  getElementById(id: string): FakeElement | null {
    return this.elements.get(id) ?? null;
  }

  createElement(): FakeElement {
    return new FakeElement();
  }
}

function createTaskpaneDocument(): FakeDocument {
  const doc = new FakeDocument([
    "sideload-msg",
    "app-body",
    "btn-analyze",
    "btn-analyze-label",
    "btn-cleanup",
    "btn-cleanup-label",
    "profile-select",
    "status-bar",
    "progress-container",
    "progress-bar",
    "progress-text",
    "results-panel",
    "results-summary",
    "results-list",
    "cleanup-section",
  ]);

  doc.getElementById("sideload-msg")!.style.display = "block";
  doc.getElementById("app-body")!.style.display = "none";
  doc.getElementById("cleanup-section")!.style.display = "none";
  doc.getElementById("results-panel")!.style.display = "block";
  doc.getElementById("progress-container")!.style.display = "none";
  doc.getElementById("profile-select")!.value = "formal";
  doc.getElementById("btn-analyze-label")!.textContent = "Analizar y sugerir";
  doc.getElementById("btn-cleanup-label")!.textContent = "Limpiar comentarios resueltos";

  return doc;
}

type ReadyCallback = (info: { host: string }) => void;

function createOffice() {
  let readyCallback: ReadyCallback | undefined;
  const office = {
    HostType: { Word: "Word" },
    onReady: vi.fn((callback: ReadyCallback) => {
      readyCallback = callback;
      return Promise.resolve();
    }),
  };

  return {
    office,
    triggerReady(info: { host: string }) {
      if (!readyCallback) {
        throw new Error("Office.onReady callback was not registered");
      }
      readyCallback(info);
    },
  };
}

async function importTaskpane() {
  vi.resetModules();
  return import("./taskpane");
}

describe("taskpane entrypoint", () => {
  let logSpy: ReturnType<typeof vi.spyOn>;
  let warnSpy: ReturnType<typeof vi.spyOn>;
  let errorSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    vi.clearAllMocks();
    vi.useFakeTimers();
    taskpaneMocks.orchestratorHandlers = [];
    taskpaneMocks.run.mockResolvedValue(undefined);
    taskpaneMocks.cleanupResolvedComments.mockResolvedValue({ deleted: 0, kept: 0 });

    logSpy = vi.spyOn(console, "log").mockImplementation(() => {});
    warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    errorSpy = vi.spyOn(console, "error").mockImplementation(() => {});

    delete (globalThis as any).document;
    delete (globalThis as any).Office;
  });

  afterEach(() => {
    logSpy.mockRestore();
    warnSpy.mockRestore();
    errorSpy.mockRestore();
    vi.useRealTimers();
    delete (globalThis as any).document;
    delete (globalThis as any).Office;
  });

  it("does nothing when Office or the DOM is unavailable", async () => {
    await expect(importTaskpane()).resolves.toMatchObject({
      bootstrapTaskpane: expect.any(Function),
    });

    expect(taskpaneMocks.wordAdapterConstructor).toHaveBeenCalledOnce();
    expect(taskpaneMocks.mastraAdapterConstructor).toHaveBeenCalledOnce();
    expect(taskpaneMocks.retryDecoratorConstructor).toHaveBeenCalledTimes(1);
  });

  it("registers Office bootstrap and wires the taskpane only for Word", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    await importTaskpane();

    expect(officeHarness.office.onReady).toHaveBeenCalledOnce();
    expect(taskpaneMocks.orchestratorHandlers).toHaveLength(7);
    expect(taskpaneMocks.retryDecoratorConstructor).toHaveBeenCalledWith(
      expect.any(Object),
      MAX_RETRIES,
      RETRY_BASE_DELAY_MS
    );

    officeHarness.triggerReady({ host: "Excel" });
    expect(doc.getElementById("sideload-msg")!.style.display).toBe("block");
    expect(doc.getElementById("btn-analyze")!.onclick).toBeNull();

    officeHarness.triggerReady({ host: "Word" });
    expect(doc.getElementById("sideload-msg")!.style.display).toBe("none");
    expect(doc.getElementById("app-body")!.style.display).toBe("flex");
    expect(doc.getElementById("btn-analyze")!.onclick).toEqual(expect.any(Function));
    expect(doc.getElementById("btn-cleanup")!.onclick).toEqual(expect.any(Function));
  });

  it("runs the pipeline with the selected profile and updates the UI from emitted events", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    taskpaneMocks.run.mockImplementationOnce(async (ctx) => {
      ctx.emitter.emitPhaseStart("reading", "Leyendo documento...");
      ctx.emitter.emitProgress(1, 4, "Analizando fragmentos...");
      ctx.emitter.emitComplete(
        [makeSuggestion()],
        { successCount: 1, failedSuggestions: [] },
        [],
        true
      );
    });

    await importTaskpane();
    officeHarness.triggerReady({ host: "Word" });

    await doc.getElementById("btn-analyze")!.onclick?.({} as MouseEvent);

    expect(taskpaneMocks.run).toHaveBeenCalledOnce();
    expect(taskpaneMocks.run.mock.calls[0][0]).toMatchObject({
      profile: "formal",
      maxChunkSize: DEFAULT_MAX_CHUNK_SIZE,
      documentPort: expect.any(Object),
      analysisPort: expect.any(Object),
      emitter: expect.any(Object),
    });
    expect(doc.getElementById("progress-container")!.style.display).toBe("block");
    expect(doc.getElementById("progress-bar")!.style.display).toBe("");
    expect(doc.getElementById("progress-bar")!.style.width).toBe("25%");
    expect(doc.getElementById("progress-text")!.textContent).toBe("Analizando fragmentos...");
    expect(doc.getElementById("results-panel")!.style.display).toBe("block");
    expect(doc.getElementById("results-summary")!.textContent).toBe(
      "Sobre selección — 1 de 1 sugerencias aplicadas como Track Changes."
    );
    expect(doc.getElementById("results-list")!.children).toHaveLength(1);
    expect(doc.getElementById("cleanup-section")!.style.display).toBe("block");
    expect(doc.getElementById("status-bar")!.textContent).toBe(
      "1 sugerencia(s) insertada(s) como Track Changes (selección)."
    );
    expect(doc.getElementById("status-bar")!.className).toBe("stylistic-status success");
    expect(doc.getElementById("btn-analyze")!.disabled).toBe(false);
    expect(doc.getElementById("profile-select")!.disabled).toBe(false);
    expect(doc.getElementById("btn-analyze-label")!.textContent).toBe("Analizar y sugerir");

    await vi.advanceTimersByTimeAsync(1000);
    expect(doc.getElementById("progress-container")!.style.display).toBe("none");

    await vi.advanceTimersByTimeAsync(4000);
    expect(doc.getElementById("status-bar")!.style.display).toBe("none");
  });

  it("ignores analyze clicks while a run is already in progress", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    const runDeferred = deferred<void>();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    taskpaneMocks.run.mockImplementationOnce(() => runDeferred.promise);

    await importTaskpane();
    officeHarness.triggerReady({ host: "Word" });

    const firstRun = doc.getElementById("btn-analyze")!.onclick?.({} as MouseEvent);
    await Promise.resolve();
    await doc.getElementById("btn-analyze")!.onclick?.({} as MouseEvent);

    expect(taskpaneMocks.run).toHaveBeenCalledOnce();
    expect(warnSpy).toHaveBeenCalledWith("⚠️ [Taskpane] Pipeline ya en ejecución — ignorando click");

    runDeferred.resolve();
    await firstRun;
  });

  it("delegates cleanup to the document adapter and hides the section when no comments remain", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;
    doc.getElementById("cleanup-section")!.style.display = "block";

    taskpaneMocks.cleanupResolvedComments.mockResolvedValueOnce({ deleted: 2, kept: 0 });

    await importTaskpane();
    officeHarness.triggerReady({ host: "Word" });

    await doc.getElementById("btn-cleanup")!.onclick?.({} as MouseEvent);

    expect(taskpaneMocks.cleanupResolvedComments).toHaveBeenCalledOnce();
    expect(doc.getElementById("cleanup-section")!.style.display).toBe("none");
    expect(doc.getElementById("status-bar")!.textContent).toBe(
      "2 comentario(s) eliminado(s), 0 conservado(s)."
    );
    expect(doc.getElementById("btn-cleanup")!.disabled).toBe(false);
    expect(doc.getElementById("btn-cleanup-label")!.textContent).toBe(
      "Limpiar comentarios resueltos"
    );
  });
});
