import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  createOffice,
  createTaskpaneDocument,
  deferred,
  getTaskpaneMocks,
  importTaskpane,
  makeSuggestion,
  resetTaskpaneHarness,
  teardownTaskpaneHarness,
} from "./TaskpaneTestHelper";
import {
  DEFAULT_MAX_CHUNK_SIZE,
  MAX_RETRIES,
  RETRY_BASE_DELAY_MS,
} from "../infrastructure/config";

describe("taskpane entrypoint", () => {
  const taskpaneMocks = getTaskpaneMocks();
  let logSpy: ReturnType<typeof vi.spyOn>;
  let warnSpy: ReturnType<typeof vi.spyOn>;
  let errorSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    resetTaskpaneHarness();
    logSpy = vi.spyOn(console, "log").mockImplementation(() => {});
    warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    errorSpy = vi.spyOn(console, "error").mockImplementation(() => {});
  });

  afterEach(() => {
    logSpy.mockRestore();
    warnSpy.mockRestore();
    errorSpy.mockRestore();
    teardownTaskpaneHarness();
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
      RETRY_BASE_DELAY_MS,
    );

    officeHarness.triggerReady({ host: "Excel" });
    expect(doc.getElementById("sideload-msg")!.style.display).toBe("block");
    expect(doc.getElementById("btn-analyze")?.onclick).toBeNull();

    taskpaneMocks.getCleanupPreview.mockResolvedValueOnce({
      deletable: 1,
      kept: 0,
    });

    officeHarness.triggerReady({ host: "Word" });
    await Promise.resolve();
    expect(doc.getElementById("sideload-msg")!.style.display).toBe("none");
    expect(doc.getElementById("app-body")!.style.display).toBe("flex");
    expect(doc.getElementById("btn-analyze")?.onclick).toEqual(
      expect.any(Function),
    );
    expect(doc.getElementById("btn-cleanup")?.onclick).toEqual(
      expect.any(Function),
    );
    expect(taskpaneMocks.getCleanupPreview).toHaveBeenCalledOnce();
    expect(doc.getElementById("cleanup-section")!.style.display).toBe("block");
  });

  it("runs the pipeline with the selected profile and updates terminal UI state from emitted events", async () => {
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
        true,
      );
    });
    taskpaneMocks.getCleanupPreview.mockResolvedValueOnce({
      deletable: 0,
      kept: 0,
    });
    taskpaneMocks.getCleanupPreview.mockResolvedValueOnce({
      deletable: 1,
      kept: 0,
    });

    await importTaskpane();
    officeHarness.triggerReady({ host: "Word" });
    await Promise.resolve();

    await doc.getElementById("btn-analyze")?.onclick?.({} as MouseEvent);
    await Promise.resolve();

    expect(taskpaneMocks.run).toHaveBeenCalledOnce();
    expect(taskpaneMocks.run.mock.calls[0][0]).toMatchObject({
      genero: "narrativa-literaria",
      maxChunkSize: DEFAULT_MAX_CHUNK_SIZE,
      documentPort: expect.any(Object),
      analysisPort: expect.any(Object),
      emitter: expect.any(Object),
    });
    expect(doc.getElementById("progress-container")!.style.display).toBe("block");
    expect(doc.getElementById("progress-bar")?.style.width).toBe("25%");
    expect(doc.getElementById("progress-text")?.textContent).toBe(
      "Analizando fragmentos...",
    );
    expect(doc.getElementById("results-summary")?.textContent).toBe(
      "Sobre selección — 1 de 1 sugerencias aplicadas como Track Changes.",
    );
    expect(doc.getElementById("results-list")?.children).toHaveLength(1);
    expect(doc.getElementById("cleanup-section")!.style.display).toBe("block");
    expect(doc.getElementById("status-bar")?.textContent).toBe(
      "1 sugerencia(s) insertada(s) como Track Changes (selección).",
    );
    expect(doc.getElementById("btn-analyze")?.disabled).toBe(false);
    expect(taskpaneMocks.getCleanupPreview).toHaveBeenCalledTimes(2);

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

    const firstRun = doc.getElementById("btn-analyze")?.onclick?.({} as MouseEvent);
    await Promise.resolve();
    await doc.getElementById("btn-analyze")?.onclick?.({} as MouseEvent);

    expect(taskpaneMocks.run).toHaveBeenCalledOnce();
    expect(warnSpy).toHaveBeenCalledWith(
      "⚠️ [Taskpane] Pipeline ya en ejecución — ignorando click",
    );

    runDeferred.resolve();
    await firstRun;
  });

  it("delegates cleanup to the document adapter and hides the section when no comments remain", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;
    doc.getElementById("cleanup-section")!.style.display = "block";

    taskpaneMocks.cleanupResolvedComments.mockResolvedValueOnce({
      deleted: 2,
      kept: 0,
    });

    await importTaskpane();
    officeHarness.triggerReady({ host: "Word" });

    await doc.getElementById("btn-cleanup")?.onclick?.({} as MouseEvent);

    expect(taskpaneMocks.cleanupResolvedComments).toHaveBeenCalledOnce();
    expect(doc.getElementById("cleanup-section")!.style.display).toBe("none");
    expect(doc.getElementById("status-bar")?.textContent).toBe(
      "2 comentario(s) eliminado(s), 0 conservado(s).",
    );
  });

  it("keeps the cleanup section visible on bootstrap when deletable comments already exist", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;
    taskpaneMocks.getCleanupPreview.mockResolvedValueOnce({
      deletable: 3,
      kept: 1,
    });

    await importTaskpane();
    officeHarness.triggerReady({ host: "Word" });
    await Promise.resolve();

    expect(doc.getElementById("cleanup-section")!.style.display).toBe("block");
    expect(taskpaneMocks.getCleanupPreview).toHaveBeenCalledOnce();
  });

  it("hides the cleanup section after analysis when no deletable comments remain", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;
    doc.getElementById("cleanup-section")!.style.display = "block";

    taskpaneMocks.run.mockImplementationOnce(async (ctx) => {
      ctx.emitter.emitComplete(
        [makeSuggestion()],
        { successCount: 1, failedSuggestions: [] },
        [],
        false,
      );
    });
    taskpaneMocks.getCleanupPreview.mockResolvedValueOnce({
      deletable: 1,
      kept: 0,
    });
    taskpaneMocks.getCleanupPreview.mockResolvedValueOnce({
      deletable: 0,
      kept: 1,
    });

    await importTaskpane();
    officeHarness.triggerReady({ host: "Word" });
    await Promise.resolve();

    await doc.getElementById("btn-analyze")?.onclick?.({} as MouseEvent);
    await Promise.resolve();

    expect(doc.getElementById("cleanup-section")!.style.display).toBe("none");
  });
});
