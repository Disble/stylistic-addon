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

function getRequiredElement(doc: ReturnType<typeof createTaskpaneDocument>, id: string) {
  const el = doc.getElementById(id);
  if (!el) {
    throw new Error(`Missing fake DOM element: ${id}`);
  }
  return el;
}

async function flushTaskpaneWork(times = 6) {
  for (let index = 0; index < times; index += 1) {
    await Promise.resolve();
  }
}

describe("taskpane entrypoint", () => {
  const taskpaneMocks = getTaskpaneMocks();
  let warnSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    resetTaskpaneHarness();
    warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
  });

  afterEach(() => {
    warnSpy.mockRestore();
    teardownTaskpaneHarness();
  });

  it("boots only for Word and wires top-level handlers plus initial CTA visibility", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    taskpaneMocks.getCleanupPreview.mockResolvedValueOnce({ deletable: 1, kept: 0 });
    taskpaneMocks.getDocumentReviewState.mockResolvedValueOnce({
      pendingStylisticArtifacts: 0,
      hasPendingStylisticArtifacts: false,
      trackChangesActive: true,
    });

    await importTaskpane();

    officeHarness.triggerReady({ host: "Excel" });
    expect(getRequiredElement(doc, "sideload-msg").style.display).toBe("block");
    expect(doc.getElementById("btn-analyze")?.onclick).toBeNull();

    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();

    expect(taskpaneMocks.orchestratorHandlers).toHaveLength(7);
    expect(taskpaneMocks.retryDecoratorConstructor).toHaveBeenCalledWith(
      expect.any(Object),
      MAX_RETRIES,
      RETRY_BASE_DELAY_MS,
    );
    expect(getRequiredElement(doc, "sideload-msg").style.display).toBe("none");
    expect(getRequiredElement(doc, "app-body").style.display).toBe("flex");
    expect(doc.getElementById("btn-analyze")?.onclick).toEqual(expect.any(Function));
    expect(doc.getElementById("btn-cleanup")?.onclick).toEqual(expect.any(Function));
    expect(doc.getElementById("btn-disable-track-changes")?.onclick).toEqual(
      expect.any(Function),
    );
    expect(getRequiredElement(doc, "cleanup-section").style.display).toBe("block");
    expect(getRequiredElement(doc, "disable-track-changes-section").style.display).toBe(
      "block",
    );
  });

  it("runs the pipeline and updates progress/results/status from emitted events", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    taskpaneMocks.run.mockImplementationOnce(async (ctx) => {
      ctx.emitter.emitPhaseStart("reading", "Leyendo documento...");
      ctx.emitter.emitProgress(1, 4, "Analizando fragmentos...");
      ctx.emitter.emitComplete(
        [makeSuggestion()],
        {
          successCount: 1,
          failedSuggestions: [],
          pendingAfter: {
            pendingStylisticArtifacts: 1,
            hasPendingStylisticArtifacts: true,
            trackChangesActive: true,
          },
          documentState: "pending-review",
          trackChangesActivatedForBatch: false,
        },
        [],
        true,
      );
    });
    taskpaneMocks.getCleanupPreview.mockResolvedValueOnce({ deletable: 0, kept: 0 });
    taskpaneMocks.getCleanupPreview.mockResolvedValueOnce({ deletable: 1, kept: 0 });

    await importTaskpane();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();

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
    expect(getRequiredElement(doc, "progress-container").style.display).toBe("block");
    expect(doc.getElementById("progress-bar")?.style.width).toBe("25%");
    expect(doc.getElementById("results-list")?.children).toHaveLength(1);
    expect(doc.getElementById("status-bar")?.textContent).toBe(
      "1 sugerencia(s) insertada(s) como Track Changes (selección).",
    );
    expect(getRequiredElement(doc, "cleanup-section").style.display).toBe("block");
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

  it("delegates cleanup and hides the section when no comments remain", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;
    getRequiredElement(doc, "cleanup-section").style.display = "block";

    taskpaneMocks.cleanupResolvedComments.mockResolvedValueOnce({ deleted: 2, kept: 0 });

    await importTaskpane();
    officeHarness.triggerReady({ host: "Word" });

    await doc.getElementById("btn-cleanup")?.onclick?.({} as MouseEvent);

    expect(taskpaneMocks.cleanupResolvedComments).toHaveBeenCalledOnce();
    expect(getRequiredElement(doc, "cleanup-section").style.display).toBe("none");
    expect(doc.getElementById("status-bar")?.textContent).toBe(
      "2 comentario(s) eliminado(s), 0 conservado(s).",
    );
  });

  it("delegates disabling Track Changes and hides the CTA afterward", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;
    getRequiredElement(doc, "disable-track-changes-section").style.display = "block";

    await importTaskpane();
    officeHarness.triggerReady({ host: "Word" });

    await doc.getElementById("btn-disable-track-changes")?.onclick?.({} as MouseEvent);

    expect(taskpaneMocks.disableTrackChanges).toHaveBeenCalledOnce();
    expect(getRequiredElement(doc, "disable-track-changes-section").style.display).toBe(
      "none",
    );
    expect(doc.getElementById("status-bar")?.textContent).toBe(
      "Control de cambios desactivado.",
    );
  });
});
