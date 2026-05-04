import type * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { AppProps } from "./components/App";
import { DEFAULT_MAX_CHUNK_SIZE, MAX_RETRIES, RETRY_BASE_DELAY_MS } from "../infrastructure/config";
import {
  createOffice,
  createTaskpaneDocument,
  deferred,
  getTaskpaneMocks,
  getTaskpaneReactMocks,
  importTaskpane,
  makeSuggestion,
  resetTaskpaneHarness,
  teardownTaskpaneHarness,
} from "./TaskpaneTestHelper";

function getRequiredElement(doc: ReturnType<typeof createTaskpaneDocument>, id: string) {
  const el = doc.getElementById(id);
  if (!el) {
    throw new Error(`Missing fake DOM element: ${id}`);
  }
  return el;
}

function getRenderedAppProps(reactMocks: ReturnType<typeof getTaskpaneReactMocks>): AppProps {
  const lastRenderCall = reactMocks.render.mock.calls[reactMocks.render.mock.calls.length - 1];
  const renderedTree = lastRenderCall?.[0] as
    | React.ReactElement<{ children?: React.ReactNode }>
    | undefined;
  const appElement = renderedTree?.props.children as React.ReactElement<AppProps> | undefined;

  if (!appElement) {
    throw new Error("Missing rendered App element in React mock tree");
  }

  return appElement.props;
}

async function importTaskpaneRuntime() {
  await importTaskpane();

  const resultsPanelStore = await import("./ResultsPanelStore");
  const taskpaneShellStore = await import("./TaskpaneShellStore");

  return {
    getResultsPanelState: resultsPanelStore.getResultsPanelState,
    resetResultsPanelState: resultsPanelStore.resetResultsPanelState,
    getTaskpaneShellState: taskpaneShellStore.getTaskpaneShellState,
    resetTaskpaneShellState: taskpaneShellStore.resetTaskpaneShellState,
  };
}

async function flushTaskpaneWork(times = 8) {
  for (let index = 0; index < times; index += 1) {
    await Promise.resolve();
  }
}

describe("taskpane entrypoint", () => {
  const taskpaneMocks = getTaskpaneMocks();
  const reactMocks = getTaskpaneReactMocks();
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

    const { getTaskpaneShellState } = await importTaskpaneRuntime();

    expect(reactMocks.createRoot).toHaveBeenCalledWith(getRequiredElement(doc, "container"));

    officeHarness.triggerReady({ host: "Excel" });
    await flushTaskpaneWork();

    expect(reactMocks.render).not.toHaveBeenCalled();

    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();

    const appProps = getRenderedAppProps(reactMocks);

    expect(reactMocks.render).toHaveBeenCalledOnce();
    expect(taskpaneMocks.orchestratorHandlers).toHaveLength(7);
    expect(taskpaneMocks.retryDecoratorConstructor).toHaveBeenCalledWith(
      expect.any(Object),
      MAX_RETRIES,
      RETRY_BASE_DELAY_MS
    );
    expect(appProps.onAnalyze).toEqual(expect.any(Function));
    expect(appProps.onCleanup).toEqual(expect.any(Function));
    expect(appProps.onDisableTrackChanges).toEqual(expect.any(Function));
    expect(getTaskpaneShellState().cleanupVisible).toBe(true);
    expect(getTaskpaneShellState().disableTrackChangesCtaVisible).toBe(true);
  });

  it("runs the pipeline and publishes progress/results/state through the stores", async () => {
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
        true
      );
    });
    taskpaneMocks.getCleanupPreview.mockResolvedValueOnce({ deletable: 0, kept: 0 });
    taskpaneMocks.getCleanupPreview.mockResolvedValueOnce({ deletable: 1, kept: 0 });

    const { getResultsPanelState, getTaskpaneShellState } = await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();

    const appProps = getRenderedAppProps(reactMocks);

    await appProps.onAnalyze();
    await flushTaskpaneWork();
    vi.advanceTimersByTime(1000);
    await flushTaskpaneWork();

    expect(taskpaneMocks.run).toHaveBeenCalledOnce();
    expect(taskpaneMocks.run.mock.calls[0][0]).toMatchObject({
      genero: "narrativa-literaria",
      maxChunkSize: DEFAULT_MAX_CHUNK_SIZE,
      documentPort: expect.any(Object),
      analysisPort: expect.any(Object),
      emitter: expect.any(Object),
    });
    expect(getTaskpaneShellState().progress.visible).toBe(false);
    expect(getResultsPanelState().visible).toBe(true);
    expect(getResultsPanelState().cards).toHaveLength(1);
    expect(getTaskpaneShellState().status.message).toBe(
      "1 sugerencia(s) insertada(s) como Track Changes (selección)."
    );
    expect(getTaskpaneShellState().cleanupVisible).toBe(true);
  });

  it("ignores analyze clicks while a run is already in progress", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    const runDeferred = deferred<void>();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    taskpaneMocks.run.mockImplementationOnce(() => runDeferred.promise);

    await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();

    const appProps = getRenderedAppProps(reactMocks);

    const firstRun = appProps.onAnalyze();
    await Promise.resolve();
    await appProps.onAnalyze();

    expect(taskpaneMocks.run).toHaveBeenCalledOnce();
    expect(warnSpy).toHaveBeenCalledWith(
      "⚠️ [Taskpane] Pipeline ya en ejecución — ignorando click"
    );

    runDeferred.resolve();
    await firstRun;
  });

  it("delegates cleanup and hides the cleanup CTA when no comments remain", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    taskpaneMocks.cleanupResolvedComments.mockResolvedValueOnce({ deleted: 2, kept: 0 });

    const { getTaskpaneShellState } = await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();

    const appProps = getRenderedAppProps(reactMocks);

    await appProps.onCleanup();

    expect(taskpaneMocks.cleanupResolvedComments).toHaveBeenCalledOnce();
    expect(getTaskpaneShellState().cleanupVisible).toBe(false);
    expect(getTaskpaneShellState().status.message).toBe(
      "2 comentario(s) eliminado(s), 0 conservado(s)."
    );
  });

  it("delegates disabling Track Changes and hides the CTA afterward", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    const { getTaskpaneShellState } = await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();

    const appProps = getRenderedAppProps(reactMocks);

    await appProps.onDisableTrackChanges();

    expect(taskpaneMocks.disableTrackChanges).toHaveBeenCalledOnce();
    expect(getTaskpaneShellState().disableTrackChangesCtaVisible).toBe(false);
    expect(getTaskpaneShellState().status.message).toBe("Control de cambios desactivado.");
  });
});
