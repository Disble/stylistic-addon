import type * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { AppProps } from "../components/App";
import {
  DEFAULT_MAX_CHUNK_SIZE,
  MAX_RETRIES,
  RETRY_BASE_DELAY_MS,
} from "../../infrastructure/config";
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
} from "../TaskpaneTestHelper";

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

  const resultsPanelStore = await import("../ResultsPanelStore");
  const taskpaneAuthStore = await import("../TaskpaneAuthStore");
  const taskpaneShellStore = await import("../TaskpaneShellStore");

  return {
    getResultsPanelState: resultsPanelStore.getResultsPanelState,
    resetResultsPanelState: resultsPanelStore.resetResultsPanelState,
    setTaskpaneAuthenticated: taskpaneAuthStore.setTaskpaneAuthenticated,
    getTaskpaneShellState: taskpaneShellStore.getTaskpaneShellState,
    resetTaskpaneShellState: taskpaneShellStore.resetTaskpaneShellState,
    setTaskpaneSelectedGenero: taskpaneShellStore.setTaskpaneSelectedGenero,
  };
}

const AUTHENTICATED_SESSION = {
  token: "token-de-prueba",
  user: {
    id: "u-1",
    email: "test@example.com",
  },
} as const;

type OfficeRuntimeStorageMock = {
  getItem: ReturnType<typeof vi.fn>;
  setItem: ReturnType<typeof vi.fn>;
  removeItem: ReturnType<typeof vi.fn>;
};

function installOfficeRuntime(storage?: OfficeRuntimeStorageMock): OfficeRuntimeStorageMock {
  const resolvedStorage =
    storage ??
    ({
      getItem: vi.fn().mockResolvedValue(null),
      setItem: vi.fn().mockResolvedValue(undefined),
      removeItem: vi.fn().mockResolvedValue(undefined),
    } satisfies OfficeRuntimeStorageMock);

  (
    globalThis as unknown as { OfficeRuntime?: { storage?: OfficeRuntimeStorageMock } }
  ).OfficeRuntime = storage ? { storage: resolvedStorage } : undefined;

  return resolvedStorage;
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
    expect(appProps.onRetryAnalysis).toEqual(expect.any(Function));
    expect(appProps.onCancelAnalysis).toEqual(expect.any(Function));
    expect(appProps.onRetryAnalysisQuery).toEqual(expect.any(Function));
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

    const { getResultsPanelState, getTaskpaneShellState, setTaskpaneAuthenticated } =
      await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();
    setTaskpaneAuthenticated(AUTHENTICATED_SESSION);

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

    const { setTaskpaneAuthenticated } = await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();
    setTaskpaneAuthenticated(AUTHENTICATED_SESSION);

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

  it("cancels backend runs when the progress session exposes active polling", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    const runDeferred = deferred<void>();
    taskpaneMocks.run.mockImplementationOnce(async (ctx) => {
      ctx.activeRunReferences = [{ chunkIndex: 0, runId: "run-0" }];
      ctx.emitter.emitProgress(1, 2, "Consultando resultado del fragmento 1 de 1...");
      await runDeferred.promise;
    });

    const { getTaskpaneShellState, setTaskpaneAuthenticated } = await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();
    setTaskpaneAuthenticated(AUTHENTICATED_SESSION);

    const appProps = getRenderedAppProps(reactMocks);
    const analyzePromise = appProps.onAnalyze();
    await flushTaskpaneWork();

    expect(getTaskpaneShellState().progress.asyncSession.activeRuns).toEqual([
      { chunkIndex: 0, runId: "run-0" },
    ]);

    await appProps.onCancelAnalysis();
    runDeferred.resolve();
    await analyzePromise;

    expect(taskpaneMocks.cancelChunkAnalysis).toHaveBeenCalledWith(0, "run-0");
    expect(getTaskpaneShellState().status.message).toBe("Análisis cancelado en backend.");
  });

  it("publishes active polling runs to the progress session before the pipeline finishes", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    const progressDeferred = deferred<void>();
    taskpaneMocks.run.mockImplementationOnce(async (ctx) => {
      ctx.activeRunReferences = [{ chunkIndex: 0, runId: "run-0" }];
      ctx.emitter.emitProgress(1, 2, "Consultando resultado del fragmento 1 de 1...");
      expect(ctx.activeRunReferences).toEqual([{ chunkIndex: 0, runId: "run-0" }]);
      expect(ctx.retryableRunReferences).toEqual([]);
      await progressDeferred.promise;
    });

    const { getTaskpaneShellState, setTaskpaneAuthenticated } = await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();
    setTaskpaneAuthenticated(AUTHENTICATED_SESSION);

    const appProps = getRenderedAppProps(reactMocks);
    const analyzePromise = appProps.onAnalyze();
    await flushTaskpaneWork();

    expect(getTaskpaneShellState().progress.asyncSession.phase).toBe("polling");
    expect(getTaskpaneShellState().progress.asyncSession.activeRuns).toEqual([
      { chunkIndex: 0, runId: "run-0" },
    ]);

    progressDeferred.resolve();
    await analyzePromise;
  });

  it("retries polling with the same runId without re-submitting chunks", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    taskpaneMocks.run.mockImplementationOnce(async (ctx) => {
      ctx.retryableRunReferences = [{ chunkIndex: 0, runId: "run-0" }];
      ctx.chunkErrors = ["Chunk 1: poll timeout"];
      ctx.emitter.emitProgress(1, 2, "Consultando resultado del fragmento 1 de 1...");
      ctx.emitter.emitAbort(
        "La consulta del análisis falló localmente. Reintentá la consulta con el mismo run."
      );
    });
    taskpaneMocks.retryPollChunkAnalysis.mockResolvedValueOnce({
      chunkIndex: 0,
      runId: "run-0",
      status: "success",
      origin: "backend",
      suggestions: [makeSuggestion()],
    });

    const { getTaskpaneShellState, getResultsPanelState, setTaskpaneAuthenticated } =
      await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();
    setTaskpaneAuthenticated(AUTHENTICATED_SESSION);

    const appProps = getRenderedAppProps(reactMocks);
    await appProps.onAnalyze();
    await flushTaskpaneWork();

    expect(getTaskpaneShellState().progress.asyncSession.retryableRuns).toEqual([
      { chunkIndex: 0, runId: "run-0" },
    ]);

    await appProps.onRetryAnalysisQuery();
    await flushTaskpaneWork();

    expect(taskpaneMocks.retryPollChunkAnalysis).toHaveBeenCalledWith({
      chunkIndex: 0,
      runId: "run-0",
    });
    expect(taskpaneMocks.run).toHaveBeenCalledOnce();
    expect(getResultsPanelState().visible).toBe(true);
  });

  it("publishes a retry-query error surface instead of returning to the idle hero", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    taskpaneMocks.run.mockImplementationOnce(async (ctx) => {
      ctx.retryableRunReferences = [{ chunkIndex: 0, runId: "run-0" }];
      ctx.chunkErrors = ["Chunk 1: poll timeout"];
      ctx.emitter.emitProgress(1, 2, "Consultando resultado del fragmento 1 de 1...");
      ctx.emitter.emitAbort(
        "La consulta del análisis falló localmente. Reintentá la consulta con el mismo run."
      );
    });

    const { getTaskpaneShellState, setTaskpaneAuthenticated } = await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();
    setTaskpaneAuthenticated(AUTHENTICATED_SESSION);

    const appProps = getRenderedAppProps(reactMocks);
    await appProps.onAnalyze();
    await flushTaskpaneWork();

    expect(getTaskpaneShellState().analysisError).toEqual({
      message: "La consulta del análisis falló localmente. Reintentá la consulta con el mismo run.",
      retryKind: "retry-query",
      visible: true,
    });
    expect(getTaskpaneShellState().progress.visible).toBe(false);
  });

  it("publishes a full-retry error surface for terminal backend failures", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    taskpaneMocks.run.mockImplementationOnce(async (ctx) => {
      ctx.chunkErrors = ["Chunk 1: backend exploded"];
      ctx.emitter.emitProgress(1, 2, "Consultando resultado del fragmento 1 de 1...");
      ctx.emitter.emitAbort("El análisis falló en 1 fragmento(s). Intenta de nuevo.");
    });

    const { getTaskpaneShellState, setTaskpaneAuthenticated } = await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();
    setTaskpaneAuthenticated(AUTHENTICATED_SESSION);

    const appProps = getRenderedAppProps(reactMocks);
    await appProps.onAnalyze();
    await flushTaskpaneWork();

    expect(getTaskpaneShellState().analysisError).toEqual({
      message: "El análisis falló en 1 fragmento(s). Intenta de nuevo.",
      retryKind: "full-retry",
      visible: true,
    });
    expect(getTaskpaneShellState().progress.visible).toBe(false);
  });

  it("clears the persisted analysis error once a retry-query succeeds", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    taskpaneMocks.run.mockImplementationOnce(async (ctx) => {
      ctx.retryableRunReferences = [{ chunkIndex: 0, runId: "run-0" }];
      ctx.chunkErrors = ["Chunk 1: poll timeout"];
      ctx.emitter.emitProgress(1, 2, "Consultando resultado del fragmento 1 de 1...");
      ctx.emitter.emitAbort(
        "La consulta del análisis falló localmente. Reintentá la consulta con el mismo run."
      );
    });
    taskpaneMocks.retryPollChunkAnalysis.mockResolvedValueOnce({
      chunkIndex: 0,
      runId: "run-0",
      status: "success",
      origin: "backend",
      suggestions: [makeSuggestion()],
    });

    const { getTaskpaneShellState, setTaskpaneAuthenticated } = await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();
    setTaskpaneAuthenticated(AUTHENTICATED_SESSION);

    const appProps = getRenderedAppProps(reactMocks);
    await appProps.onAnalyze();
    await flushTaskpaneWork();
    await appProps.onRetryAnalysisQuery();
    await flushTaskpaneWork();

    expect(getTaskpaneShellState().analysisError.visible).toBe(false);
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

  it("hydrates the persisted analysis profile only when the stored id is known", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    const storage = installOfficeRuntime({
      getItem: vi.fn().mockResolvedValue("ensayo-academico"),
      setItem: vi.fn().mockResolvedValue(undefined),
      removeItem: vi.fn().mockResolvedValue(undefined),
    });
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    const { getTaskpaneShellState } = await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();

    expect(storage.getItem).toHaveBeenCalledWith("stylistic.preferences.analysis-profile.v1");
    expect(getTaskpaneShellState().selectedGenero).toBe("ensayo-academico");
  });

  it("ignores invalid persisted analysis profile ids and keeps the default", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    installOfficeRuntime({
      getItem: vi.fn().mockResolvedValue("perfil-inexistente"),
      setItem: vi.fn().mockResolvedValue(undefined),
      removeItem: vi.fn().mockResolvedValue(undefined),
    });
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    const { getTaskpaneShellState } = await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();

    expect(getTaskpaneShellState().selectedGenero).toBe("narrativa-literaria");
  });

  it("does NOT auto-persist analysis profile changes after bootstrap", async () => {
    // Settings is now draft + Save: changing the shell store directly must
    // not write to OfficeRuntime. Persistence only happens via the save flow.
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    const storage = installOfficeRuntime({
      getItem: vi.fn().mockResolvedValue(null),
      setItem: vi.fn().mockResolvedValue(undefined),
      removeItem: vi.fn().mockResolvedValue(undefined),
    });
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    const { setTaskpaneSelectedGenero: setSelectedGeneroFromStore } = await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();
    storage.setItem.mockClear();

    setSelectedGeneroFromStore("general");
    await flushTaskpaneWork();

    expect(storage.setItem).not.toHaveBeenCalled();
  });

  it("exposes loadPreferences as an App prop wired to the backend endpoint", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    installOfficeRuntime();
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;
    const fetchMock = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          correctionInstructions: "Vigilá X.",
          correctionInstructionsMaxLength: 4000,
        }),
        { status: 200, headers: { "content-type": "application/json" } }
      )
    );
    (globalThis as any).fetch = fetchMock;

    await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();

    const appProps = getRenderedAppProps(reactMocks);
    const result = await appProps.loadPreferences();

    expect(result).toEqual({
      correctionInstructions: "Vigilá X.",
      correctionInstructionsMaxLength: 4000,
    });
    expect(fetchMock).toHaveBeenCalledWith(
      "http://localhost:4111/user/preferences",
      expect.objectContaining({ method: "GET" })
    );
  });

  it("exposes savePreferences as an App prop and persists profile + instructions atomically", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    const storage = installOfficeRuntime({
      getItem: vi.fn().mockResolvedValue(null),
      setItem: vi.fn().mockResolvedValue(undefined),
      removeItem: vi.fn().mockResolvedValue(undefined),
    });
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;

    const putResponseBody = {
      correctionInstructions: "Vigilá Y.",
      correctionInstructionsMaxLength: 4000,
    };
    const fetchMock = vi.fn().mockImplementation(
      async () =>
        new Response(JSON.stringify(putResponseBody), {
          status: 200,
          headers: { "content-type": "application/json" },
        })
    );
    (globalThis as any).fetch = fetchMock;

    const { getTaskpaneShellState } = await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();
    storage.setItem.mockClear();

    const appProps = getRenderedAppProps(reactMocks);
    const result = await appProps.savePreferences("Vigilá Y.", "general");

    expect(result).toEqual(putResponseBody);
    expect(fetchMock).toHaveBeenCalledWith(
      "http://localhost:4111/user/preferences",
      expect.objectContaining({
        method: "PUT",
        body: JSON.stringify({ correctionInstructions: "Vigilá Y." }),
      })
    );
    expect(storage.setItem).toHaveBeenCalledWith(
      "stylistic.preferences.analysis-profile.v1",
      "general"
    );
    expect(getTaskpaneShellState().selectedGenero).toBe("general");
  });

  it("does NOT persist the profile when the backend save fails", async () => {
    const doc = createTaskpaneDocument();
    const officeHarness = createOffice();
    const storage = installOfficeRuntime({
      getItem: vi.fn().mockResolvedValue(null),
      setItem: vi.fn().mockResolvedValue(undefined),
      removeItem: vi.fn().mockResolvedValue(undefined),
    });
    (globalThis as any).document = doc;
    (globalThis as any).Office = officeHarness.office;
    const fetchMock = vi.fn().mockResolvedValue(
      new Response(JSON.stringify({ error: "unauthenticated" }), {
        status: 401,
        headers: { "content-type": "application/json" },
      })
    );
    (globalThis as any).fetch = fetchMock;

    const { getTaskpaneShellState } = await importTaskpaneRuntime();
    officeHarness.triggerReady({ host: "Word" });
    await flushTaskpaneWork();
    storage.setItem.mockClear();
    const initialProfile = getTaskpaneShellState().selectedGenero;

    const appProps = getRenderedAppProps(reactMocks);

    await expect(appProps.savePreferences("Vigilá Y.", "general")).rejects.toMatchObject({
      reason: "unauthenticated",
    });
    expect(storage.setItem).not.toHaveBeenCalled();
    expect(getTaskpaneShellState().selectedGenero).toBe(initialProfile);
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
