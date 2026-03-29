import { DEFAULT_MAX_CHUNK_SIZE, MAX_RETRIES, RETRY_BASE_DELAY_MS } from "../infrastructure/config";
import type { InsertionResult, Suggestion } from "../domain/types";

const taskpaneMocks = vi.hoisted(() => ({
  orchestratorHandlers: [] as unknown[],
  run: vi.fn<(ctx: any) => Promise<void>>(),
  wordAdapterConstructor: vi.fn(),
  cleanupResolvedComments: vi.fn<() => Promise<{ deleted: number; kept: number }>>(),
  acceptSuggestion: vi.fn(),
  rejectSuggestion: vi.fn(),
  mastraAdapterConstructor: vi.fn(),
  retryDecoratorConstructor: vi.fn(),
  feedbackSendFeedback: vi.fn<(payload: any) => Promise<void>>(),
}));

vi.mock("../adapters/word/WordAdapter", () => ({
  WordAdapter: class {
    constructor() {
      taskpaneMocks.wordAdapterConstructor();
    }

    cleanupResolvedComments() {
      return taskpaneMocks.cleanupResolvedComments();
    }

    acceptSuggestion() {
      return taskpaneMocks.acceptSuggestion();
    }

    rejectSuggestion() {
      return taskpaneMocks.rejectSuggestion();
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

vi.mock("../adapters/mastra/FeedbackAdapter", () => ({
  FeedbackAdapter: class {
    sendFeedback(payload: any) {
      return taskpaneMocks.feedbackSendFeedback(payload);
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
    type: "track-change",
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

class FakeClassList {
  private classes = new Set<string>();

  add(...names: string[]) {
    for (const n of names) this.classes.add(n);
  }

  remove(...names: string[]) {
    for (const n of names) this.classes.delete(n);
  }

  toggle(name: string): boolean {
    if (this.classes.has(name)) {
      this.classes.delete(name);
      return false;
    } else {
      this.classes.add(name);
      return true;
    }
  }

  contains(name: string): boolean {
    return this.classes.has(name);
  }

  toString(): string {
    return Array.from(this.classes).join(" ");
  }
}

class FakeElement {
  style = { display: "", width: "" };
  /** className string — synced with classList */
  get className(): string {
    return this.classList.toString();
  }
  set className(value: string) {
    const cl = new FakeClassList();
    for (const c of value.split(/\s+/).filter(Boolean)) cl.add(c);
    this._classList = cl;
  }
  private _classList = new FakeClassList();
  get classList(): FakeClassList {
    return this._classList;
  }

  disabled = false;
  value = "";
  onclick: ((ev: MouseEvent) => any) | null = null;
  /** DOM children appended via appendChild */
  children: FakeElement[] = [];
  /** Parent element reference (set when appended) */
  parentElement: FakeElement | null = null;

  private text = "";
  private html = "";
  private listeners = new Map<string, Array<(ev: any) => void>>();

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
    // Clear children when innerHTML is wiped
    if (value === "") {
      for (const child of this.children) {
        child.parentElement = null;
      }
      this.children = [];
    }
  }

  appendChild(child: FakeElement): FakeElement {
    child.parentElement = this;
    this.children.push(child);
    return child;
  }

  addEventListener(event: string, handler: (ev: any) => void) {
    if (!this.listeners.has(event)) this.listeners.set(event, []);
    this.listeners.get(event)!.push(handler);
  }

  /** Synchronously fires all "click" listeners */
  click() {
    for (const handler of this.listeners.get("click") ?? []) {
      handler({} as MouseEvent);
    }
  }

  /**
   * Searches appended children by CSS attribute selector or class selector.
   * Supports: [attr="val"], [attr], .classname
   * Also searches recursively into children's children.
   */
  querySelector(selector: string): FakeElement | null {
    for (const child of this.children) {
      if (matchesSelector(child, selector)) return child;
      const nested = child.querySelector(selector);
      if (nested) return nested;
    }
    return null;
  }

  querySelectorAll(selector: string): FakeElement[] {
    const results: FakeElement[] = [];
    for (const child of this.children) {
      if (matchesSelector(child, selector)) results.push(child);
      const nested = child.querySelectorAll(selector);
      for (const n of nested) results.push(n);
    }
    return results;
  }

  /** Removes this element from its parent */
  remove() {
    if (this.parentElement) {
      const idx = this.parentElement.children.indexOf(this);
      if (idx !== -1) this.parentElement.children.splice(idx, 1);
      this.parentElement = null;
    }
  }

  setAttribute(name: string, value: string) {
    (this as any)[`_attr_${name}`] = value;
  }

  getAttribute(name: string): string | null {
    return (this as any)[`_attr_${name}`] ?? null;
  }
}

function matchesSelector(el: FakeElement, selector: string): boolean {
  // [attr="value"]
  const attrEqMatch = selector.match(/^\[([^\]="]+)="([^"]+)"\]$/);
  if (attrEqMatch) {
    const [, attr, val] = attrEqMatch;
    return (el as any)[attr] === val || (el as any)[`_attr_${attr}`] === val;
  }
  // [attr]
  const attrMatch = selector.match(/^\[([^\]="]+)\]$/);
  if (attrMatch) {
    const [, attr] = attrMatch;
    return (el as any)[attr] !== undefined || (el as any)[`_attr_${attr}`] !== undefined;
  }
  // .className
  const classMatch = selector.match(/^\.(.+)$/);
  if (classMatch) {
    return el.classList.contains(classMatch[1]);
  }
  return false;
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

  createElement(_tagName?: string): FakeElement {
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
  doc.getElementById("profile-select")!.value = "narrativa-literaria";
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
    taskpaneMocks.feedbackSendFeedback.mockResolvedValue(undefined);

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
      genero: "narrativa-literaria",
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

// ---------------------------------------------------------------------------
// Helper: render results via the pipeline emitter and return the list children
// ---------------------------------------------------------------------------

async function renderViaEmitter(
  doc: FakeDocument,
  suggestions: Suggestion[],
  failedIds: string[] = []
): Promise<FakeElement[]> {
  const officeHarness = createOffice();
  (globalThis as any).document = doc;
  (globalThis as any).Office = officeHarness.office;

  const failedSuggestions = suggestions.filter((s) => failedIds.includes(s.id));
  const result: InsertionResult = {
    successCount: suggestions.length - failedSuggestions.length,
    failedSuggestions,
  };

  taskpaneMocks.run.mockImplementationOnce(async (ctx: any) => {
    ctx.emitter.emitComplete(suggestions, result, [], false);
  });

  await importTaskpane();
  officeHarness.triggerReady({ host: "Word" });
  await doc.getElementById("btn-analyze")!.onclick?.({} as MouseEvent);

  return doc.getElementById("results-list")!.children;
}

describe("Accept/Reject buttons", () => {
  let logSpy: ReturnType<typeof vi.spyOn>;
  let warnSpy: ReturnType<typeof vi.spyOn>;
  let errorSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    vi.clearAllMocks();
    vi.useFakeTimers();
    taskpaneMocks.orchestratorHandlers = [];
    taskpaneMocks.run.mockResolvedValue(undefined);
    taskpaneMocks.cleanupResolvedComments.mockResolvedValue({ deleted: 0, kept: 0 });
    taskpaneMocks.feedbackSendFeedback.mockResolvedValue(undefined);
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "accepted",
      trackedChangesAffected: 2,
      commentDeleted: true,
    });
    taskpaneMocks.rejectSuggestion.mockResolvedValue({
      status: "rejected",
      trackedChangesAffected: 2,
      commentDeleted: true,
    });

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

  it("4.0 — comment-only suggestion: no diff block, 'Entendido'/'Ignorar' button labels, type badge shown", async () => {
    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-co", type: "comment-only", suggestedText: undefined });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];

    // No diff block (result-change span) should be present
    expect(li.querySelector(".result-change")).toBeNull();

    // Accept button should read "Entendido", not "✓"
    const acceptBtn = li.querySelector('[data-action="accept"]');
    expect(acceptBtn).not.toBeNull();
    expect(acceptBtn!.textContent).toBe("Entendido");

    // Reject button should read "Ignorar", not "✗"
    const rejectBtn = li.querySelector('[data-action="reject"]');
    expect(rejectBtn).not.toBeNull();
    expect(rejectBtn!.textContent).toBe("Ignorar");

    // Comment-only type badge must be visible
    const typeBadge = li.querySelector(".result-type-badge--comment");
    expect(typeBadge).not.toBeNull();
    expect(typeBadge!.textContent).toBe("comentario");
  });

  it("4.0b — track-change suggestion: diff block present, '✓'/'✗' button labels", async () => {
    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-tc", type: "track-change", suggestedText: "texto sugerido" });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];

    // Diff block must be present for track-change
    expect(li.querySelector(".result-change")).not.toBeNull();

    const acceptBtn = li.querySelector('[data-action="accept"]');
    expect(acceptBtn).not.toBeNull();
    expect(acceptBtn!.textContent).toBe("✓");

    const rejectBtn = li.querySelector('[data-action="reject"]');
    expect(rejectBtn).not.toBeNull();
    expect(rejectBtn!.textContent).toBe("✗");
  });

  it("4.1 — renders accept and reject buttons for non-failed suggestions", async () => {
    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-1" });
    const s2 = makeSuggestion({ id: "s-2" });

    const liItems = await renderViaEmitter(doc, [s1, s2]);

    expect(liItems).toHaveLength(2);

    const li1 = liItems[0];
    const acceptBtn1 = li1.querySelector('[data-action="accept"]');
    const rejectBtn1 = li1.querySelector('[data-action="reject"]');
    expect(acceptBtn1).not.toBeNull();
    expect(rejectBtn1).not.toBeNull();
    expect(acceptBtn1!.getAttribute("data-suggestion-id")).toBe("s-1");
    expect(rejectBtn1!.getAttribute("data-suggestion-id")).toBe("s-1");

    const li2 = liItems[1];
    const acceptBtn2 = li2.querySelector('[data-action="accept"]');
    expect(acceptBtn2).not.toBeNull();
    expect(acceptBtn2!.getAttribute("data-suggestion-id")).toBe("s-2");
  });

  it("4.2 — failed suggestions do NOT have accept/reject buttons", async () => {
    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-fail" });

    const liItems = await renderViaEmitter(doc, [s1], ["s-fail"]);

    expect(liItems).toHaveLength(1);
    expect(liItems[0].querySelector('[data-action="accept"]')).toBeNull();
    expect(liItems[0].querySelector('[data-action="reject"]')).toBeNull();
  });

  it("4.3 — clicking Accept: adds result-accepted class and removes buttons", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "accepted",
      trackedChangesAffected: 2,
      commentDeleted: true,
    });

    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-1" });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];
    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;

    acceptBtn.click();
    // Flush microtasks (async handler)
    await Promise.resolve();
    await Promise.resolve();

    expect(li.classList.contains("result-accepted")).toBe(true);
    expect(li.querySelector('[data-action="accept"]')).toBeNull();
    expect(li.querySelector('[data-action="reject"]')).toBeNull();
  });

  it("4.4 — clicking Reject: adds result-rejected class and removes buttons", async () => {
    taskpaneMocks.rejectSuggestion.mockResolvedValue({
      status: "rejected",
      trackedChangesAffected: 2,
      commentDeleted: true,
    });

    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-1" });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];
    const rejectBtn = li.querySelector('[data-action="reject"]') as FakeElement;

    rejectBtn.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(li.classList.contains("result-rejected")).toBe(true);
    expect(li.querySelector('[data-action="reject"]')).toBeNull();
    expect(li.querySelector('[data-action="accept"]')).toBeNull();
  });

  it("4.5 — already-resolved: adds class, shows '(ya resuelto)' note, does NOT send feedback", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "already-resolved",
      trackedChangesAffected: 0,
      commentDeleted: false,
    });

    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-1" });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];
    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;

    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(li.classList.contains("result-already-resolved")).toBe(true);
    const noteSpan = li.querySelector(".result-already-resolved-note");
    expect(noteSpan).not.toBeNull();
    expect(noteSpan!.textContent).toBe("(ya resuelto)");
    // already-resolved must NOT send feedback — we don't know if user accepted or rejected
    expect(taskpaneMocks.feedbackSendFeedback).not.toHaveBeenCalled();
  });

  it("4.8 — cc-not-found: adds amber class, shows '(aplicación falló)' note, does NOT send feedback", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "cc-not-found",
      trackedChangesAffected: 0,
      commentDeleted: false,
    });

    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-1" });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];
    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;

    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(li.classList.contains("result-cc-not-found")).toBe(true);
    const noteSpan = li.querySelector(".result-cc-not-found-note");
    expect(noteSpan).not.toBeNull();
    expect(noteSpan!.textContent).toBe("(aplicación falló)");
    // Terminal UI — actions div removed, buttons not re-enabled
    expect(li.querySelector(".result-actions")).toBeNull();
    expect(taskpaneMocks.feedbackSendFeedback).not.toHaveBeenCalled();
  });

  it("4.9 — double-click guard: documentPort called exactly once", async () => {
    const { promise: firstCall, resolve: resolveFirst } = deferred<{ status: string; trackedChangesAffected: number; commentDeleted: boolean }>();
    taskpaneMocks.acceptSuggestion.mockReturnValue(firstCall);

    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-1" });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];
    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;

    // First click — SM transitions to "resolving"
    acceptBtn.click();
    await Promise.resolve();

    // Second click while first is still in-flight — SM guard rejects it
    acceptBtn.click();
    await Promise.resolve();

    // Resolve the first call
    resolveFirst({ status: "accepted", trackedChangesAffected: 1, commentDeleted: false });
    await Promise.resolve();
    await Promise.resolve();

    // Adapter must have been called exactly once
    expect(taskpaneMocks.acceptSuggestion).toHaveBeenCalledTimes(1);
    expect(li.classList.contains("result-accepted")).toBe(true);
  });

  it("4.10 — error retry: second click succeeds after error state", async () => {
    taskpaneMocks.acceptSuggestion
      .mockResolvedValueOnce({ status: "error", trackedChangesAffected: 0, commentDeleted: false, error: "timeout" })
      .mockResolvedValueOnce({ status: "accepted", trackedChangesAffected: 1, commentDeleted: true });

    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-1" });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];
    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;

    // First click — error
    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();
    expect(acceptBtn.disabled).toBe(false); // re-enabled after error

    // Second click — succeeds
    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();
    expect(li.classList.contains("result-accepted")).toBe(true);
    expect(taskpaneMocks.acceptSuggestion).toHaveBeenCalledTimes(2);
  });

  it("4.6 — error case: buttons re-enabled and showStatus called", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "error",
      trackedChangesAffected: 0,
      commentDeleted: false,
      error: "El documento está protegido",
    });

    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-1" });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];
    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;

    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();

    // Buttons should be re-enabled
    expect(acceptBtn.disabled).toBe(false);
    const rejectBtn = li.querySelector('[data-action="reject"]') as FakeElement;
    expect(rejectBtn.disabled).toBe(false);
    // Status bar should show error message
    expect(doc.getElementById("status-bar")!.textContent).toBe("El documento está protegido");
    expect(doc.getElementById("status-bar")!.className).toBe("stylistic-status error");
  });

  it("4.7 — not-found case: buttons re-enabled", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "not-found",
      trackedChangesAffected: 0,
      commentDeleted: false,
      error: "Texto no encontrado",
    });

    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-1" });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];
    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;

    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(acceptBtn.disabled).toBe(false);
    const rejectBtn = li.querySelector('[data-action="reject"]') as FakeElement;
    expect(rejectBtn.disabled).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Feedback Button + Accordion
// ---------------------------------------------------------------------------

describe("Feedback button + accordion", () => {
  let logSpy: ReturnType<typeof vi.spyOn>;
  let warnSpy: ReturnType<typeof vi.spyOn>;
  let errorSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    vi.clearAllMocks();
    vi.useFakeTimers();
    taskpaneMocks.orchestratorHandlers = [];
    taskpaneMocks.run.mockResolvedValue(undefined);
    taskpaneMocks.cleanupResolvedComments.mockResolvedValue({ deleted: 0, kept: 0 });
    taskpaneMocks.feedbackSendFeedback.mockResolvedValue(undefined);
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "accepted",
      trackedChangesAffected: 2,
      commentDeleted: true,
    });
    taskpaneMocks.rejectSuggestion.mockResolvedValue({
      status: "rejected",
      trackedChangesAffected: 2,
      commentDeleted: true,
    });

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

  it("F.1 — renderResults() injects 💬 button for each non-failed suggestion", async () => {
    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-1" });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];

    const feedbackBtn = li.querySelector('[data-action="feedback"]');
    expect(feedbackBtn).not.toBeNull();
    expect(feedbackBtn!.getAttribute("aria-label")).toBe("Dejar feedback");
  });

  it("F.2 — renderResults() injects .feedback-accordion with .feedback-textarea per non-failed suggestion", async () => {
    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-1" });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];

    const accordion = li.querySelector(".feedback-accordion");
    expect(accordion).not.toBeNull();

    const textarea = li.querySelector(".feedback-textarea");
    expect(textarea).not.toBeNull();
  });

  it("F.3 — failed suggestions do NOT have a 💬 button", async () => {
    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-fail" });

    const liItems = await renderViaEmitter(doc, [s1], ["s-fail"]);
    const li = liItems[0];

    expect(li.querySelector('[data-action="feedback"]')).toBeNull();
    expect(li.querySelector(".feedback-accordion")).toBeNull();
  });

  it("F.4 — clicking 💬 button toggles .feedback-accordion--open on accordion", async () => {
    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-1" });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];

    const feedbackBtn = li.querySelector('[data-action="feedback"]') as FakeElement;
    const accordion = li.querySelector(".feedback-accordion") as FakeElement;

    expect(accordion.classList.contains("feedback-accordion--open")).toBe(false);

    feedbackBtn.click();
    expect(accordion.classList.contains("feedback-accordion--open")).toBe(true);

    feedbackBtn.click();
    expect(accordion.classList.contains("feedback-accordion--open")).toBe(false);
  });

  it("F.5 — accept sends positive feedback with correct payload", async () => {
    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({
      id: "s-1",
      category: "Redundancia",
      originalText: "completamente necesario",
      suggestedText: "necesario",
      justification: "Ya implica completitud.",
      severity: "high",
    });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];

    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;
    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();
    // Flush the void sendFeedback microtask
    await Promise.resolve();

    expect(taskpaneMocks.feedbackSendFeedback).toHaveBeenCalledWith(
      expect.objectContaining({
        rating: "positive",
        category: "Redundancia",
        originalText: "completamente necesario",
        suggestedText: "necesario",
        justification: "Ya implica completitud.",
        severity: "high",
      })
    );
  });

  it("F.6 — reject sends negative feedback with correct payload", async () => {
    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({
      id: "s-1",
      category: "Muletilla",
      originalText: "básicamente",
      suggestedText: "",
      justification: "Frase de relleno.",
      severity: "medium",
    });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];

    const rejectBtn = li.querySelector('[data-action="reject"]') as FakeElement;
    rejectBtn.click();
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();

    expect(taskpaneMocks.feedbackSendFeedback).toHaveBeenCalledWith(
      expect.objectContaining({
        rating: "negative",
        category: "Muletilla",
        originalText: "básicamente",
        suggestedText: "",
        justification: "Frase de relleno.",
        severity: "medium",
      })
    );
  });

  it("F.7 — empty textarea => comment absent from payload", async () => {
    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-1" });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];

    // textarea value is empty (default)
    const textarea = li.querySelector(".feedback-textarea") as FakeElement;
    textarea.value = "";

    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;
    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();

    const payload = taskpaneMocks.feedbackSendFeedback.mock.calls[0][0];
    expect(payload).not.toHaveProperty("comment");
  });

  it("F.8 — textarea with text => comment present in payload", async () => {
    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-1" });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];

    const textarea = li.querySelector(".feedback-textarea") as FakeElement;
    textarea.value = "Muy buen cambio";

    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;
    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();

    expect(taskpaneMocks.feedbackSendFeedback).toHaveBeenCalledWith(
      expect.objectContaining({ comment: "Muy buen cambio" })
    );
  });

  it("F.9 — payload includes justification field", async () => {
    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-1", justification: "Es más claro" });

    const liItems = await renderViaEmitter(doc, [s1]);
    const li = liItems[0];

    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;
    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();

    expect(taskpaneMocks.feedbackSendFeedback).toHaveBeenCalledWith(
      expect.objectContaining({ justification: "Es más claro" })
    );
  });
});
