import { vi } from "vitest";

import type {
  ApplySuggestionsResult,
  Suggestion,
  SuggestionApplicationFailure,
} from "../domain/types";

const hoistedTaskpaneMocks = vi.hoisted(() => ({
  orchestratorHandlers: [] as unknown[],
  run: vi.fn<(ctx: any) => Promise<void>>(),
  wordAdapterConstructor: vi.fn(),
  getDocumentReviewState: vi.fn<
    () => Promise<{
      pendingStylisticArtifacts: number;
      hasPendingStylisticArtifacts: boolean;
      trackChangesActive: boolean;
    }>
  >(),
  getCleanupPreview:
    vi.fn<() => Promise<{ deletable: number; kept: number }>>(),
  cleanupResolvedComments:
    vi.fn<() => Promise<{ deleted: number; kept: number }>>(),
  acceptSuggestion: vi.fn(),
  rejectSuggestion: vi.fn(),
  disableTrackChanges: vi.fn<() => Promise<void>>(),
  navigateToText: vi.fn<(target: Suggestion | string) => Promise<void>>(),
  mastraAdapterConstructor: vi.fn(),
  retryDecoratorConstructor: vi.fn(),
  feedbackSendFeedback: vi.fn<(payload: any) => Promise<void>>(),
}));

/**
 * Returns the shared mock registry used by taskpane presentation tests.
 */
export function getTaskpaneMocks() {
  return hoistedTaskpaneMocks;
}

vi.mock("../adapters/word/WordAdapter", () => ({
  WordAdapter: class {
    constructor() {
      hoistedTaskpaneMocks.wordAdapterConstructor();
    }

    cleanupResolvedComments() {
      return hoistedTaskpaneMocks.cleanupResolvedComments();
    }

    getCleanupPreview() {
      return hoistedTaskpaneMocks.getCleanupPreview();
    }

    getDocumentReviewState() {
      return hoistedTaskpaneMocks.getDocumentReviewState();
    }

    acceptSuggestion(suggestion: any) {
      return hoistedTaskpaneMocks.acceptSuggestion(suggestion);
    }

    rejectSuggestion(suggestion: any) {
      return hoistedTaskpaneMocks.rejectSuggestion(suggestion);
    }

    disableTrackChanges() {
      return hoistedTaskpaneMocks.disableTrackChanges();
    }

    navigateToText(target: Suggestion | string) {
      return hoistedTaskpaneMocks.navigateToText(target);
    }
  },
}));

vi.mock("../adapters/mastra/MastraAdapter", () => ({
  MastraAdapter: class {
    constructor() {
      hoistedTaskpaneMocks.mastraAdapterConstructor();
    }

    /** Keeps mock shape explicit and avoids constructor-only class lint issues. */
    ping(): number {
      return 0;
    }
  },
}));

vi.mock("../adapters/RetryAnalysisDecorator", () => ({
  RetryAnalysisDecorator: class {
    constructor(...args: unknown[]) {
      hoistedTaskpaneMocks.retryDecoratorConstructor(...args);
    }

    /** Keeps mock shape explicit and avoids constructor-only class lint issues. */
    ping(): number {
      return 0;
    }
  },
}));

vi.mock("../domain/pipeline/PipelineOrchestrator", () => ({
  PipelineOrchestrator: class {
    constructor(handlers: unknown[]) {
      hoistedTaskpaneMocks.orchestratorHandlers = handlers;
    }

    run(ctx: any) {
      return hoistedTaskpaneMocks.run(ctx);
    }
  },
}));

vi.mock("../adapters/mastra/FeedbackAdapter", () => ({
  FeedbackAdapter: class {
    sendFeedback(payload: any) {
      return hoistedTaskpaneMocks.feedbackSendFeedback(payload);
    }
  },
}));

/**
 * Builds a canonical taskpane suggestion fixture.
 */
export function makeSuggestion(
  overrides: Partial<Suggestion> = {},
): Suggestion {
  const anchor = overrides.anchor ?? "texto original";
  return {
    id: "s-1",
    context: overrides.context ?? `Contexto con ${anchor}.`,
    anchor,
    suggestedText: "texto sugerido",
    justification: "Mas claro",
    category: "Claridad",
    severity: "medium",
    type: "track-change",
    ...overrides,
  };
}

/**
 * Creates a promise with externally controlled resolution.
 */
export function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

/**
 * Escapes user-provided text before fake DOM HTML snapshots.
 */
function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

export class FakeClassList {
  private readonly classes = new Set<string>();

  add(...names: string[]) {
    for (const name of names) {
      this.classes.add(name);
    }
  }

  remove(...names: string[]) {
    for (const name of names) {
      this.classes.delete(name);
    }
  }

  toggle(name: string): boolean {
    if (this.classes.has(name)) {
      this.classes.delete(name);
      return false;
    }

    this.classes.add(name);
    return true;
  }

  contains(name: string): boolean {
    return this.classes.has(name);
  }

  toString(): string {
    return Array.from(this.classes).join(" ");
  }
}

export class FakeElement {
  style = { display: "", width: "" };
  dataset: Record<string, string> = {};

  get className(): string {
    return this.classList.toString();
  }

  set className(value: string) {
    const classList = new FakeClassList();
    for (const name of value.split(/\s+/).filter(Boolean)) {
      classList.add(name);
    }
    this._classList = classList;
  }

  private _classList = new FakeClassList();

  get classList(): FakeClassList {
    return this._classList;
  }

  disabled = false;
  value = "";
  onclick: ((ev: MouseEvent) => any) | null = null;
  children: FakeElement[] = [];
  parentElement: FakeElement | null = null;

  private text = "";
  private html = "";
  private readonly listeners = new Map<string, Array<(ev: any) => void>>();

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
    if (value === "") {
      for (const child of this.children) {
        child.parentElement = null;
      }
      this.children = [];
    }
  }

  appendChild(child: FakeElement): FakeElement {
    if (child.parentElement) {
      const previousIndex = child.parentElement.children.indexOf(child);
      if (previousIndex !== -1) {
        child.parentElement.children.splice(previousIndex, 1);
      }
    }

    child.parentElement = this;
    this.children.push(child);
    return child;
  }

  /** Inserts a child before the given reference node, matching DOM semantics. */
  insertBefore(child: FakeElement, referenceNode: FakeElement | null): FakeElement {
    if (referenceNode === null) {
      return this.appendChild(child);
    }

    if (child.parentElement) {
      const previousIndex = child.parentElement.children.indexOf(child);
      if (previousIndex !== -1) {
        child.parentElement.children.splice(previousIndex, 1);
      }
    }

    const referenceIndex = this.children.indexOf(referenceNode);
    if (referenceIndex === -1) {
      throw new Error("Reference node is not a child of this element");
    }

    child.parentElement = this;
    this.children.splice(referenceIndex, 0, child);
    return child;
  }

  /** Inserts this node immediately before a sibling, matching Element.before(). */
  before(node: FakeElement): void {
    if (!this.parentElement) {
      return;
    }

    this.parentElement.insertBefore(node, this);
  }

  addEventListener(event: string, handler: (ev: any) => void) {
    if (!this.listeners.has(event)) {
      this.listeners.set(event, []);
    }
    this.listeners.get(event)?.push(handler);
  }

  click() {
    if (this.onclick) {
      this.onclick({} as MouseEvent);
    }
    for (const handler of this.listeners.get("click") ?? []) {
      handler({} as MouseEvent);
    }
  }

  querySelector(selector: string): FakeElement | null {
    for (const child of this.children) {
      if (matchesSelector(child, selector)) {
        return child;
      }
      const nested = child.querySelector(selector);
      if (nested) {
        return nested;
      }
    }
    return null;
  }

  querySelectorAll(selector: string): FakeElement[] {
    const results: FakeElement[] = [];
    for (const child of this.children) {
      if (matchesSelector(child, selector)) {
        results.push(child);
      }
      results.push(...child.querySelectorAll(selector));
    }
    return results;
  }

  remove() {
    if (this.parentElement) {
      const index = this.parentElement.children.indexOf(this);
      if (index !== -1) {
        this.parentElement.children.splice(index, 1);
      }
      this.parentElement = null;
    }
  }

  setAttribute(name: string, value: string) {
    (this as any)[`_attr_${name}`] = value;
    if (name.startsWith("data-")) {
      const datasetKey = name
        .slice(5)
        .replace(/-([a-z])/g, (_, letter: string) => letter.toUpperCase());
      this.dataset[datasetKey] = value;
    }
  }

  getAttribute(name: string): string | null {
    return (this as any)[`_attr_${name}`] ?? null;
  }
}

function matchesSelector(el: FakeElement, selector: string): boolean {
  const attrEqMatch = /^\[([^\]="]+)="([^"]+)"\]$/.exec(selector);
  if (attrEqMatch) {
    const [, attr, value] = attrEqMatch;
    if (attr.startsWith("data-")) {
      const datasetKey = attr
        .slice(5)
        .replace(/-([a-z])/g, (_, letter: string) => letter.toUpperCase());
      return el.dataset[datasetKey] === value;
    }

    return (
      (el as any)[attr] === value || (el as any)[`_attr_${attr}`] === value
    );
  }

  const attrMatch = /^\[([^\]="]+)\]$/.exec(selector);
  if (attrMatch) {
    const [, attr] = attrMatch;
    if (attr.startsWith("data-")) {
      const datasetKey = attr
        .slice(5)
        .replace(/-([a-z])/g, (_, letter: string) => letter.toUpperCase());
      return el.dataset[datasetKey] !== undefined;
    }

    return (
      (el as any)[attr] !== undefined ||
      (el as any)[`_attr_${attr}`] !== undefined
    );
  }

  const classMatch = /^\.(.+)$/.exec(selector);
  if (classMatch) {
    return el.classList.contains(classMatch[1]);
  }

  return false;
}

export class FakeDocument {
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

/** Returns a required fake element by id or throws in tests. */
function getRequiredElement(doc: FakeDocument, id: string): FakeElement {
  const el = doc.getElementById(id);
  if (!el) {
    throw new Error(`Missing fake DOM element: ${id}`);
  }
  return el;
}

/**
 * Creates the fake taskpane DOM with the default initial visibility state.
 */
export function createTaskpaneDocument(): FakeDocument {
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
    "disable-track-changes-section",
    "btn-disable-track-changes",
    "btn-disable-track-changes-label",
  ]);

  getRequiredElement(doc, "sideload-msg").style.display = "block";
  getRequiredElement(doc, "app-body").style.display = "none";
  getRequiredElement(doc, "cleanup-section").style.display = "none";
  getRequiredElement(doc, "disable-track-changes-section").style.display =
    "none";
  getRequiredElement(doc, "results-panel").style.display = "block";
  getRequiredElement(doc, "progress-container").style.display = "none";
  getRequiredElement(doc, "profile-select").value = "narrativa-literaria";
  getRequiredElement(doc, "btn-analyze-label").textContent =
    "Analizar y sugerir";
  getRequiredElement(doc, "btn-cleanup-label").textContent =
    "Limpiar comentarios resueltos";
  getRequiredElement(doc, "btn-disable-track-changes-label").textContent =
    "Desactivar control de cambios";

  return doc;
}

type ReadyCallback = (info: { host: string }) => void;

/**
 * Creates a controllable `Office.onReady` harness for taskpane tests.
 */
export function createOffice() {
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

/**
 * Re-imports the taskpane entrypoint after resetting Vitest module state.
 */
export async function importTaskpane() {
  vi.resetModules();
  return import("./taskpane");
}

/**
 * Resets taskpane spies and fake globals to a known baseline.
 */
export function resetTaskpaneHarness() {
  const taskpaneMocks = getTaskpaneMocks();
  vi.resetAllMocks();
  vi.useFakeTimers();
  taskpaneMocks.orchestratorHandlers = [];
  taskpaneMocks.run.mockResolvedValue(undefined);
  taskpaneMocks.getCleanupPreview.mockResolvedValue({
    deletable: 0,
    kept: 0,
  });
  taskpaneMocks.cleanupResolvedComments.mockResolvedValue({
    deleted: 0,
    kept: 0,
  });
  taskpaneMocks.feedbackSendFeedback.mockResolvedValue(undefined);
  taskpaneMocks.disableTrackChanges.mockResolvedValue(undefined);
  taskpaneMocks.acceptSuggestion.mockResolvedValue({
    status: "accepted",
    trackedChangesAffected: 2,
    commentDeleted: true,
    pendingAfter: {
      pendingStylisticArtifacts: 1,
      hasPendingStylisticArtifacts: true,
      trackChangesActive: true,
    },
    documentState: "pending-review",
    feedbackStatus: "sent",
    taskpaneState: {
      documentState: "pending-review",
      showDisableTrackChangesCta: false,
      showCleanupSection: false,
    },
  });
  taskpaneMocks.rejectSuggestion.mockResolvedValue({
    status: "rejected",
    trackedChangesAffected: 2,
    commentDeleted: true,
    pendingAfter: {
      pendingStylisticArtifacts: 1,
      hasPendingStylisticArtifacts: true,
      trackChangesActive: true,
    },
    documentState: "pending-review",
    feedbackStatus: "sent",
    taskpaneState: {
      documentState: "pending-review",
      showDisableTrackChangesCta: false,
      showCleanupSection: false,
    },
  });

  delete (globalThis as any).document;
  delete (globalThis as any).Office;
}

/**
 * Restores globals and timers for taskpane tests.
 */
export function teardownTaskpaneHarness() {
  vi.useRealTimers();
  delete (globalThis as any).document;
  delete (globalThis as any).Office;
}

/**
 * Renders results by emitting a pipeline completion event through the entrypoint.
 */
export async function renderViaEmitter(
  doc: FakeDocument,
  suggestions: Suggestion[],
  failedIds: string[] = [],
): Promise<FakeElement[]> {
  const taskpaneMocks = getTaskpaneMocks();
  const officeHarness = createOffice();
  (globalThis as any).document = doc;
  (globalThis as any).Office = officeHarness.office;

  const failedSuggestions: SuggestionApplicationFailure[] = suggestions
    .filter((suggestion) => failedIds.includes(suggestion.id))
    .map((suggestion) => ({
      suggestion,
      reason: "not-found",
      message: "Anchor no encontrado en el contexto",
    }));
  const result: ApplySuggestionsResult = {
    successCount: suggestions.length - failedSuggestions.length,
    failedSuggestions,
    pendingAfter: {
      pendingStylisticArtifacts: suggestions.length - failedSuggestions.length,
      hasPendingStylisticArtifacts:
        suggestions.length - failedSuggestions.length > 0,
      trackChangesActive: suggestions.length - failedSuggestions.length > 0,
    },
    documentState:
      suggestions.length - failedSuggestions.length > 0
        ? "pending-review"
        : "idle",
    trackChangesActivatedForBatch:
      suggestions.length - failedSuggestions.length > 0,
  };

  taskpaneMocks.run.mockImplementationOnce(async (ctx: any) => {
    ctx.emitter.emitComplete(suggestions, result, [], false);
  });

  await importTaskpane();
  officeHarness.triggerReady({ host: "Word" });
  await doc.getElementById("btn-analyze")?.onclick?.({} as MouseEvent);

  return doc.getElementById("results-list")?.children ?? [];
}
