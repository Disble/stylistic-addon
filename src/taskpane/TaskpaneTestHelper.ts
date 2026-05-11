import { vi } from "vitest";
import * as React from "react";

import type { Suggestion, SuggestionNavigationResult } from "../domain/suggestion/Suggestion.types";
import "./TaskpaneFluentMocks";
export {
  createOffice,
  createTaskpaneDocument,
  FakeClassList,
  FakeDocument,
  FakeElement,
} from "./TaskpaneDomTestHelper";

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
  getCleanupPreview: vi.fn<() => Promise<{ deletable: number; kept: number }>>(),
  cleanupResolvedComments: vi.fn<() => Promise<{ deleted: number; kept: number }>>(),
  applySuggestions: vi.fn<(suggestions: any[]) => Promise<any>>(),
  acceptSuggestion: vi.fn(),
  rejectSuggestion: vi.fn(),
  disableTrackChanges: vi.fn<() => Promise<void>>(),
  navigateToText: vi.fn<(target: Suggestion | string) => Promise<SuggestionNavigationResult>>(),
  mastraAdapterConstructor: vi.fn(),
  retryDecoratorConstructor: vi.fn(),
  feedbackSendFeedback: vi.fn<(payload: any) => Promise<void>>(),
  cancelChunkAnalysis: vi.fn<(chunkIndex: number, runId: string) => Promise<any>>(),
  retryPollChunkAnalysis:
    vi.fn<(reference: { chunkIndex: number; runId: string }) => Promise<any>>(),
}));

const hoistedReactMocks = vi.hoisted(() => ({
  createRoot: vi.fn(),
  render: vi.fn(),
}));

/**
 * Returns the shared mock registry used by taskpane presentation tests.
 */
export function getTaskpaneMocks() {
  return hoistedTaskpaneMocks;
}

/** Returns shared React-root mocks used by the taskpane entrypoint tests. */
export function getTaskpaneReactMocks() {
  return hoistedReactMocks;
}

vi.mock("react-dom/client", () => ({
  createRoot: vi.fn((...args: unknown[]) => {
    hoistedReactMocks.createRoot(...args);
    return {
      render: vi.fn((...renderArgs: unknown[]) => {
        hoistedReactMocks.render(...renderArgs);
        simulateMountedReactTree(renderArgs[0]);
      }),
    };
  }),
}));

/**
 * Simulates React mount effects in tests by walking the rendered element tree
 * and invoking any `onMount` callback props exposed by shell components.
 */
function simulateMountedReactTree(node: unknown): void {
  if (!isReactElementLike(node)) {
    return;
  }

  const onMount = node.props.onMount;
  if (typeof onMount === "function") {
    onMount();
  }

  const children = React.Children.toArray(node.props.children);
  for (const child of children) {
    simulateMountedReactTree(child);
  }
}

/** Returns true when the value looks like a React element object. */
function isReactElementLike(
  value: unknown
): value is React.ReactElement<{ children?: React.ReactNode; onMount?: unknown }> {
  return typeof value === "object" && value !== null && "props" in value;
}

vi.mock("../adapters/word/WordAdapter", () => ({
  WordAdapter: class {
    constructor() {
      hoistedTaskpaneMocks.wordAdapterConstructor();
    }

    cleanupResolvedComments() {
      return hoistedTaskpaneMocks.cleanupResolvedComments();
    }

    applySuggestions(suggestions: any[]) {
      return hoistedTaskpaneMocks.applySuggestions(suggestions);
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

    subscribeSelectionChanges() {
      return () => {};
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

    cancelChunkAnalysis(chunkIndex: number, runId: string) {
      return hoistedTaskpaneMocks.cancelChunkAnalysis(chunkIndex, runId);
    }

    retryPollChunkAnalysis(reference: { chunkIndex: number; runId: string }) {
      return hoistedTaskpaneMocks.retryPollChunkAnalysis(reference);
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
export function makeSuggestion(overrides: Partial<Suggestion> = {}): Suggestion {
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
 * Re-imports the taskpane entrypoint after resetting the module graph.
 * Tests that assert store state must import the stores again after this call
 * so they observe the same module instances used by the entrypoint.
 */
export async function importTaskpane() {
  vi.resetModules();
  return import("./index");
}

/**
 * Resets taskpane spies and fake globals to a known baseline.
 */
export function resetTaskpaneHarness() {
  const taskpaneMocks = getTaskpaneMocks();
  const reactMocks = getTaskpaneReactMocks();
  vi.resetAllMocks();
  vi.useFakeTimers();
  taskpaneMocks.orchestratorHandlers = [];
  reactMocks.createRoot.mockReset();
  reactMocks.render.mockReset();
  taskpaneMocks.run.mockResolvedValue(undefined);
  taskpaneMocks.getCleanupPreview.mockResolvedValue({
    deletable: 0,
    kept: 0,
  });
  taskpaneMocks.cleanupResolvedComments.mockResolvedValue({
    deleted: 0,
    kept: 0,
  });
  taskpaneMocks.applySuggestions.mockResolvedValue({
    successCount: 1,
    failedSuggestions: [],
    pendingAfter: {
      pendingStylisticArtifacts: 1,
      hasPendingStylisticArtifacts: true,
      trackChangesActive: true,
    },
    documentState: "pending-review",
    trackChangesActivatedForBatch: false,
  });
  taskpaneMocks.feedbackSendFeedback.mockResolvedValue(undefined);
  taskpaneMocks.cancelChunkAnalysis.mockResolvedValue({ canceled: true });
  taskpaneMocks.retryPollChunkAnalysis.mockResolvedValue({
    chunkIndex: 0,
    runId: "run-0",
    status: "retryable-failure",
    origin: "frontend-retryable",
    suggestions: [],
    error: "Chunk 1: poll timeout",
  });
  taskpaneMocks.cancelChunkAnalysis.mockClear();
  taskpaneMocks.retryPollChunkAnalysis.mockClear();
  taskpaneMocks.disableTrackChanges.mockResolvedValue(undefined);
  taskpaneMocks.navigateToText.mockResolvedValue({ status: "navigated" });
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
  delete (globalThis as any).OfficeRuntime;
}

/**
 * Restores globals and timers for taskpane tests.
 */
export function teardownTaskpaneHarness() {
  vi.useRealTimers();
  delete (globalThis as any).document;
  delete (globalThis as any).Office;
  delete (globalThis as any).OfficeRuntime;
}
