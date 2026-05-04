import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { SelectionSnapshot } from "../../domain/selection/SelectionSnapshot.types";
import { WordSelectionMonitorAdapter } from "./WordSelectionMonitorAdapter";

type SelectionChangedHandler = (event: { type: string }) => void;

interface OfficeHarness {
  trigger: () => void;
  addHandlerAsync: ReturnType<typeof vi.fn>;
  removeHandlerAsync: ReturnType<typeof vi.fn>;
}

function setupOfficeHarness(): OfficeHarness {
  let registeredHandler: SelectionChangedHandler | undefined;
  const addHandlerAsync = vi.fn(
    (
      _eventType: string,
      handler: SelectionChangedHandler,
      callback?: (asyncResult: { status: string }) => void
    ) => {
      registeredHandler = handler;
      callback?.({ status: "succeeded" });
    }
  );
  const removeHandlerAsync = vi.fn(
    (_eventType: string, _options?: unknown, callback?: (r: { status: string }) => void) => {
      registeredHandler = undefined;
      callback?.({ status: "succeeded" });
    }
  );

  (globalThis as any).Office = {
    EventType: { DocumentSelectionChanged: "documentSelectionChanged" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    context: {
      document: {
        addHandlerAsync,
        removeHandlerAsync,
      },
    },
  };

  return {
    trigger: () => registeredHandler?.({ type: "documentSelectionChanged" }),
    addHandlerAsync,
    removeHandlerAsync,
  };
}

function setupWordRunMock(selectionTexts: string[]) {
  const queue = [...selectionTexts];
  const wordRun = vi.fn(async (callback: (ctx: any) => Promise<unknown>) => {
    const text = queue.shift() ?? "";
    const selection = {
      text,
      load: vi.fn(),
    };
    const ctx = {
      document: { getSelection: () => selection },
      sync: vi.fn(async () => {}),
    };
    return callback(ctx);
  });
  (globalThis as any).Word = { run: wordRun };
  return wordRun;
}

describe("WordSelectionMonitorAdapter", () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
    delete (globalThis as any).Office;
    delete (globalThis as any).Word;
  });

  it("emits an initial empty snapshot when no selection is active", async () => {
    setupOfficeHarness();
    setupWordRunMock([""]);
    const monitor = new WordSelectionMonitorAdapter();
    const listener = vi.fn<(snapshot: SelectionSnapshot) => void>();

    monitor.subscribe(listener);
    await vi.runAllTimersAsync();

    expect(listener).toHaveBeenCalledWith({
      hasSelection: false,
      charCount: 0,
      preview: "",
    });
  });

  it("emits a snapshot with charCount and truncated preview when selection has text", async () => {
    const office = setupOfficeHarness();
    const longSelection = "A".repeat(150);
    setupWordRunMock(["", longSelection]);
    const monitor = new WordSelectionMonitorAdapter({ previewMaxChars: 80 });
    const listener = vi.fn<(snapshot: SelectionSnapshot) => void>();

    monitor.subscribe(listener);
    await vi.runAllTimersAsync();
    office.trigger();
    await vi.runAllTimersAsync();

    const lastCall = listener.mock.calls[listener.mock.calls.length - 1][0];
    expect(lastCall.hasSelection).toBe(true);
    expect(lastCall.charCount).toBe(150);
    expect(lastCall.preview).toBe(`${"A".repeat(80)}…`);
  });

  it("treats a whitespace-only selection as no selection", async () => {
    const office = setupOfficeHarness();
    setupWordRunMock(["", "   \n\t  "]);
    const monitor = new WordSelectionMonitorAdapter();
    const listener = vi.fn<(snapshot: SelectionSnapshot) => void>();

    monitor.subscribe(listener);
    await vi.runAllTimersAsync();
    office.trigger();
    await vi.runAllTimersAsync();

    const lastCall = listener.mock.calls[listener.mock.calls.length - 1][0];
    expect(lastCall.hasSelection).toBe(false);
    expect(lastCall.charCount).toBe(0);
    expect(lastCall.preview).toBe("");
  });

  it("registers the Office selection-changed handler exactly once for multiple subscribers", async () => {
    const office = setupOfficeHarness();
    setupWordRunMock(["", "", "hola"]);
    const monitor = new WordSelectionMonitorAdapter();
    const listenerA = vi.fn<(snapshot: SelectionSnapshot) => void>();
    const listenerB = vi.fn<(snapshot: SelectionSnapshot) => void>();

    monitor.subscribe(listenerA);
    monitor.subscribe(listenerB);
    await vi.runAllTimersAsync();
    office.trigger();
    await vi.runAllTimersAsync();

    expect(office.addHandlerAsync).toHaveBeenCalledTimes(1);
    expect(listenerA).toHaveBeenCalled();
    expect(listenerB).toHaveBeenCalled();
  });

  it("stops notifying a listener after its unsubscribe is called", async () => {
    const office = setupOfficeHarness();
    setupWordRunMock(["", "", "actualizado"]);
    const monitor = new WordSelectionMonitorAdapter();
    const listener = vi.fn<(snapshot: SelectionSnapshot) => void>();

    const unsubscribe = monitor.subscribe(listener);
    await vi.runAllTimersAsync();
    unsubscribe();
    listener.mockClear();
    office.trigger();
    await vi.runAllTimersAsync();

    expect(listener).not.toHaveBeenCalled();
  });
});
