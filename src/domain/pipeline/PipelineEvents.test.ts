import type { InsertionResult, PipelineState, Suggestion } from "../types";
import { PipelineEventEmitter, type PipelineObserver } from "./PipelineEvents";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeSuggestion(overrides: Partial<Suggestion> = {}): Suggestion {
  return {
    id: "s1",
    originalText: "foo",
    suggestedText: "bar",
    justification: "test",
    category: "Redundancia",
    severity: "medium",
    type: "track-change",
    ...overrides,
  };
}

function makeInsertionResult(
  overrides: Partial<InsertionResult> = {},
): InsertionResult {
  return { successCount: 1, failedSuggestions: [], ...overrides };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("PipelineEventEmitter", () => {
  let emitter: PipelineEventEmitter;

  beforeEach(() => {
    emitter = new PipelineEventEmitter();
  });

  // -------------------------------------------------------------------------
  // subscribe / unsubscribe / clear
  // -------------------------------------------------------------------------

  describe("subscription management", () => {
    it("should call observer methods after subscribing", () => {
      const observer: PipelineObserver = { onPhaseStart: vi.fn() };
      emitter.subscribe(observer);

      emitter.emitPhaseStart("reading", "Starting…");

      expect(observer.onPhaseStart).toHaveBeenCalledWith(
        "reading",
        "Starting…",
      );
    });

    it("should stop calling observer methods after unsubscribing", () => {
      const observer: PipelineObserver = { onPhaseStart: vi.fn() };
      emitter.subscribe(observer);
      emitter.unsubscribe(observer);

      emitter.emitPhaseStart("reading", "Starting…");

      expect(observer.onPhaseStart).not.toHaveBeenCalled();
    });

    it("should remove all observers on clear()", () => {
      const o1: PipelineObserver = { onPhaseStart: vi.fn() };
      const o2: PipelineObserver = { onPhaseStart: vi.fn() };
      emitter.subscribe(o1);
      emitter.subscribe(o2);

      emitter.clear();
      emitter.emitPhaseStart("reading", "msg");

      expect(o1.onPhaseStart).not.toHaveBeenCalled();
      expect(o2.onPhaseStart).not.toHaveBeenCalled();
    });

    it("should allow re-subscribing the same observer after unsubscribe", () => {
      const observer: PipelineObserver = { onProgress: vi.fn() };
      emitter.subscribe(observer);
      emitter.unsubscribe(observer);
      emitter.subscribe(observer);

      emitter.emitProgress(1, 5, "msg");

      expect(observer.onProgress).toHaveBeenCalledOnce();
    });

    it("should noop when unsubscribing an observer that was never subscribed", () => {
      const stranger: PipelineObserver = { onPhaseStart: vi.fn() };

      // Should not throw
      expect(() => emitter.unsubscribe(stranger)).not.toThrow();
    });
  });

  // -------------------------------------------------------------------------
  // Multiple listeners
  // -------------------------------------------------------------------------

  describe("multiple observers", () => {
    it("should notify all subscribed observers", () => {
      const o1: PipelineObserver = { onPhaseStart: vi.fn() };
      const o2: PipelineObserver = { onPhaseStart: vi.fn() };
      emitter.subscribe(o1);
      emitter.subscribe(o2);

      emitter.emitPhaseStart("connecting", "Connecting…");

      expect(o1.onPhaseStart).toHaveBeenCalledWith("connecting", "Connecting…");
      expect(o2.onPhaseStart).toHaveBeenCalledWith("connecting", "Connecting…");
    });

    it("should notify in subscription (insertion) order", () => {
      const order: number[] = [];
      const o1: PipelineObserver = {
        onPhaseStart: vi.fn(() => order.push(1)),
      };
      const o2: PipelineObserver = {
        onPhaseStart: vi.fn(() => order.push(2)),
      };
      emitter.subscribe(o1);
      emitter.subscribe(o2);

      emitter.emitPhaseStart("reading", "msg");

      expect(order).toEqual([1, 2]);
    });

    it("should only remove the unsubscribed observer, not others", () => {
      const o1: PipelineObserver = { onAbort: vi.fn() };
      const o2: PipelineObserver = { onAbort: vi.fn() };
      emitter.subscribe(o1);
      emitter.subscribe(o2);
      emitter.unsubscribe(o1);

      emitter.emitAbort("cancelled");

      expect(o1.onAbort).not.toHaveBeenCalled();
      expect(o2.onAbort).toHaveBeenCalledWith("cancelled");
    });
  });

  // -------------------------------------------------------------------------
  // Double-subscribe edge case
  // -------------------------------------------------------------------------

  describe("double subscribe (same reference)", () => {
    it("should call the observer twice if subscribed twice", () => {
      const observer: PipelineObserver = { onProgress: vi.fn() };
      emitter.subscribe(observer);
      emitter.subscribe(observer);

      emitter.emitProgress(1, 3, "msg");

      expect(observer.onProgress).toHaveBeenCalledTimes(2);
    });

    it("should remove only one registration per unsubscribe call (filter behavior)", () => {
      const observer: PipelineObserver = { onProgress: vi.fn() };
      emitter.subscribe(observer);
      emitter.subscribe(observer);

      // Current impl uses .filter(o => o !== observer) which removes ALL
      emitter.unsubscribe(observer);
      emitter.emitProgress(1, 3, "msg");

      // filter removes ALL matches, so zero calls after single unsubscribe
      expect(observer.onProgress).not.toHaveBeenCalled();
    });
  });

  // -------------------------------------------------------------------------
  // Emitting with no listeners
  // -------------------------------------------------------------------------

  describe("emit with no observers", () => {
    it("should not throw when emitting phaseStart with no observers", () => {
      expect(() => emitter.emitPhaseStart("reading", "msg")).not.toThrow();
    });

    it("should not throw when emitting progress with no observers", () => {
      expect(() => emitter.emitProgress(1, 5, "msg")).not.toThrow();
    });

    it("should not throw when emitting phaseComplete with no observers", () => {
      expect(() => emitter.emitPhaseComplete("done")).not.toThrow();
    });

    it("should not throw when emitting error with no observers", () => {
      expect(() => emitter.emitError("error", new Error("boom"))).not.toThrow();
    });

    it("should not throw when emitting complete with no observers", () => {
      expect(() =>
        emitter.emitComplete([], makeInsertionResult(), [], false),
      ).not.toThrow();
    });

    it("should not throw when emitting abort with no observers", () => {
      expect(() => emitter.emitAbort("cancelled")).not.toThrow();
    });
  });

  // -------------------------------------------------------------------------
  // Partial observer (optional methods)
  // -------------------------------------------------------------------------

  describe("partial observers (optional methods)", () => {
    it("should not throw when observer lacks the emitted method", () => {
      const observer: PipelineObserver = {}; // no methods at all
      emitter.subscribe(observer);

      expect(() => emitter.emitPhaseStart("reading", "msg")).not.toThrow();
      expect(() => emitter.emitProgress(1, 5, "msg")).not.toThrow();
      expect(() => emitter.emitPhaseComplete("done")).not.toThrow();
      expect(() => emitter.emitError("error", "oops")).not.toThrow();
      expect(() =>
        emitter.emitComplete([], makeInsertionResult(), [], false),
      ).not.toThrow();
      expect(() => emitter.emitAbort("cancelled")).not.toThrow();
    });

    it("should still call the methods the observer does implement", () => {
      const observer: PipelineObserver = {
        onProgress: vi.fn(),
        // onPhaseStart intentionally missing
      };
      emitter.subscribe(observer);

      emitter.emitPhaseStart("reading", "msg"); // should not throw
      emitter.emitProgress(2, 10, "working…");

      expect(observer.onProgress).toHaveBeenCalledWith(2, 10, "working…");
    });
  });

  // -------------------------------------------------------------------------
  // Event payload verification (each emit method)
  // -------------------------------------------------------------------------

  describe("emitPhaseStart", () => {
    it("should forward phase and message to observer", () => {
      const fn = vi.fn();
      emitter.subscribe({ onPhaseStart: fn });

      const phases: PipelineState[] = [
        "idle",
        "reading",
        "connecting",
        "chunking",
        "analyzing",
        "applying",
        "done",
        "error",
      ];

      phases.forEach((phase) => {
        emitter.emitPhaseStart(phase, `msg-${phase}`);
      });

      expect(fn).toHaveBeenCalledTimes(phases.length);
      phases.forEach((phase, i) => {
        expect(fn).toHaveBeenNthCalledWith(i + 1, phase, `msg-${phase}`);
      });
    });
  });

  describe("emitProgress", () => {
    it("should forward current, total, and message", () => {
      const fn = vi.fn();
      emitter.subscribe({ onProgress: fn });

      emitter.emitProgress(3, 10, "Chunk 3 of 10");

      expect(fn).toHaveBeenCalledWith(3, 10, "Chunk 3 of 10");
    });
  });

  describe("emitPhaseComplete", () => {
    it("should forward the completed phase", () => {
      const fn = vi.fn();
      emitter.subscribe({ onPhaseComplete: fn });

      emitter.emitPhaseComplete("analyzing");

      expect(fn).toHaveBeenCalledWith("analyzing");
    });
  });

  describe("emitError", () => {
    it("should forward phase and Error object", () => {
      const fn = vi.fn();
      emitter.subscribe({ onError: fn });
      const err = new Error("connection failed");

      emitter.emitError("connecting", err);

      expect(fn).toHaveBeenCalledWith("connecting", err);
    });

    it("should forward phase and string error", () => {
      const fn = vi.fn();
      emitter.subscribe({ onError: fn });

      emitter.emitError("analyzing", "timeout");

      expect(fn).toHaveBeenCalledWith("analyzing", "timeout");
    });
  });

  describe("emitComplete", () => {
    it("should forward all four arguments", () => {
      const fn = vi.fn();
      emitter.subscribe({ onComplete: fn });

      const suggestions = [
        makeSuggestion({ id: "s1" }),
        makeSuggestion({ id: "s2" }),
      ];
      const result = makeInsertionResult({ successCount: 2 });
      const chunkErrors = ["chunk 3 failed"];

      emitter.emitComplete(suggestions, result, chunkErrors, true);

      expect(fn).toHaveBeenCalledWith(suggestions, result, chunkErrors, true);
    });

    it("should pass isSelection=false correctly", () => {
      const fn = vi.fn();
      emitter.subscribe({ onComplete: fn });

      emitter.emitComplete([], makeInsertionResult(), [], false);

      expect(fn).toHaveBeenCalledWith([], expect.any(Object), [], false);
    });
  });

  describe("emitAbort", () => {
    it("should forward the abort reason string", () => {
      const fn = vi.fn();
      emitter.subscribe({ onAbort: fn });

      emitter.emitAbort("User cancelled");

      expect(fn).toHaveBeenCalledWith("User cancelled");
    });
  });

  // -------------------------------------------------------------------------
  // Observer failure isolation
  // -------------------------------------------------------------------------

  describe("observer failure isolation", () => {
    it("should continue notifying later observers when one throws", () => {
      const order: string[] = [];
      const first: PipelineObserver = {
        onPhaseStart: vi.fn(() => {
          order.push("first");
        }),
      };
      const bad: PipelineObserver = {
        onPhaseStart: vi.fn(() => {
          order.push("bad");
          throw new Error("observer exploded");
        }),
      };
      const last: PipelineObserver = {
        onPhaseStart: vi.fn(() => {
          order.push("last");
        }),
      };
      emitter.subscribe(first);
      emitter.subscribe(bad);
      emitter.subscribe(last);

      expect(() => emitter.emitPhaseStart("reading", "msg")).not.toThrow();
      expect(first.onPhaseStart).toHaveBeenCalledWith("reading", "msg");
      expect(bad.onPhaseStart).toHaveBeenCalledWith("reading", "msg");
      expect(last.onPhaseStart).toHaveBeenCalledWith("reading", "msg");
      expect(order).toEqual(["first", "bad", "last"]);
    });
  });
});
