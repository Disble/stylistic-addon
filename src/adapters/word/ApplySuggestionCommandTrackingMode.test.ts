import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ApplySuggestionCommand } from "./ApplySuggestionCommand";
import { WordTextLocatorAdapter } from "./WordTextLocatorAdapter";
import {
  installWordContext,
  makeSuggestion,
} from "./ApplySuggestionCommandTestHelper";

const textLocator = new WordTextLocatorAdapter();

describe("ApplySuggestionCommand tracking-mode guards", () => {
  beforeEach(() => {
    vi.spyOn(console, "log").mockImplementation(() => undefined);
    vi.spyOn(console, "warn").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("rejects insert-only suggestions without touching Word", async () => {
    const env = installWordContext();

    const result = await new ApplySuggestionCommand(
      makeSuggestion({ anchor: "", context: "", suggestedText: "texto sugerido" }),
      textLocator,
    ).execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Insert-only suggestions require anchor text",
    });
    expect(env.context.document.body.search).not.toHaveBeenCalled();
  });

  it("returns insertion errors without reading or mutating changeTrackingMode", async () => {
    const env = installWordContext({
      initialTrackingMode: "off",
      insertError: new Error("insert failed"),
    });

    const result = await new ApplySuggestionCommand(
      makeSuggestion(),
      textLocator,
    ).execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "insert failed",
    });
    expect(env.context.document.changeTrackingMode).toBe("off");
    expect(env.context.document.load).not.toHaveBeenCalledWith(
      "changeTrackingMode",
    );
  });

  it("does not load or read changeTrackingMode for track-change suggestions", async () => {
    const env = installWordContext({ initialTrackingMode: "trackMine" });

    const result = await new ApplySuggestionCommand(
      makeSuggestion(),
      textLocator,
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.context.document.load).not.toHaveBeenCalledWith(
      "changeTrackingMode",
    );
    expect(env.context.document.changeTrackingMode).toBe("trackMine");
  });

  it("does not load or mutate changeTrackingMode for comment-only suggestions", async () => {
    const env = installWordContext({ initialTrackingMode: "trackMine" });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({ type: "comment-only", suggestedText: undefined }),
      textLocator,
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.context.document.load).not.toHaveBeenCalledWith("changeTrackingMode");
    expect(env.context.document.changeTrackingMode).toBe("trackMine");
  });
});
