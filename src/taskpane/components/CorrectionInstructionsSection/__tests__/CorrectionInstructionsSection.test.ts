import { describe, expect, it, vi } from "vitest";
import { useCorrectionInstructionsSection } from "../CorrectionInstructionsSection.hooks";

vi.mock("../CorrectionInstructionsSection.styles", () => ({
  useCorrectionInstructionsSectionStyles: () => ({
    root: "root",
    header: "header",
    title: "title",
    description: "description",
    textarea: "textarea",
    footer: "footer",
    counter: "counter",
    counterLimit: "counterLimit",
  }),
}));

describe("useCorrectionInstructionsSection", () => {
  it("derives the live X / max counter label", () => {
    const state = useCorrectionInstructionsSection("hola", 4000);
    expect(state.counterLabel).toBe("4 / 4000");
  });

  it("flags isAtLimit when the value reaches the maxLength", () => {
    const state = useCorrectionInstructionsSection("x".repeat(4000), 4000);
    expect(state.isAtLimit).toBe(true);
  });

  it("does not flag isAtLimit below the maxLength", () => {
    const state = useCorrectionInstructionsSection("x".repeat(3999), 4000);
    expect(state.isAtLimit).toBe(false);
  });

  it("treats an empty string as a valid 0-length state", () => {
    const state = useCorrectionInstructionsSection("", 4000);
    expect(state.counterLabel).toBe("0 / 4000");
    expect(state.isAtLimit).toBe(false);
  });

  it("exposes the resolved Griffel classes for the parent component", () => {
    const state = useCorrectionInstructionsSection("hola", 4000);
    expect(state.classes.root).toBe("root");
    expect(state.classes.counter).toBe("counter");
    expect(state.classes.counterLimit).toBe("counterLimit");
  });
});
