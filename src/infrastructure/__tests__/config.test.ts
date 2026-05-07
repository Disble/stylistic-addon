import { describe, expect, it } from "vitest";
import { DEFAULT_PROFILES, FEEDBACK_WORKFLOW_ID } from "../config";

describe("DEFAULT_PROFILES", () => {
  it("contains all expected profile ids", () => {
    const ids = DEFAULT_PROFILES.map((p) => p.id);
    expect(ids).toContain("general");
    expect(ids).toContain("narrativa-literaria");
    expect(ids).toContain("ensayo-academico");
    expect(ids).toContain("periodismo-cultural");
  });

  it("narrativa-literaria profile has correct label", () => {
    const fiction = DEFAULT_PROFILES.find((p) => p.id === "narrativa-literaria");
    expect(fiction).toBeDefined();
    expect(fiction?.label).toBe("Literatura de ficción");
  });

  it("has exactly 4 profiles", () => {
    expect(DEFAULT_PROFILES).toHaveLength(4);
  });

  it("narrativa-literaria is the first profile (default)", () => {
    expect(DEFAULT_PROFILES[0].id).toBe("narrativa-literaria");
  });
});

describe("FEEDBACK_WORKFLOW_ID", () => {
  it('equals "feedback-workflow"', () => {
    expect(FEEDBACK_WORKFLOW_ID).toBe("feedback-workflow");
  });
});

describe("DEFAULT_PROFILES", () => {
  it("contains all expected profile ids", () => {
    const ids = DEFAULT_PROFILES.map((p) => p.id);
    expect(ids).toContain("general");
    expect(ids).toContain("narrativa-literaria");
    expect(ids).toContain("ensayo-academico");
    expect(ids).toContain("periodismo-cultural");
  });

  it("narrativa-literaria profile has correct label", () => {
    const fiction = DEFAULT_PROFILES.find((p) => p.id === "narrativa-literaria");
    expect(fiction).toBeDefined();
    expect(fiction?.label).toBe("Literatura de ficción");
  });

  it("has exactly 4 profiles", () => {
    expect(DEFAULT_PROFILES).toHaveLength(4);
  });

  it("narrativa-literaria is the first profile (default)", () => {
    expect(DEFAULT_PROFILES[0].id).toBe("narrativa-literaria");
  });
});
