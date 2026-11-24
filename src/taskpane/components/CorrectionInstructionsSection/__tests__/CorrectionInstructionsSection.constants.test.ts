import { describe, expect, it } from "vitest";
import {
  CORRECTION_INSTRUCTIONS_DESCRIPTION,
  CORRECTION_INSTRUCTIONS_PLACEHOLDER,
} from "../CorrectionInstructionsSection.constants";

describe("correction instructions copy", () => {
  it("speaks neutral Spanish, without voseo", () => {
    expect(CORRECTION_INSTRUCTIONS_DESCRIPTION).toBe(
      "Reglas que quieres que el corrector vigile en todos tus documentos. " +
        "Por ejemplo: subordinadas demasiado largas, muletillas o adverbios en -mente."
    );
    expect(CORRECTION_INSTRUCTIONS_PLACEHOLDER).toBe(
      "Vigila subordinadas demasiado largas, muletillas y adverbios -mente cuando no aporten precisión."
    );
  });
});
