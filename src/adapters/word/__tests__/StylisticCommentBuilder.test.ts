import { describe, expect, it } from "vitest";

import { buildStylisticCommentContent, isStylisticComment } from "../StylisticCommentBuilder";

describe("StylisticComment helpers", () => {
  it("builds the persisted comment content shape", () => {
    expect(buildStylisticCommentContent("Claridad", "Mas claro")).toBe("[Claridad]\nMas claro");
  });

  it("identifies Stylistic comments by content even with a human author name", () => {
    expect(
      isStylisticComment({
        authorName: "Usuario de prueba",
        content: "[gramática]\nRedundancia pronominal.",
      })
    ).toBe(true);
  });

  it("identifies Stylistic comments when Word returns CRLF line endings", () => {
    expect(
      isStylisticComment({
        authorName: "Usuario de prueba",
        content: "[gramática]\r\nRedundancia pronominal.",
      })
    ).toBe(true);
  });

  it("identifies Stylistic comments when Word returns CR line endings", () => {
    expect(
      isStylisticComment({
        authorName: "Usuario de prueba",
        content: "[gramática]\rRedundancia pronominal.",
      })
    ).toBe(true);
  });

  it("does not treat authorName alone as a Stylistic signal", () => {
    expect(
      isStylisticComment({
        authorName: "Stylistic",
        content: "Comentario cualquiera",
      })
    ).toBe(false);
  });
});
