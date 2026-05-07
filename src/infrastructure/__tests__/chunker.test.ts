import { splitText } from "../chunker";

describe("splitText", () => {
  it("returns empty array for empty or whitespace-only input", () => {
    expect(splitText("")).toEqual([]);
    expect(splitText("   \n\n  ")).toEqual([]);
  });

  it("returns a single chunk when text is under maxChunkSize", () => {
    const chunks = splitText("Hello world.\n\nSecond paragraph.", 1000);

    expect(chunks).toHaveLength(1);
    expect(chunks[0].text).toBe("Hello world.\n\nSecond paragraph.");
    expect(chunks[0].index).toBe(0);
    expect(chunks[0].total).toBe(1);
    expect(chunks[0].startOffset).toBe(0);
  });

  it("splits into multiple chunks at paragraph boundaries", () => {
    const text = "Paragraph one.\n\nParagraph two.\n\nParagraph three.";
    const chunks = splitText(text, 30);

    expect(chunks.length).toBeGreaterThan(1);
    for (const chunk of chunks) {
      expect(chunk.total).toBe(chunks.length);
      expect(chunk.index).toBeGreaterThanOrEqual(0);
      expect(chunk.text.length).toBeGreaterThan(0);
    }
  });
});
