import { describe, expect, it } from "vitest";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const { transformManifest, MODULE_BASE_URL, PRODUCTION_ADDIN_ID } = require("./addinManifest.cjs");

const manifest = [
  "<Id>b821964a-4f1e-4939-83c8-0ddba48f3dff</Id>",
  "<Version>1.0.0.0</Version>",
  '<SupportUrl DefaultValue="https://localhost:3000"/>',
  '<SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>',
  '<bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html"/>',
].join("\n");

describe("transformManifest", () => {
  it("writes the package version as the four-part manifest version", () => {
    const result = transformManifest(manifest, { version: "1.4.2" });

    expect(result).toContain("<Version>1.4.2.0</Version>");
    expect(result).toContain('DefaultValue="https://localhost:3000/taskpane.html"');
    expect(result).toContain("<Id>b821964a-4f1e-4939-83c8-0ddba48f3dff</Id>");
  });

  it("points a module build under /m/stylistic-word-addin/ with the production id", () => {
    const result = transformManifest(manifest, {
      version: "1.0.0",
      baseUrl: MODULE_BASE_URL,
      id: PRODUCTION_ADDIN_ID,
    });

    expect(MODULE_BASE_URL).toBe("https://localhost:3000/m/stylistic-word-addin/");
    expect(result).toContain(
      'DefaultValue="https://localhost:3000/m/stylistic-word-addin/taskpane.html"'
    );
    expect(result).toContain(
      'DefaultValue="https://localhost:3000/m/stylistic-word-addin/commands.html"'
    );
    expect(result).toContain("<Id>5d0c6b1e-7f2a-4c3e-9b8d-1a2f3e4c5d6e</Id>");
  });

  it("rejects a version that is not major.minor.patch", () => {
    expect(() => transformManifest(manifest, { version: "1.0.0-beta.1" })).toThrow(
      'package.json version "1.0.0-beta.1" must be major.minor.patch'
    );
  });
});
