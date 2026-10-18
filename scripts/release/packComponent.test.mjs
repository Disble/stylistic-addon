import { afterEach, describe, expect, it } from "vitest";
import { createHash, generateKeyPairSync, verify } from "node:crypto";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { canonicalJson, keyId, packComponent, rawPublicKey } from "./packComponent.mjs";

const scratch = [];

function makeDist() {
  const root = mkdtempSync(join(tmpdir(), "pack-component-"));
  scratch.push(root);
  const dist = join(root, "dist");
  mkdirSync(join(dist, "assets"), { recursive: true });
  writeFileSync(join(dist, "manifest.xml"), "<OfficeApp/>");
  writeFileSync(join(dist, "taskpane.html"), "<html><head></head></html>");
  writeFileSync(join(dist, "taskpane.js.map"), "{}".repeat(50));
  writeFileSync(join(dist, "assets", "icon-16.png"), "png");
  return { root, dist, out: join(root, "release") };
}

const release = {
  id: "stylistic-word-addin",
  kind: "office-addin",
  version: "1.0.0",
  channel: "stable",
  requires: { api: ["v1"], scopes: ["status"] },
  provides: {},
  releaseNotes: "https://github.com/Disble/stylistic-addon/releases/tag/v1.0.0",
};

afterEach(() => {
  scratch.splice(0).forEach((dir) => rmSync(dir, { recursive: true, force: true }));
});

describe("packComponent", () => {
  it("zips the build without source maps and signs a component.json that pins the zip", async () => {
    const { root, dist, out } = makeDist();
    const { privateKey, publicKey } = generateKeyPairSync("ed25519");

    const { zip, component } = await packComponent(
      { ...release, root, dir: dist, out },
      privateKey
    );

    const zipBytes = readFileSync(zip);
    expect(zip).toBe(join(out, "stylistic-word-addin-1.0.0.zip"));
    expect(component).toEqual({
      ...release,
      sha256: createHash("sha256").update(zipBytes).digest("hex"),
      size: zipBytes.length,
      maxUncompressedSize: "<OfficeApp/>".length + "<html><head></head></html>".length + 3,
    });
    expect(JSON.parse(readFileSync(join(out, "component.json"), "utf8"))).toEqual(component);

    const envelope = JSON.parse(readFileSync(join(out, "component.json.sig"), "utf8"));
    expect(envelope.keyId).toBe(keyId(rawPublicKey(publicKey)));
    expect(
      verify(
        null,
        Buffer.from(canonicalJson(component)),
        publicKey,
        Buffer.from(envelope.signature, "base64")
      )
    ).toBe(true);
  });

  it("refuses a component.json the hub would reject", async () => {
    const { root, dist, out } = makeDist();
    const { privateKey } = generateKeyPairSync("ed25519");

    await expect(
      packComponent({ ...release, version: "v1.0.0", root, dir: dist, out }, privateKey)
    ).rejects.toThrow('Invalid component.json: version "v1.0.0"');
    expect(existsSync(out)).toBe(false);
  });

  it("refuses a build or output directory outside the root", async () => {
    const { root, dist } = makeDist();
    const { privateKey } = generateKeyPairSync("ed25519");
    const escape = join(root, "..", "escaped-release");

    await expect(
      packComponent({ ...release, root, dir: dist, out: escape }, privateKey)
    ).rejects.toThrow(`--out "${escape}" is outside ${root}`);
    await expect(
      packComponent({ ...release, root: dist, dir: root, out: join(dist, "release") }, privateKey)
    ).rejects.toThrow(`--dir "${root}" is outside ${dist}`);
    expect(existsSync(escape)).toBe(false);
  });
});

describe("canonicalJson", () => {
  it("sorts object keys at every depth and drops whitespace", () => {
    expect(canonicalJson({ b: 1, a: { d: [2, { f: 3, e: 4 }], c: "x" } })).toBe(
      '{"a":{"c":"x","d":[2,{"e":4,"f":3}]},"b":1}'
    );
  });
});
