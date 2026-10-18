/* eslint-disable no-undef */

/**
 * Packs the add-in build as a signed Stylistic Hub component.
 *
 * Mirrors stylistic-hub `scripts/release/pack-component.ts` and
 * `src/shared/signing/keys.ts` (format: stylistic-hub docs/signed-releases.md):
 * `<id>-<version>.zip` of the build without source maps, its `component.json`,
 * and `component.json.sig`, an ed25519 envelope `{ keyId, signature }` over the
 * canonical JSON, signed with the PKCS#8 PEM in STYLISTIC_COMPONENT_KEY.
 *
 * Usage:
 *   node scripts/release/packComponent.mjs --dir <dir> --out <dir> --id <id> --kind <kind>
 *     --version <version> --channel <stable|beta> --release-notes <url> [--requires <json>] [--provides <json>]
 */

import { createHash, createPrivateKey, createPublicKey, sign } from "node:crypto";
import {
  createWriteStream,
  mkdirSync,
  readFileSync,
  readdirSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { join, relative, sep } from "node:path";
import { pipeline } from "node:stream/promises";
import { fileURLToPath } from "node:url";
import { parseArgs } from "node:util";
import yazl from "yazl";

const NAME = /^[a-z][a-z0-9-]{1,40}$/;
const VERSION = /^\d+\.\d+\.\d+(-beta\.\d+)?$/;
const SHA256 = /^[0-9a-f]{64}$/;
const CHANNELS = ["stable", "beta"];

/** The raw 32 bytes of an ed25519 public key. */
export function rawPublicKey(key) {
  if (key.asymmetricKeyType !== "ed25519") throw new Error("Not an ed25519 public key");
  // An ed25519 SubjectPublicKeyInfo is a fixed 12-byte header followed by the raw key.
  return key.export({ type: "spki", format: "der" }).subarray(12);
}

/** A key's id: the first 16 hex characters of the sha256 of its raw public key. */
export function keyId(publicKey) {
  return createHash("sha256").update(publicKey).digest("hex").slice(0, 16);
}

/** JSON with object keys sorted at every depth, so the signed bytes never depend on key order. */
export function canonicalJson(value) {
  if (Array.isArray(value)) return `[${value.map(canonicalJson).join(",")}]`;
  if (value && typeof value === "object") {
    const entries = Object.entries(value).sort(([a], [b]) => (a < b ? -1 : 1));
    return `{${entries.map(([key, item]) => `${JSON.stringify(key)}:${canonicalJson(item)}`).join(",")}}`;
  }
  return JSON.stringify(value);
}

/** The signature envelope of `document`: who signed it and the signature over its canonical JSON. */
function signEnvelope(document, privateKey) {
  return {
    keyId: keyId(rawPublicKey(createPublicKey(privateKey))),
    signature: sign(null, Buffer.from(canonicalJson(document)), privateKey).toString("base64"),
  };
}

/** Throws unless `component` matches the hub's `component.json` schema. */
function assertComponent(component) {
  const isObject = (value) => Boolean(value) && typeof value === "object" && !Array.isArray(value);
  const isBytes = (value) => Number.isInteger(value) && value >= 0;
  const isUrl = (value) => URL.canParse(value);
  const checks = [
    ["id", NAME.test(component.id)],
    ["kind", NAME.test(component.kind)],
    ["version", VERSION.test(component.version)],
    ["channel", CHANNELS.includes(component.channel)],
    ["sha256", SHA256.test(component.sha256)],
    ["size", isBytes(component.size)],
    ["maxUncompressedSize", isBytes(component.maxUncompressedSize)],
    ["requires", isObject(component.requires)],
    ["provides", isObject(component.provides)],
    ["releaseNotes", typeof component.releaseNotes === "string" && isUrl(component.releaseNotes)],
  ];
  const invalid = checks.filter(([, valid]) => !valid).map(([field]) => field);

  if (invalid.length > 0) {
    const fields = invalid.map((field) => `${field} ${JSON.stringify(component[field])}`);
    throw new Error(`Invalid component.json: ${fields.join(", ")}`);
  }
}

/** Zips `dir` without source maps and writes the signed `component.json` beside the zip. */
export async function packComponent({ dir, out, ...release }, privateKey) {
  const files = readdirSync(dir, { recursive: true, withFileTypes: true })
    .filter((entry) => entry.isFile() && !entry.name.endsWith(".map"))
    .map((entry) => join(entry.parentPath, entry.name))
    .sort();
  const zipfile = new yazl.ZipFile();
  let uncompressed = 0;

  for (const file of files) {
    zipfile.addFile(file, relative(dir, file).split(sep).join("/"));
    uncompressed += statSync(file).size;
  }
  zipfile.end();

  mkdirSync(out, { recursive: true });
  const zip = join(out, `${release.id}-${release.version}.zip`);
  await pipeline(zipfile.outputStream, createWriteStream(zip));

  const bytes = readFileSync(zip);
  const component = {
    ...release,
    sha256: createHash("sha256").update(bytes).digest("hex"),
    size: bytes.length,
    maxUncompressedSize: uncompressed,
  };
  assertComponent(component);

  writeFileSync(join(out, "component.json"), `${JSON.stringify(component, null, 2)}\n`);
  writeFileSync(
    join(out, "component.json.sig"),
    `${JSON.stringify(signEnvelope(component, privateKey), null, 2)}\n`
  );
  return { zip, component };
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const { values } = parseArgs({
    options: {
      dir: { type: "string" },
      out: { type: "string" },
      id: { type: "string" },
      kind: { type: "string" },
      version: { type: "string" },
      channel: { type: "string" },
      "release-notes": { type: "string" },
      requires: { type: "string", default: "{}" },
      provides: { type: "string", default: "{}" },
    },
  });
  const pem = process.env.STYLISTIC_COMPONENT_KEY;
  if (!pem)
    throw new Error("STYLISTIC_COMPONENT_KEY is not set: it holds the signing key's PKCS#8 PEM");

  const { zip, component } = await packComponent(
    {
      dir: values.dir ?? "",
      out: values.out ?? "",
      id: values.id ?? "",
      kind: values.kind ?? "",
      version: values.version ?? "",
      channel: values.channel ?? "",
      requires: JSON.parse(values.requires),
      provides: JSON.parse(values.provides),
      releaseNotes: values["release-notes"] ?? "",
    },
    createPrivateKey(pem)
  );
  console.log(`${zip}: ${component.size} bytes, sha256 ${component.sha256}`);
}
