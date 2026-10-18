/* eslint-disable no-undef */

/**
 * Manifest rewrites shared by every webpack build and the release workflow.
 *
 * `package.json` is the only version source: each build writes it into the
 * manifest as the four-part `<Version>` Office expects. A module build points
 * every URL under the path Stylistic Hub serves the module from and carries
 * the hub's production add-in id (stylistic-hub docs/addin-contract.md).
 */

/** Origin the manifest source names; Stylistic Hub serves from it too. */
const LOCAL_BASE_URL = "https://localhost:3000/";
/** Component id of the signed add-in module. */
const MODULE_ID = "stylistic-word-addin";
/** Where Stylistic Hub serves the add-in module. */
const MODULE_BASE_URL = `${LOCAL_BASE_URL}m/${MODULE_ID}/`;
/** Add-in id Word knows the installed hub's add-in by; `Stylistic.docx` is built for it. */
const PRODUCTION_ADDIN_ID = "5d0c6b1e-7f2a-4c3e-9b8d-1a2f3e4c5d6e";

/** Returns the manifest with the package version, and optionally another base URL and id. */
function transformManifest(content, { version, baseUrl = LOCAL_BASE_URL, id }) {
  if (!/^\d+\.\d+\.\d+$/.test(version)) {
    throw new Error(`package.json version "${version}" must be major.minor.patch`);
  }

  let manifest = content
    .toString()
    .replace(/<Version>[^<]*<\/Version>/, `<Version>${version}.0</Version>`);

  if (baseUrl !== LOCAL_BASE_URL) {
    manifest = manifest.split(LOCAL_BASE_URL).join(baseUrl);
  }

  if (id) {
    manifest = manifest.replace(/<Id>[^<]*<\/Id>/, `<Id>${id}</Id>`);
  }

  return manifest;
}

module.exports = {
  LOCAL_BASE_URL,
  MODULE_BASE_URL,
  MODULE_ID,
  PRODUCTION_ADDIN_ID,
  transformManifest,
};
