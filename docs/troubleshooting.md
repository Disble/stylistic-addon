# Troubleshooting

Common issues and their solutions when developing or using Stylistic.

## Development Issues

### `bun run start` fails with certificate errors

**Symptom:** The dev server starts but Word rejects the connection, or the browser shows `ERR_CERT_AUTHORITY_INVALID`.

**Cause:** The self-signed development certificates haven't been generated or trusted.

**Fix:**

```bash
npx office-addin-dev-certs install
```

This generates and trusts a local CA certificate. Restart with `bun run start` after installing.

---

### Task pane shows "Please sideload your add-in"

**Symptom:** The add-in loads in Word but the task pane only shows the sideload message.

**Cause:** `Office.onReady()` didn't detect `Office.HostType.Word`, which usually means the add-in wasn't properly sideloaded.

**Fix:**

1. Make sure you ran `bun run start` (not just the dev server).
2. If using Word Online, manually sideload: **Insert** > **Add-ins** > **Upload My Add-in** > select `manifest.xml`.
3. Check the browser console (F12) for errors — the Office.js CDN may be blocked.

---

### Changes don't appear after editing source files

**Symptom:** You modified `.ts` or `.html` files but the task pane shows old content.

**Cause:** Webpack may not have rebuilt, or the task pane is cached.

**Fix:**

1. If using `bun run watch`, check that webpack recompiled (look for output in the terminal).
2. Close and reopen the task pane in Word.
3. Hard-refresh: in Word Online, press `Ctrl+Shift+R`. In Word Desktop, close and reopen the document.

---

### `bun run validate` reports manifest errors

**Symptom:** The validator flags issues in `manifest.xml`.

**Common causes:**

- **Missing icon files:** Ensure all referenced PNG files exist in `assets/` at the correct sizes (16, 32, 64, 80, 128px).
- **URL mismatch:** In production builds, update the `urlProd` variable in `webpack.config.js` from `https://www.contoso.com/` to your actual deployment URL.
- **Requirement set version:** The manifest requires WordApi 1.6. If testing on an older version of Word, this will fail at load time (not validation time).

**Note:** For manifest-specific validation, run `bun run manifest:validate`. The `bun run validate` command runs lint + architecture rails + filename checks + complexity checks + React rails + typecheck. It does **not** validate the manifest.

---

### `bun run build` fails with TypeScript errors after adding new files

**Symptom:** `bun run build` fails with type errors in new `.ts` files.

**Cause:** The Babel-based build pipeline doesn't perform type checking — it strips types and transpiles. If you see TypeScript errors, they come from the editor (VS Code) or a separate `tsc` check.

**Fix:**

1. Make sure your new file is under `src/` (not excluded in `tsconfig.json`).
2. Verify imports use relative paths (e.g., `"../lib/types"`, not absolute paths).
3. Run `bun run typecheck` to see all type errors.

---

## UI / Layout Issues

### Taskpane content sticks to the top and the rest stays grey

**Symptom:** The white card with the workflow only spans the height of its
content. Below it, the body background color (grey `#fafafa`) is visible. The
`SettingsToolbar` does not pin to the bottom.

**Cause:** The full-height chain is broken. `#app-body` declares
`height: 100%`, but its parents must each be `100%` too — including the `<div>`
that `FluentProvider` injects between `#container` and the React subtree. If
any link in the chain is auto-sized, `#app-body` collapses to its content
height and the workflow's `flex: 1` has no remaining space to grow into.

**Fix:** keep this rule in `taskpane.css` so the chain reaches `#app-body`:

```css
#container,
#container > * {
  height: 100%;
}
```

`> *` covers the `FluentProvider` wrapper without coupling to the internal
`.fui-FluentProvider` class.

---

### Vertical scrollbar appears on the taskpane even though content fits

**Symptom:** A scrollbar shows up on the right edge of the taskpane while the
visible content clearly fits in the viewport (e.g. only the profile selector
and analyze button are rendered).

**Cause:** `#app-body` declares `height: 100%` *and* a non-zero `padding`. The
default `box-sizing` is `content-box`, so the rendered height is
`100% + padding-top + padding-bottom`, overflowing the parent by the padding
amount and forcing the scrollbar.

**Fix:** add `box-sizing: border-box` to `#app-body` (already in
`taskpane.css`). Whenever you introduce a new full-height container with
padding, repeat this — `content-box` is the silent killer of Office task pane
layouts.

---

## Backend Connection Issues

### Login dialog closes with error 12006 but Google eventually succeeds

**Symptom:** The taskpane reports that the authentication dialog closed before
login completed, while the dialog or backend logs show the OAuth callback/session
eventually succeeded.

**Cause:** Do not treat every `DialogEventReceived` 12006 as a fatal login
failure. In real Word hosts, 12006 can appear while the dialog is still
navigating. The confirmed implementation resolves login only from
`DialogMessageReceived`, following the official Office fallback-auth sample.

**Fix:** Verify `OfficeDialogAuthAdapter` does not reject from
`DialogEventReceived`. The dialog page must load Office.js, wait for
`Office.onReady()`, and post `{ type: "stylistic-auth-success", session }` with
`Office.context.ui.messageParent`.

---

### SonarQube warns about missing SRI on `office.js`

**Symptom:** Sonar raises `Web:S5725` on `taskpane.html`, `auth-dialog.html`, or
`commands.html` because the Office.js CDN script does not declare
`integrity="..."`.

**Cause:** Microsoft documents Office.js as a CDN-hosted script that must be
loaded from `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` inside
`<head>`. The platform is versioned by the `/1/` channel and served by Microsoft,
but Microsoft does **not** publish a stable SRI hash for this URL. Adding a hash
we compute ourselves would be brittle because the CDN can legitimately update the
payload while preserving the supported `/1/` contract.

**Safe decision:** Keep the official Microsoft CDN reference without SRI unless
Microsoft starts publishing supported integrity metadata for Office.js. Do **not**
pin an ad-hoc hash or mirror the file locally for production just to silence the
warning; either option risks breaking Office host compatibility or missing vendor
updates.

Reference: Microsoft Learn — *Referencing the Office JavaScript API library from
its CDN*.

---

### Login reaches Google but returns `state_security_mismatch`

**Symptom:** Google redirects back to the backend, but backend logs show Better
Auth errors similar to `state_security_mismatch` or `State not persisted
correctly`.

**Cause:** The Office Dialog flow crosses taskpane/backend/provider runtimes, and
the temporary signed OAuth state cookie can be unavailable on callback.

**Fix:** Backend auth must keep OAuth state in the database with:

```ts
account: {
  storeStateStrategy: "database",
  skipStateCookieCheck: true,
}
```

The `verification` table must also exist. In the backend, run:

```bash
bun run db:auth:apply
```

---

### Login succeeds but analysis says the backend is unauthorized

**Symptom:** The taskpane shows an active session, but workflow calls fail with
401/unauthorized.

**Cause:** The Mastra client was created without the latest Better Auth bearer
token, or a stale client instance kept old headers after login/logout.

**Fix:** Mastra workflow adapters must use `MastraClientFactory`, which creates a
client with the current token snapshot for each call. Do not reintroduce a module
level `new MastraClient({ baseUrl })` singleton.

---

### "Backend no disponible"

**Symptom:** Clicking "Analizar y sugerir" shows this error immediately.

**Cause:** The Mastra server is not running or the `stylistic-workflow` is not registered.

**Fix:**

1. Verify the Mastra server is running at the URL configured in `src/infrastructure/config.ts` (default: `http://localhost:4111`).
2. Check that the `stylistic-workflow` is registered in the Mastra configuration.
3. Verify CORS is enabled for `https://localhost:3000` on the Mastra server.
4. Check the browser console (F12) for CORS or network errors.

---

### Analysis takes too long or times out

**Symptom:** The progress bar stalls on a chunk, and eventually an error is reported.

**Cause:** The backend AI processing took longer than expected, or the chunk is too large.

**Fix:**

1. Check the Mastra server logs for errors or slow responses.
2. Reduce `DEFAULT_MAX_CHUNK_SIZE` in `src/infrastructure/config.ts` to send smaller chunks.
3. Ensure the backend model has sufficient resources for processing.

---

### The backend run completed, but the taskpane still shows `Reintentar análisis`

**Symptom:** Real host evidence shows the Mastra run already finished, but the
analysis error surface renders the full resubmit CTA instead of
`Reintentar consulta`.

**Cause:** The retry-mode decision in the taskpane is driven by whether the
frontend still has a retryable `runId`, not by the button copy itself. A real
Mastra `runById(runId)` response can briefly report `status: "success"` before
the success payload is readable/valid for the frontend. If the adapter collapses
that shape into a terminal `failed`, the taskpane loses the preserved `runId`
and falls back to the full resubmit path even though the backend workflow already
completed.

**Fix:** Treat `status: "success"` with an unreadable payload as a
frontend-local retryable query failure. Keep the original `runId`, surface the
retry-query hero, and let the user re-poll the same backend run without
re-submitting chunks.

---

### Some chunks fail but others succeed

**Symptom:** The results show "N fragmento(s) con error" alongside successful suggestions.

**Cause:** Individual chunks failed after retries (network issue, backend overload, or model error).

**Expected behavior.** The add-in applies suggestions from successful chunks and reports failures. No manual retry is needed for the successful portion — those Track Changes are already in the document.

---

## Runtime Issues

### "El documento está protegido o es de solo lectura"

**Symptom:** The analysis completes but Track Changes cannot be applied.

**Cause:** The document has editing restrictions (DRM, password protection, or read-only mode).

**Fix:** Remove the protection from the document (**Review** > **Restrict Editing** > **Stop Protection**), or open a non-protected copy.

---

### Suggestions don't appear as tracked changes

**Symptom:** The results panel shows suggestions as "applied", but no tracked changes are visible in the document.

---

### Analysis fails with a document-identity error

**Symptom:** Analysis or feedback fails with an error saying Stylistic could not
read or persist the document identity.

**Cause:** The add-in now requires `Office.context.document.settings` to persist
the stable `documentUuid` used by both workflows. If the host does not expose
document settings, or if saving settings fails, the frontend fails closed.

**Fix:**

1. Verify the add-in is running in a Word host that supports document settings.
2. Make sure the document is editable and not blocked by host restrictions that
   prevent saving settings metadata.
3. If the host is expected to support settings, inspect the Office runtime and
   console for `Office.context.document.settings` availability and `saveAsync`
   failures.

**Possible causes:**

1. **Track Changes display is off.** Go to **Review** > **All Markup** (ensure it's not set to "No Markup" or "Original").
2. **The suggestion matched and replaced identical text.** If `anchor` and `suggestedText` are the same after casing, Word won't show a change.
3. **The tracking mode was already `TrackAll`.** The changes were made and are there — check the Review pane (**Review** > **Reviewing Pane**).

**What to inspect in developer logs:** The expected evidence depends on the
Track Changes subtype:

- **replace:** pre-mutation scope should report `changeTrackingMode: 'TrackAll'`,
  and the annotation resolver should select a candidate whose `current` equals
  the suggested text and whose `original` is `''`. If the replacement is visible
  but `original` is also the suggested text, Word likely made an untracked
  replacement; that is a Track Changes lifecycle bug, not a display issue.
- **delete-only:** `suggestedText` is `""`. The mutation range can be empty in
  real Word; inspect the operational wrapper reviewed text/tracked changes
  instead of expecting an inserted-side Content Control.
- **formatting:** `suggestedText` may be markdown transport such as `*post
  mortem*` or `**PRIME**`. Word should expose a `Formatted` tracked change and
  the target range font state should change; `reviewedText.current` and
  `reviewedText.original` usually remain equal because text content did not
  change.

Apply-time operational wrapper creation must not be used to toggle Track Changes
off and back on. The wrapper only defines mutation/identity scope; the batch
apply workflow owns Track Changes activation.

---

### "No se encontraron sugerencias editoriales"

**Symptom:** The status bar shows this message even though the document contains text.

**Cause:** The backend workflow didn't find any editorial issues in the text.

**Expected behavior.** This is a valid result — the text may already be well-written for the selected genre.

---

### Many suggestions show as "No encontrado"

**Symptom:** The results panel marks many suggestions as not found in the document.

**Possible causes:**

1. **The text was already replaced** by a previous suggestion. Since each suggestion is applied in its own `Word.run`, earlier replacements may remove text that later suggestions target.
2. **Backend returned non-exact matches.** The `anchor` from the workflow must be an exact, case-sensitive substring of the document text. See [api-contract.md](api-contract.md) for details.
3. **Hidden characters.** The document may contain non-breaking spaces, soft hyphens, or other invisible characters that break the match.

---

### Clicking a suggestion does not move the cursor

**Symptom:** Clicking a suggestion card in the taskpane leaves the Word cursor in
place and the card shows an informational note similar to
`(no se pudo ubicar la sugerencia de forma segura)`.

**Expected behavior.** Navigation is now intentionally conservative. The add-in
will only move the cursor when it can locate a safe target:

1. for `track-change`, the operational wrapper Content Control with valid
   subtype-aware operational-wrapper metadata,
2. for `comment-only`, the canonical comment-only Content Control,
3. if the artifact is missing, the exact `anchor` inside the localized `context`.

If the artifact is ambiguous/corrupt, Word cannot be queried, or the context
cannot be localized, the add-in refuses to navigate. This prevents the older bug
where a global anchor search could select an unrelated occurrence in the table of
contents or a heading.

**What to inspect:** Check whether the document still contains the Stylistic
Content Controls for the suggestion. If metadata is missing, verify that the
backend suggestion still satisfies `context.includes(anchor)` and that the
context text exists in the document after user edits.

---

### Accept/Reject appears to do nothing

**Symptom:** Clicking **Accept** or **Reject** appears inert, or the card enters
resolving state briefly and then returns without a visible document change.

**Common real-host cause:** The resolution workflow successfully locates the
Stylistic artifact, but a later step reads a proxy-backed Word property such as
`ContentControl.tag` that was never loaded. Real Word then throws an error like:

> `The property 'tag' is not available. Before reading the property's value, call the load method on the containing object and call "context.sync()" on the associated request context.`

This can look like an inert button when the taskpane only surfaces a generic
resolution failure.

**Why tests can miss it:** permissive Office.js mocks often expose `tag` and
`title` as plain fields, so a GREEN suite does not prove the real host loaded
those properties correctly.

**What to inspect in logs:**

1. Confirm the workflow emitted `locate succeeded` for the attempt.
2. Check the next failure event or caught error metadata for an unloaded proxy
   property (`tag`, `title`, etc.).
3. Inspect the artifact locator and verify it loads every identity field that
   downstream resolution steps read.

**Fix direction:** treat the locator as the contract boundary. If downstream code
reads `selectedCc.tag` or `selectedCc.title`, the locator must load those
properties before returning the selected artifact bundle.

---

## Comment Cleanup Issues

### Accepted/rejected suggestions leave invisible Stylistic metadata

**Symptom:** A suggestion appears accepted or rejected in Word, but later runs fail
to insert or resolve nearby suggestions, or OOXML inspection still shows markers
such as `stylistic:track-change:{id}` or
`stylistic-operational-wrapper:{id}`.

**Cause:** Native Word tracked-change resolution does not guarantee removal of
the add-in-owned metadata Content Controls. If those wrappers are deleted while
Track Changes is enabled, Word can preserve the cleanup itself as another pending
revision, leaving metadata residue in the document package.

**Expected production behavior:** After a `track-change` suggestion is accepted
or rejected, the resolution workflow now:

1. resolves the native tracked changes,
2. deletes the colocated Stylistic comment,
3. temporarily disables Track Changes,
4. deletes exact metadata Content Controls for the resolved suggestion with
   `delete(true)`,
5. restores the previous Track Changes mode,
6. only then computes the final document review state.

This cleanup is exact-tag based and must not delete unrelated user comments,
foreign Content Controls, or metadata for other pending Stylistic suggestions.

**If residue persists:** Capture the browser console logs for the
`cleanup-metadata` phase and inspect the reported `deletedContentControlCount`,
`deletedContentControls`, and any failure metadata. Real Word evidence wins over
green tests; update the mock model first if the host disproves the current
contract.

---

### "Limpiar comentarios" button doesn't appear

**Symptom:** After applying suggestions, the cleanup button is not visible.

**Cause:** No suggestions were successfully applied. The button only appears when `successCount > 0`.

**Fix:** Verify that suggestions were applied by checking the results panel.

---

### Cleanup deletes 0 comments

**Symptom:** Clicking "Limpiar comentarios resueltos" shows "0 eliminado(s)".

**Cause:** All tracked changes are still pending — none have been accepted or rejected yet.

**Expected behavior.** The cleanup only removes comments for resolved tracked changes. Accept or reject some tracked changes first, then click cleanup.

---

### Cleanup keeps comments it shouldn't

**Symptom:** After accepting all tracked changes, some Stylistic comments remain.

**Possible cause:** The comment cleanup uses range colocation (comparing the document position of comments and tracked changes). If a comment's anchor range was shifted by other edits, the position comparison may not detect it as orphaned.

**Fix:** Run the cleanup again — after the first pass removes some comments, the remaining ones may now be correctly identified. If comments persist, manually delete them from the Review pane.

---

## Word Online Specific

### Task pane doesn't appear after sideloading

**Fix:** After uploading the manifest, close and reopen the document. Then go to **Home** tab and look for the **Show Task Pane** button in the ribbon.

---

### Tracked changes work but look different from desktop

**Expected behavior.** Word Online has a slightly different Track Changes UI (the strikethrough and underline rendering may differ). The changes are functionally identical and will appear correctly when the document is opened in Word Desktop.

---

## Getting Help

If your issue isn't listed here:

1. Check the browser console (F12) for JavaScript errors.
2. Run `bun run manifest:validate` to check the manifest.
3. Run `bun run validate` to run lint + rails + typecheck without building.
4. Try `bun run stop && bun run start` for a clean restart.
4. Verify the Mastra server is running and responsive.
5. Open an issue in the repository with:
   - The error message (screenshot or text)
   - Your Word version (Desktop/Online, version number)
   - Your OS and Node.js version
   - Whether the Mastra backend is running
