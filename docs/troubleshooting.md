# Troubleshooting

Common issues and their solutions when developing or using Stylistic.

## Development Issues

### `npm start` fails with certificate errors

**Symptom:** The dev server starts but Word rejects the connection, or the browser shows `ERR_CERT_AUTHORITY_INVALID`.

**Cause:** The self-signed development certificates haven't been generated or trusted.

**Fix:**

```bash
npx office-addin-dev-certs install
```

This generates and trusts a local CA certificate. Restart with `npm start` after installing.

---

### Task pane shows "Please sideload your add-in"

**Symptom:** The add-in loads in Word but the task pane only shows the sideload message.

**Cause:** `Office.onReady()` didn't detect `Office.HostType.Word`, which usually means the add-in wasn't properly sideloaded.

**Fix:**

1. Make sure you ran `npm start` (not just the dev server).
2. If using Word Online, manually sideload: **Insert** > **Add-ins** > **Upload My Add-in** > select `manifest.xml`.
3. Check the browser console (F12) for errors — the Office.js CDN may be blocked.

---

### Changes don't appear after editing source files

**Symptom:** You modified `.ts` or `.html` files but the task pane shows old content.

**Cause:** Webpack may not have rebuilt, or the task pane is cached.

**Fix:**

1. If using `npm run watch`, check that webpack recompiled (look for output in the terminal).
2. Close and reopen the task pane in Word.
3. Hard-refresh: in Word Online, press `Ctrl+Shift+R`. In Word Desktop, close and reopen the document.

---

### `npm run validate` reports manifest errors

**Symptom:** The validator flags issues in `manifest.xml`.

**Common causes:**

- **Missing icon files:** Ensure all referenced PNG files exist in `assets/` at the correct sizes (16, 32, 64, 80, 128px).
- **URL mismatch:** In production builds, update the `urlProd` variable in `webpack.config.js` from `https://www.contoso.com/` to your actual deployment URL.
- **Requirement set version:** The manifest requires WordApi 1.6. If testing on an older version of Word, this will fail at load time (not validation time).

**Note:** For manifest-specific validation, run `npm run manifest:validate`. The `npm run validate` command runs lint + filename checks + complexity checks.

---

### Build fails with TypeScript errors after adding new files

**Symptom:** `npm run build` fails with type errors in new `.ts` files.

**Cause:** The Babel-based build pipeline doesn't perform type checking — it strips types and transpiles. If you see TypeScript errors, they come from the editor (VS Code) or a separate `tsc` check.

**Fix:**

1. Make sure your new file is under `src/` (not excluded in `tsconfig.json`).
2. Verify imports use relative paths (e.g., `"../lib/types"`, not absolute paths).
3. Run `npx tsc --noEmit` to see all type errors.

---

## Backend Connection Issues

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

**Possible causes:**

1. **Track Changes display is off.** Go to **Review** > **All Markup** (ensure it's not set to "No Markup" or "Original").
2. **The suggestion matched and replaced identical text.** If `anchor` and `suggestedText` are the same after casing, Word won't show a change.
3. **The tracking mode was already `TrackAll`.** The changes were made and are there — check the Review pane (**Review** > **Reviewing Pane**).

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
   `compound-v2` metadata,
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
2. Run `npm run manifest:validate` to check the manifest.
3. Run `npm run validate` to run lint + filename checks + complexity checks.
3. Try `npm stop && npm start` for a clean restart.
4. Verify the Mastra server is running and responsive.
5. Open an issue in the repository with:
   - The error message (screenshot or text)
   - Your Word version (Desktop/Online, version number)
   - Your OS and Node.js version
   - Whether the Mastra backend is running
