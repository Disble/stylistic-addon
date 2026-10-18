---
name: stylistic-addon-release
description: >
  Publishes a version of the add-in as a signed Stylistic Hub module.
  Trigger: release, publish a version, bump the version, tag vX.Y.Z, signed release,
  publicar versión, sacar release, nueva versión del add-in.
license: Apache-2.0
metadata:
  author: stylistic
  version: "1.0"
---

# Release

## Activation Contract

Use when publishing a new add-in version for Stylistic Hub. A release is the GitHub release `v<version>` on Disble/stylistic-addon. It holds `stylistic-word-addin-<version>.zip`, `component.json` and `component.json.sig`, which `.github/workflows/release.yml` builds and signs when a tag is pushed. The format belongs to the hub: `docs/signed-releases.md` and `docs/addin-contract.md` (Modules) in stylistic-hub `main`.

## Hard Rules

- **One version source:** `version` in `package.json`, as `major.minor.patch`. Never edit the manifest `<Version>`; every build writes it as `<version>.0`. There are no beta versions: Office needs four numeric parts, so `scripts/release/addinManifest.cjs` rejects `-beta.N`.
- **Tagging:** tag only a commit that is already on `main`, as `v<version>`. The workflow fails on any other tag. Never move or reuse a tag. A broken release is fixed with the next patch version.
- **Confirmation:** pushing the tag publishes a signed artifact that cannot be cleanly undone. Get the user's explicit yes for that exact version first. Never delete a release or tag without the user.
- **Signing:** `STYLISTIC_COMPONENT_KEY` exists only as a repository secret. Never ask for it, print it, or sign locally with it; local checks use a throwaway key.
- **Pushing:** SSH to origin fails with `publickey`, so push through gh (the user authorized it on 2026-10-18):
  `git -c credential.helper= -c "credential.helper=!gh auth git-credential" push https://github.com/Disble/stylistic-addon.git <ref>`
  That push does not update `origin/*`. Fetch the same way, then use `FETCH_HEAD`.
- **Staging:** stage explicit paths and never `git commit -a`. `.atl/skill-registry.md` carries someone else's uncommitted change.

## Decision Gates

| Condition | Action |
|---|---|
| `git status --short` shows changes beyond `.atl/skill-registry.md` and `odd/` | Stop; ask what to do with them. |
| The tag `v<version>` already exists (`git ls-remote --tags <url> v<version>`) | Bump to the next version; never re-tag. |
| The new version is not greater than the latest tag | Stop and ask. |
| CI or SonarCloud is red on the version PR | Fix it on the branch first. |
| A route group the task pane calls changed (`src/adapters/hub`) | Update `requires.scopes` in `release.yml` to match stylistic-hub `docs/api.md` (Scopes) before releasing. |
| The hub changed `docs/addin-contract.md` or `docs/signed-releases.md` on `main` | Re-read them; adapt `build:module`, the packer or the workflow first. |

## Execution Steps

1. **Prepare.** Fetch `main` through gh (see Pushing). Create `chore/release-<version>` from `FETCH_HEAD`. Set `version` in `package.json`; the lockfile does not store it. Commit as `chore(release): <version>`.
2. **Check locally.** Run `bun run validate`, `bun run test` and `bun run build:module`. Then confirm `dist/manifest.xml` has `<Version><version>.0`, `<Id>5d0c6b1e-7f2a-4c3e-9b8d-1a2f3e4c5d6e` and `SourceLocation` under `https://localhost:3000/m/stylistic-word-addin/`.
3. **Merge.** Push the branch and open a PR to `main`. Wait for `check` and SonarCloud to pass, then merge with a merge commit (`gh pr merge <n> --merge`). Wait for CI on `main` to pass.
4. **Confirm.** Ask the user: "¿Publico `v<version>` sobre `<merge sha>`?" Then STOP and wait.
5. **Tag.** Run `git tag -a v<version> <merge sha> -m "v<version>"` and push the tag (see Pushing).
6. **Watch.** Run `gh run list --workflow release.yml --limit 1` until it completes. On failure, read the log (`gh run view <id> --log-failed`). If the failure came before `gh release create`, nothing was published: fix it on `main` and release the next patch version.
7. **Verify the published assets** with the hub's own verifier and its real key list:
   ```bash
   d=$(mktemp -d) && gh release download v<version> -D "$d"
   (cd ../stylistic-hub && bun ../stylistic-addon/.claude/skills/stylistic-addon-release/verify-release.ts "$d" <version>)
   python -c "import zipfile,re,sys;z=zipfile.ZipFile(sys.argv[1]);m=z.read('manifest.xml').decode();print(sum(n.endswith('.map') for n in z.namelist()),'maps',re.search('<Id>[^<]+',m).group(),re.search('<Version>[^<]+',m).group(),re.search('SourceLocation DefaultValue=.[^\"]+',m).group())" "$d/stylistic-word-addin-<version>.zip"
   ```
   Expect `component: ok, artifact: ok`, 0 maps and the manifest values from step 2. Remove `$d` afterwards. The script reads `keys/keys.json` from the hub's `main`, but it imports the verifier from the hub's working tree, which may be on another branch. When the hub's signing code changed, check it out on `main` first.
8. **Tell the hub team** (peer session `stylistic-hub-2d`, or the maintainer). Send the release URL plus the catalog URLs:
   - `artifactUrl`: `https://github.com/Disble/stylistic-addon/releases/download/v<version>/stylistic-word-addin-<version>.zip`
   - `componentUrl`: `https://github.com/Disble/stylistic-addon/releases/download/v<version>/component.json`

   Hubs only see a version once `catalog/catalog.source.json` in stylistic-hub lists it. That is a reviewed hub-side PR, and their maintainer decides it.

## Output Contract

Report:
- the version and the tagged merge SHA;
- the PR, the `main` CI run and the release run URLs;
- the release URL;
- the verifier result (`component`/`artifact`);
- the zip size and sha256 prefix;
- that the hub team was notified, and the catalog step as pending on their side.

Report any failed or skipped step plainly.

## References

- `README.md` → Releasing
- `.github/workflows/release.yml`, `scripts/release/packComponent.mjs`, `scripts/release/addinManifest.cjs`
- stylistic-hub `main`: `docs/signed-releases.md`, `docs/addin-contract.md`, `docs/api.md` (Scopes), `keys/keys.json` (component key `a1baaad201312a0f`, scope `stylistic-addon`)
