# Releasing AppPackager

Maintainer procedure for publishing a release. Run every command from the repository root in Git Bash with `gh` signed in. This file is excluded from the release zip.

## 1. Preconditions

- On `main`, working tree clean, `main` level with `origin/main`.
- Every commit to `main` is part of a release. A commit that is not followed by a tag in the same session, including a documentation-only commit, blocks the suite installer build until the next release. Hold such edits and commit them with the next release.
- The full test suite passes under Windows PowerShell 5.1 with Pester 5, and again under PowerShell 7. Windows PowerShell is the shipping host and the build host; a failure under either host is a defect, and the fix must be host-neutral (no change that breaks 5.1):

```bash
powershell -NoProfile -Command 'Import-Module Pester -MinimumVersion 5.0; $r = Invoke-Pester -Path Tests, Packagers -PassThru -Output None; "result {0} passed {1} failed {2} failedContainers {3} failedBlocks {4}" -f $r.Result, $r.PassedCount, $r.FailedCount, $r.FailedContainersCount, $r.FailedBlocksCount'
pwsh -NoProfile -Command 'Import-Module Pester -MinimumVersion 5.0; $r = Invoke-Pester -Path Tests, Packagers -PassThru -Output None; "result {0} passed {1} failed {2} failedContainers {3} failedBlocks {4}" -f $r.Result, $r.PassedCount, $r.FailedCount, $r.FailedContainersCount, $r.FailedBlocksCount'
```

Each run must print `result Passed` with all three failure counts at 0. A test file that fails to load raises `failedContainers`, not `failed`. Any other output stops the release.

## 2. Version

The version is `YYYY.MM.DD.BBBB`: the release date, then a four-digit, zero-padded build number that increases by 1 per release and never resets. Keep the zero-padded text everywhere; `[version]` drops the leading zeros. Set the same version in three places:

- `start-apppackager.ps1`, header line `Version    : <ver>`
- `install.ps1`, header line `Version    : <ver>`
- `README.md`, the `-Version <ver>` value in the install example

`Packagers/AppPackagerCommon.psd1` `ModuleVersion` has its own counter and changes only when the module's exported surface changes.

## 3. Changelog

The maintainer writes the `CHANGELOG.md` entry. Commit the file exactly as written.

Entry shape: `## [<ver>] - <yyyy-MM-dd>`, then `###` sections by kind of change. Bullets are verb-first, one change per line, at most 15 plain words, with no file names, ticket references or rationale. A bugfix-only release may be the single line `- Various bug fixes.`

## 4. Merge, tag and push

Do the release work from sections 2 and 3 on a branch, not on `main`:

```bash
git switch -c chore/release-<ver>
git push -u origin chore/release-<ver>
```

Open a pull request into `main`. Wait for the CodeRabbit review and CI to finish, then squash-merge it. Do not push `main` directly. Then create the tag on the merged commit and push only the tag:

```bash
git switch main && git pull --ff-only origin main
git tag -a v<ver> -m v<ver>
git push origin v<ver>
```

The tag is annotated and its message is the tag name.

## 5. Build the assets

Build from the tag, never from the working tree. Use an output folder outside the repository.

```bash
OUT=<folder outside the repo>
git archive --format=zip v<ver> -o "$OUT/AppPackager-<ver>.zip" -- . ":(exclude)Tests" ":(exclude)*.Tests.ps1" ":(exclude)RELEASING.md"
cp "$OUT/AppPackager-<ver>.zip" "$OUT/AppPackager.zip"
(cd "$OUT" && sha256sum -b AppPackager-<ver>.zip AppPackager.zip > checksums.txt)
git show v<ver>:install.ps1 | unix2dos > "$OUT/install.ps1"
```

- The exclusions on the command line keep tests and this file out of the zip. GitHub's own Source code archives are built without them and include the tests.
- `AppPackager.zip` is a byte-identical copy that serves the `releases/latest/download` bootstrap URL.
- `checksums.txt` holds exactly the two zip lines in `sha256sum` binary format (`<hash> *<name>`).
- `install.ps1` comes from the tag blob converted to CRLF. A working tree copy can carry bare LF line endings.

Check the zip before publishing. This must print 0:

```bash
unzip -Z1 "$OUT/AppPackager.zip" | grep -c -E '^Tests/|\.Tests\.ps1$|^RELEASING\.md$'
```

## 6. Release notes

Write `notes.md` in the output folder:

- No title line inside the notes; the release title is the tag.
- A `##` headline with one concrete outcome metric, for example `## 283 packagers`.
- The changelog entry's `###` sections, unchanged.
- The footer `Full changelog: [CHANGELOG.md](https://github.com/jasonulbright/app-packager/blob/v<ver>/CHANGELOG.md)`.

No emoji, no marketing wording, and nothing about signing certificates.

## 7. Publish

```bash
gh release create v<ver> --title v<ver> --notes-file "$OUT/notes.md" \
  "$OUT/AppPackager-<ver>.zip" "$OUT/AppPackager.zip" "$OUT/checksums.txt" "$OUT/install.ps1"
```

## 8. Verify the published release

```bash
mkdir -p "$OUT/verify" && cd "$OUT/verify"
gh release download v<ver> -R jasonulbright/app-packager -D . --clobber
sha256sum -c checksums.txt
unzip -p AppPackager.zip start-apppackager.ps1 | grep -m1 'Version'
git -C <repo> ls-remote --tags origin v<ver>
```

Both zips must report OK, the version line must show `<ver>`, and the remote tag must exist. The release must not be a draft and must list four assets.

## 9. After the release

Every user-visible change and packager in the release is ported to the desktop build in the same catch-up.
