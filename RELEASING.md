# Releasing AppPackager

Maintainer procedure for publishing a release. Run every command from the repository root in Git Bash with `gh` signed in. This file is excluded from the release zip.

## 1. Preconditions

- On `main`, working tree clean, `main` level with `origin/main`.
- The full test suite passes under PowerShell 7 with Pester 5:

```bash
pwsh -NoProfile -Command "$r = Invoke-Pester -Path Tests, Packagers -PassThru -Output None; 'passed {0} failed {1}' -f $r.PassedCount, $r.FailedCount"
```

Any failure stops the release.

## 2. Version

Versions increase by 0.0.1 per release. Set the same version in three places:

- `start-apppackager.ps1`, header line `Version    : <ver>`
- `install.ps1`, header line `Version    : <ver>`
- `README.md`, the `-Version <ver>` value in the install example

`Packagers/AppPackagerCommon.psd1` `ModuleVersion` has its own counter and changes only when the module's exported surface changes.

## 3. Changelog

The maintainer writes the `CHANGELOG.md` entry. Commit the file exactly as written.

Entry shape: `## [<ver>] - <yyyy-MM-dd>`, then `###` sections by kind of change. Bullets are verb-first, one change per line, at most 15 plain words, with no file names, ticket references or rationale. A bugfix-only release may be the single line `- Various bug fixes.`

## 4. Tag and push

```bash
git tag -a v<ver> -m v<ver>
git push origin main v<ver>
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
