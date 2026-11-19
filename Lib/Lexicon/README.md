# ConfigMgr lexicon

A data file and a checker for the user-facing text of the suite and of APD. The canonical data files are in the vault (`wiki/concepts/configmgr-lexicon/`); this folder is a vendored copy of them plus the engine. Run `sync-lexicon.ps1` in the suite root to refresh this copy from the vault and to vendor it into the other repositories. The text must use the vocabulary of Configuration Manager and the voice of Microsoft Learn.

The lexicon holds terms, rules and citations. It holds no text copied from Microsoft Learn.

## Files

| File | Purpose |
|---|---|
| `configmgr-lexicon.json` | The entries. ASCII, UTF-8 without BOM. |
| `lexicon-allowlist.json` | Justified exceptions. Each entry needs a `reason`. The file ships empty. |
| `Lexicon.psm1` | The engine: finds the user-facing text and matches it. Windows PowerShell 5.1 and PowerShell 7. |
| `Invoke-LexiconCheck.ps1` | Command line report. Use it for APD (`-RepoType Apd`) or for one repository by hand. |
| `enforce.marker` | When this file exists next to the lexicon, the Pester guard fails on a hit. Without it, the guard reports only. |

## Entry schema

```json
{
  "id": "LX-001",
  "category": "vague-noun",
  "kind": "flag",
  "preferred": ["hardware inventory", "discovery data"],
  "flagged": ["\\bfacts?\\b"],
  "allowWhen": ["\\bin fact\\b"],
  "context": "Where the word is allowed.",
  "replacement": "The required replacement, or the instruction to name the exact class, view or object.",
  "source": ["https://learn.microsoft.com/...", "code:repo/path"],
  "status": "verified",
  "severity": "error",
  "notes": ""
}
```

- `scope` is `text` (user-facing text, the default) or `tracked-files` (banned names).
- `kind` is `flag` (has `flagged` expressions) or `term` (vocabulary only).
- `flagged` and `allowWhen` are .NET regular expressions. They ignore case unless the expression starts a group with `(?-i:`.
- When one `allowWhen` expression matches the same text item, the hit is dropped.
- `source` holds a Microsoft Learn URL, `code:` plus a repository path, `vault` plus a decision, or `owner` plus a date. The guard test fails on any other value.
- `status` is `verified` (a Learn page or the code shows it) or `unverified`.
- `severity` is `error`, `warning` or `info`. Info never fails a check.
- Ids are never reused.

## What the checker reads

README.md, docs\**\*.md, the `[Unreleased]` part of CHANGELOG.md, XAML text attributes (`Content`, `Header`, `Text`, `ToolTip`, `Title`, `Watermark` and similar), strings that go to message boxes, dialogs, status lines, logs, warnings and errors in `.ps1` and `.psm1` files, `throw` messages, and comment-based help (`.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.EXAMPLE` text). It also reads the `Description`, `Synopsis` and `ReleaseNotes` of module manifests (`.psd1`) outside `Lib` and `Tests`: the manifest Description shows in `Get-Module` and in the PowerShell Gallery. For APD (`-RepoType Apd`) it also reads the XAML under `src` and sentence-like string literals in `src\**\*.cs`.

It skips code identifiers, cmdlet and class names, URLs, fenced code blocks, inline code, and the folders `Lib`, `Tests`, `installer`, `lexicon`, `.git`.

It does not read: `.nsi` installer text, JSON definition files, images.

## Run it

```powershell
Invoke-Pester -Path .\Tests\Lexicon.Guard.Tests.ps1
.\lexicon\Invoke-LexiconCheck.ps1 -Root C:\work\app-packager-desktop -RepoType Apd -OutputPath hits.csv
```

`-FailOn Error` makes the script exit with 1 when a hit has severity error. APD CI can run the script that way.

## Vendor it into another repository

Run `sync-lexicon.ps1` (it copies the `lexicon` folder to `Lib\Lexicon` in each repository, and `Tests\Lexicon.Guard.Tests.ps1` to its `Tests` folder; `-Check` reports drift). The guard finds `lexicon` or `Lib\Lexicon`. A vendored copy without `enforce.marker` reports only. Add `enforce.marker` in the fix pull request that clears the hits of that repository.

## Banned names

Entries with `"scope": "tracked-files"` (category `banned`) are matched against every line of every tracked text file, not only the user-facing text. They cover names that the owner bans from tracked files, release notes, workflow comments and commit text (vault decision D-32 of app-packager-suite). The owner extends the list: add a pattern to the entry `LX-040`. The patterns use character classes such as `[X]yz`, so that the lexicon file does not contain the banned names itself.

The entry `LX-041` (category `banned-format-detection`) reports the package file format separately, so that the owner can allow-list code that detects the format as an installer type.

`Invoke-LexiconCheck.ps1` reports both kinds of hit. The Pester guard checks `LX-040` in its own test (report-only until `enforce.marker` exists).

## Allow-list

```json
{ "schema": 1, "allow": [ { "file": "docs/reference/glossary.md", "rule": "LX-001", "phrase": "facts?", "reason": "The glossary defines the word." } ] }
```

`file` is a wildcard on the repository-relative path with forward slashes. `rule` is an id or `*`. `phrase` is an optional regular expression on the matched text. `line` is optional. `reason` is required.
