# Authoring a new packager

A packager is a single `package-<app>.ps1` file in `Packagers/`. It has two phases: **Stage** (downloads installer, extracts metadata, writes wrappers + manifest locally) and **Package** (reads manifest, copies to network, creates MECM Application). The GUI and the One Click workflow call the same two functions.

Every packager follows the same skeleton. Copy one of the templates from this folder and swap in the app-specific bits.

```
Samples/
  package-template-msi.ps1      <-- start here for MSI installers
  package-template-exe.ps1      <-- start here for EXE installers
  AUTHORING.md                  <-- this file
```

Representative existing packagers to reference:

| Pattern | Example |
|---|---|
| MSI with vendor-page scrape | `package-7zip.ps1` |
| MSI with direct download URL | `package-chrome.ps1` (Google publishes a stable MSI URL) |
| EXE with vendor-page scrape | `package-audacity.ps1` |
| EXE with GitHub releases API | `package-bitwarden.ps1`, `package-keepass.ps1` |
| EXE with extracted ARP metadata | `package-adobereader.ps1` (Adobe ships a self-extracting archive) |
| MSI with company-wide ODT config | `package-m365apps-x64.ps1` (reads Packager Preferences) |
| MSIX/APPX enterprise installer | `package-teams-new.ps1` |
| User-context file detection | `package-vscode-user.ps1`, `package-postman.ps1` |

## The header block

```powershell
<#
Vendor: Igor Pavlov
App: 7-Zip (x64)
CMName: 7-Zip
VendorUrl: https://www.7-zip.org/
CPE: cpe:2.3:a:7-zip:7-zip:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.7-zip.org/history.txt
DownloadPageUrl: https://www.7-zip.org/download.html
UpdateCadenceDays: 90
#>
```

The GUI parses these tags with `Get-PackagerMetadata`:

| Tag | Required | Purpose |
|---|---|---|
| `Vendor` | Yes | Main grid Vendor column; network content path (`\\share\Applications\<Vendor>\<App>\<Version>`) |
| `App` | Yes | Main grid Application column; network content path `<App>` segment |
| `CMName` | No (defaults to App) | Name used when querying MECM for the currently-deployed version. Some apps have MECM names that differ from their product names |
| `VendorUrl` | No | Ctrl+Click in main grid opens this URL |
| `CPE` | No | NVD Common Platform Enumeration string used by Version Monitor for CVE lookups |
| `ReleaseNotesUrl` | No | Shown in Version Monitor HTML report Links column |
| `DownloadPageUrl` | No | Shown in Version Monitor HTML report Links column |
| `UpdateCadenceDays` | No (default 7) | Default cadence for One Click Report runs. Per-app overrides live in One Click Settings |

## The param block

Every packager accepts the same parameters so the GUI can invoke any of them uniformly. Copy-paste this block verbatim:

```powershell
param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 15,
    [int]$MaximumRuntimeMins = 30,
    [string]$LogPath,
    [switch]$GetLatestVersionOnly,
    [switch]$StageOnly,
    [switch]$PackageOnly
)
```

Do not add or rename parameters that the GUI passes. Packager-specific configuration belongs in Packager Preferences (see `OPTIONS_AUTHORING.md`) or in the packager's inline config block.

## The four operating modes

| Invocation | Behavior |
|---|---|
| no switches | Stage then Package |
| `-StageOnly` | Stage only |
| `-PackageOnly` | Package only |
| `-GetLatestVersionOnly` | Write just the latest version string to stdout and exit. Prefer no logging, no staging, and a lightweight vendor query |

The GUI uses `-GetLatestVersionOnly` to populate the Latest column in the grid and to evaluate cadence in One Click. It should be fast and network-light. Prefer a vendor API, release feed, or HTML scrape. If the vendor exposes the authoritative version only inside the installer or package manifest, keep the file operation bounded, cache under `DownloadRoot`, and make that behavior obvious in comments and docs.

## Stage phase

Stage does everything that can be done locally, without the MECM console and without touching the network share. In order:

1. Resolve the current version (scrape the vendor page, hit a release API, or read a known-stable URL).
2. Download the installer to `<DownloadRoot>\<AppSubfolder>\<installer>`.
3. Derive detection metadata. For MSIs use `Get-MsiPropertyMap` - `ProductName`, `ProductVersion`, `ProductCode`, `Manufacturer` are written directly into registry at install time. The ARP uninstall key for a standard MSI is `SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\<ProductCode>`, so detection never requires a temp install.
4. Create a versioned content folder at `<DownloadRoot>\<AppSubfolder>\<Version>\`. Copy the installer in.
5. Call `Write-ContentWrappers` with install/uninstall `.ps1` content. The shared module auto-generates `install.bat` and `uninstall.bat` wrappers around them.
6. Call `Write-StageManifest` to serialize AppName / Publisher / SoftwareVersion / Detection / InstallArgs into `stage-manifest.json`.

The stage manifest is the contract between Stage and Package. Package reads it and never needs to re-download the installer, re-parse MSI properties, or re-probe the vendor.

## Package phase

Package does everything that requires the MECM console and the network share:

1. Read the stage manifest.
2. Verify `FileServerPath` is reachable and writable (`Test-NetworkShareAccess`).
3. Compute the network content path via `Get-NetworkAppRoot`: `<FileServerPath>\Applications\<Vendor>\<App>\<Version>\`.
4. Copy the staged content folder to the network path (skip `stage-manifest.json` - it stays local).
5. Call `New-MECMApplicationFromManifest` with the manifest. The shared module creates the CM Application, adds the Script deployment type with the right detection method, applies the right runtime limits, and (if Auto-distribute is configured in MECM Preferences) kicks off content distribution to the DP group.

`New-MECMApplicationFromManifest` is idempotent but fail-closed: if an application with the same name exists in MECM, it validates that the expected deployment type also exists and returns the existing CI_ID. If the application exists without the expected deployment type, it throws because that usually means a prior package run partially created the app. Fix or remove the partial application before re-running Package.

## Detection types

The stage manifest's `Detection` block tells the shared module which MECM detection method to create. Five types are supported:

| Type | Used when | Fields |
|---|---|---|
| `RegistryKeyValue` | MSI-based installs (ARP uninstall key has `DisplayVersion`) | `RegistryKeyRelative`, `ValueName`, `ExpectedValue` or `DisplayVersion`, `Operator`, `Is64Bit` |
| `RegistryKey` | Install writes a registry key but no reliable version value | `RegistryKeyRelative`, `Is64Bit` |
| `File` | Install drops a versioned EXE in a well-known path | `Path`, `FileName`, `Version`, `Operator` |
| `Script` | No registry or file works; need a custom PowerShell detection snippet | `ScriptBlock` |
| `Compound` | Multiple clauses joined by AND or OR | `Clauses`, `Connector` |

For MSIs, `RegistryKeyValue` with the `ProductCode` as the key name is always correct and never requires a temp install. See `package-7zip.ps1` for the canonical pattern.

For EXE installers, you usually have to install once into a disposable VM, note where it writes ARP or a versioned binary, then code that path into the packager.

## Version string formatting

The GUI's Latest column and the One Click cadence gate compare version strings with `Compare-SemVer`. That helper normalizes two inputs to the same significant-part count before comparing, so `26.00` and `26.00.00.0` compare equal. Pick one format and use it consistently in:

- the string returned by `-GetLatestVersionOnly`
- the `SoftwareVersion` field in the stage manifest
- the `DisplayVersion` field inside the Detection block

If the vendor's raw version differs from what Windows writes to the ARP registry (as with 7-Zip - display `26.00`, registry `26.00.00.0`), store the raw value under `DisplayVersion` in the detection block and the display value under `SoftwareVersion`.

## Scraping patterns

**Vendor download page regex.** Fetch the HTML with `curl.exe -L --fail --silent --show-error`. Use a narrow regex that targets the href attribute and captures the version as digits. Sort candidates by the captured version and take the highest. Resolve the relative href against the page's base URL. See `Resolve-7ZipX64MsiUrl` in `package-7zip.ps1`.

**GitHub releases API.** For apps published on GitHub, hit `https://api.github.com/repos/<owner>/<repo>/releases/latest` and parse the JSON. The `tag_name` field is usually the version; `assets[]` contains the download URLs. See `package-bitwarden.ps1`, `package-keepass.ps1`.

**Direct stable URL.** Google Chrome Enterprise, Microsoft Edge, and a few others publish an MSI at a URL that always redirects to the current version. You can download it unconditionally and read the version from the MSI's `ProductVersion`. See `package-chrome.ps1`.

**Vendor JSON/API endpoint.** Audacity, Chrome's Version History API, and similar. `Invoke-RestMethod` the endpoint, dig out the version field. See `package-audacity.ps1`, `package-chrome.ps1`.

**Avoid** beta/nightly channel pages, non-deterministic mirror lists, and pages that require authentication. If the vendor puts the current version behind a login, the app is not a candidate for an automated packager.

## Logging

`Initialize-Logging -LogPath $LogPath` sets up the log sink. All subsequent `Write-Log` calls write to both the console (so the GUI log pane streams them live via `-LogTextBox`) and the log file. Use `-Level ERROR` for fatal conditions and `-Level WARN` for survivable ones; INFO is the default.

Align label:value lines to the width shown in existing packagers (30 chars for the label, colon, then value). The GUI log pane is monospaced; the alignment makes runs legible.

## Admin elevation

Package requires admin because content copy to network shares can require it in some environments, and CM module cmdlets need it. Stage should avoid admin-only work whenever possible, but a few packagers need elevation to extract vendor archives, inspect installer metadata, or run vendor tooling. Check `Test-IsAdmin` only where the phase actually needs elevation, and exit with a clear message when elevation is required.

## Common mistakes

**Em-dashes in PS files.** PowerShell 5.1 parses em-dash (U+2014) as an invalid token inside executable code. Use plain hyphens. Em-dashes in `.md` docs are fine; not in `.ps1`.

**Mandatory + [string] parameters.** `[Parameter(Mandatory)][string]$Foo` rejects empty strings at runtime with a prompt. Add `[AllowEmptyString()]` or redesign the check.

**`-Encoding UTF8`.** In PS 5.1 that writes a BOM. Use `-Encoding ASCII` for pure-ASCII text, or `-Encoding UTF8NoBOM` if available, or `Set-Content -Encoding UTF8 -NoNewline` plus a manual write.

**Silent vendor-page changes.** Scrape regexes are brittle. Run `-GetLatestVersionOnly` after any packager change and verify the value. Run it periodically against production; if the vendor restructures their page, catch it before One Click does.

**`-StageOnly` side effects in production.** Stage is supposed to be safe and repeatable. Don't write to the registry, don't touch the network share, don't start services. The shared module's helpers stay in the sandbox; your code should too.

## Minimum viable test

Before submitting a new packager:

1. Run `.\Tests\Invoke-PackagerSmoke.ps1`. Confirm all offline parse, metadata, and parameter-contract checks pass.
2. Run `.\Packagers\package-<app>.ps1 -GetLatestVersionOnly`. Confirm the output is the current version in a clean format.
3. Run `.\Packagers\package-<app>.ps1 -StageOnly`. Confirm `<DownloadRoot>\<AppSubfolder>\<Version>\` has: the installer, `install.bat`, `install.ps1`, `uninstall.bat`, `uninstall.ps1`, `stage-manifest.json`.
4. Run `.\Packagers\package-<app>.ps1 -PackageOnly -SiteCode <yourcode> -FileServerPath <yourshare>`. Confirm the network content landed and the MECM app was created with the right detection and deployment type.
5. Deploy to a test VM. Install, detect, uninstall, confirm ARP entry is gone. Re-install, re-detect, re-uninstall. Reboot. Re-detect.
6. Drop the packager into the main grid and Ctrl+Click its vendor URL; confirm the link opens the right page.

Only after all six pass does it belong in the shipping tree.
