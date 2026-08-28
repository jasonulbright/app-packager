# Changelog

## [1.4.0.8] - 2026-08-27

### Added

- **Network split for the remaining M365 SKUs.** The x86 Apps, Project
  (x64/x86), and Visio (x64/x86) packagers gain the same Network split
  shipped for Apps x64 in 1.4.0.7: Online deployment type gated to VPN
  with existence detection, Managed precached fallback, per-SKU
  detection executables and 32/64-bit registry views.

## [1.4.0.7] - 2026-08-27

### Added

- **M365 Apps (x64) network split.** With the Network split selected,
  one application carries both deploy modes: an Online (CDN-direct)
  deployment type gated to VPN-connected machines with existence-only
  detection, and the Managed (version-pinned, precached) deployment
  type as the unconditional on-site fallback. The Managed payload
  stages at the content root, the Online payload under `online\`. The
  M365DeployMode choice is ignored with a logged warning when the
  split is active — the split stages both. The split app is named
  `M365 Apps for Enterprise (x64) (<channel>)`. The other five M365
  SKU packagers keep single-mode behavior for now.

## [1.4.0.6] - 2026-08-27

### Added

- **Firefox language split.** With the Language split selected and
  culture codes entered in OS languages, Firefox stages one MSI per
  language (Mozilla locale resolved from the culture code, falling back
  to the primary subtag: de-DE finds de) plus the en-US payload as the
  unconditional fallback. Each language's deployment type is gated to
  its OS language; the split app is named `Mozilla Firefox (x64)`.
  A language with no Mozilla installer fails the run instead of
  packaging without it.

## [1.4.0.5] - 2026-08-27

### Changed

- **Split apps drop the bitness from their names.** An application whose
  deployment types cover both architectures no longer carries an
  architecture marker: the 7-Zip split creates `7-Zip` (instead of the
  MSI ProductName `7-Zip <version> (x64 edition)`) and the Firefox split
  creates `Mozilla Firefox (en-US)` (the payload is still en-US only,
  so the language stays). Chrome's name had no bitness to drop.
  Single-deployment-type runs keep their existing names, so apps
  packaged without the split keep upgrading in place. A site that
  packaged before enabling the split gets the new name as a new
  application; retire the old-named one after the first split run.

## [1.4.0.4] - 2026-08-27

### Added

- **Chrome and Firefox ARM64 variant splits.** With the Architecture
  split selected, Chrome also stages Google's ARM64 enterprise MSI
  (version-matched against x64, detected on its own ProductCode ARP
  key) and Firefox stages Mozilla's win64-aarch64 exe installer (same
  install path, so the existing file detection and helper.exe uninstall
  carry over). Both network copies are now recursive. Without the
  split, both packagers behave as before.

## [1.4.0.3] - 2026-08-27

### Added

- **7-Zip: first variant-split consumer.** With the Architecture split
  selected, the Stage phase also downloads the ARM64 exe installer
  (version-matched against the staged x64 MSI, or the run fails),
  stages it under `arm64\` with its own wrappers, and writes a
  two-entry `DeploymentTypes` manifest: ARM64 gated to ARM64 CPUs
  (priority 1), x64 gated to x64 (priority 2). The exe variant detects
  on the exe installer's fixed `Uninstall\7-Zip` ARP key. The Package
  phase network copy is now recursive so variant subfolders ship.
  Without the split selected, nothing changes.

## [1.4.0.2] - 2026-08-27

### Added

- **Variant split selection (GUI).** Packagers declare stageable
  variants with a `SupportsVariants:` header tag (comma-separated:
  `Architecture`, `Language`, `Network`, parsed like the other header
  tags). The Deployment Conditions grid gains a Variant split column,
  enabled only for declaring packagers; the selection persists per app
  under `DeploymentConditions.Apps.<packager>.Split` and reaches the
  packager child as `APP_PACKAGER_VARIANTS` JSON on every Package path
  (GUI, One Click, `-BatchMode`). `AppPackagerCommon` 0.0.17 adds
  `Get-RequestedPackagerVariants` for packagers to consume the request;
  no packager implements a split yet — those land with the first
  consumers.

## [1.4.0.1] - 2026-08-27

### Added

- **Multi-deployment-type manifests (common layer).** A stage manifest
  may carry an optional `DeploymentTypes` array; each entry names a
  variant (`NameSuffix` — the deployment type becomes
  `<AppName> - <NameSuffix>`), an optional `ContentSubpath` under the
  network content root, and its own commands, detection, behavior
  overrides, and `Requirements`. Entries are created in manifest order —
  CM assigns priority by creation order and installs the first
  deployment type whose requirements pass — so the most specific variant
  is listed first and the unconditional fallback (an entry with no
  `Requirements`) last. On a version replace the whole set is staged,
  the old set removed, and the staged names promoted. Manifests without
  the array behave exactly as before. On a multi-DT manifest the
  `APP_PACKAGER_REQUIREMENTS` environment JSON is ignored with a logged
  warning: per-entry `Requirements` are authoritative, so an
  unconditional fallback cannot be gated by accident.
  `AppPackagerCommon` 0.0.16 adds `Get-ManifestDeploymentTypeSpecs`.
  No packager or GUI changes; GUI variant selection arrives with a
  later increment.

## [1.4.0.0] - 2026-08-17

### Added

- **Deployment Conditions.** A fifth Options panel attaches CM
  requirement rules to the deployment type each Package run creates,
  per app: architecture (x64 / ARM64 via a WQL global condition on
  `Win32_Processor.Architecture` — numeric, so culture-invariant and
  immune to new-OS-release list churn), OS language (culture codes
  against the site's built-in Operating System Language condition), and
  network context (VPN only / on-site only via a Boolean script global
  condition matching configurable VPN adapter description patterns).
  Requirement rules evaluate on the client at deployment evaluation
  time — no collections, no collection-evaluation load. Global
  conditions are created on first use and matched by name, so renaming
  one in the panel attaches to a condition the site already has.
  Per-app selections persist to `AppPackager.preferences.json`;
  condition names and adapter patterns persist to
  `Packagers/condition-templates.json` with built-in defaults until the
  panel writes it. Applies to Package, One Click Stage-and-Package, and
  `-BatchMode` runs.
- `AppPackagerCommon` 0.0.15: `Get-ConditionTemplates`,
  `Save-ConditionTemplates`, `New-VpnConditionScriptText`,
  `Get-OrCreateGlobalConditionFromTemplate`,
  `Get-DeploymentTypeRequirementSpecs`, and
  `New-DeploymentTypeRequirementRules`.
  `New-MECMApplicationFromManifest` consumes requirement specs from a
  manifest `Requirements` array or the `APP_PACKAGER_REQUIREMENTS`
  environment JSON the GUI sets per app. Requirement resolution runs
  before any application or deployment type is created, so a spec that
  cannot be built fails the run instead of packaging without the
  operator's rules.

## [1.3.0.0] - 2026-08-17

### Added

- **Drop-to-package intake.** Drag an `.msi` or `.exe` installer onto the
  window: the vendored `InstallerAnalysisCommon` module (at
  `Lib\InstallerAnalysisCommon\`) identifies the installer engine, reads
  MSI property tables, and predicts silent switches and the ARP uninstall
  key; an editable preview dialog shows the resulting manifest before
  anything runs. MSI identity is authoritative and detection uses the
  ProductCode ARP key. For every other engine the values are predictions
  and Stage + Package stays disabled until the operator confirms them —
  Stage alone is always available. Multiple dropped installers queue
  through one background pipeline run with the usual progress overlay,
  cancel, and per-item log lines.
- **Save as Packager.** A drop can graduate into the catalog: the app
  writes `Packagers/package-<app>.ps1` from the matching template with
  identity, folder segments, and installer filename filled in. The
  download-source resolution remains the template's TODO — a dropped
  file carries no origin URL — and the generated script satisfies grid
  discovery immediately. Existing packagers are never overwritten.
- `AppPackagerCommon` 0.0.14: `Get-InstallerAnalysis`, `New-AdHocStage`,
  `Invoke-AdHocPackage`, `New-PackagerFromDrop`. Ad-hoc staging emits the
  same schema-v3 `stage-manifest.json` (hashes included) as the packager
  templates, so `New-MECMApplicationFromManifest` consumes it unchanged.

## [1.2.0.1] - 2026-08-16

### Changed

- **Vendored `SuiteCommon` 0.3.2.** Window restore applies the saved
  geometry before maximizing, so un-maximizing returns to the saved size
  instead of the XAML defaults.

## [1.2.0.0] - 2026-08-16

### Changed

- **Shared plumbing now comes from the vendored `SuiteCommon` module**
  (0.3.0, at `Lib\SuiteCommon\`), joining the rest of the tool suite.
  `AppPackagerCommon` drops its own logging trio and CM drive mechanics:
  `Connect-CMSite` stays as a thin wrapper that keeps the app-side
  provider resolution chain (preferences, the AdminUI connect script,
  `APP_PACKAGER_CM_PROVIDER`) and its fail-fast guard, then delegates
  the drive and session work to the shared core with site verification
  off. `APP_PACKAGER_VERBOSE` bridges to `SUITE_VERBOSE` at import, so
  existing packager-script habits keep working. The GUI shell drops its
  title-bar drag block, window-state persistence, button theming, and
  message dialog for the shared implementations; preferences (nested
  schema) and the owner-chrome dialog helper stay app-side by design.
- Behavior gains from the shared layer: title-bar hook state no longer
  leaks when a window closes, a maximized close persists the
  pre-maximize geometry instead of full-screen extents, an off-screen
  saved position is clamped into the nearest monitor instead of
  discarding the saved size, message-dialog icons render as glyphs, and
  Escape closes OK-only dialogs.
- **Pipeline teardown no longer blocks the UI thread.** Canceling an
  in-flight run and closing the window both used a synchronous
  `PowerShell.Stop()` / `Runspace.Close()`, which froze for as long as a
  packager was stuck inside a CM/CIM call; teardown now stops
  asynchronously into a reaped graveyard and closes the runspace async.

## [1.1.0.7] - 2026-08-14

### Added

- **Content Layout toggle: nested or flat share folders.** MECM
  Preferences gains a Content Layout selector: Nested
  (`Applications\Vendor\App\Version`, the default and unchanged
  behavior) or Flat (`Applications\Vendor-App-Version`, one folder per
  package, for org conventions that mandate it). New exported
  `Get-NetworkContentPath` is the single construction point for both
  shapes; all 92 packagers plus the PSADT template now call it (their
  new `-ContentLayout` parameter defaults to Nested, so CLI behavior is
  identical unless asked). The GUI passes the preference through
  Package, One Click, and -BatchMode runs, and the two path-
  reconstructing fallbacks (package-integrity verification and the
  .intunewin post-step) are layout-aware. Existing content is never
  moved; the choice applies to future Package runs. Deliberately a
  binary toggle rather than a free-form pattern: two enumerable layouts
  keep every path-touching feature testable against both. Idea credit:
  stephannn (PR #2), reshaped.

## [1.1.0.6] - 2026-08-14

### Added

- **PSADT support finished.** The `package-psadt.ps1.template` skeleton
  (previously TODO-throws in every phase) is now a functional packager for
  the wrap-a-wrap case: point it at an existing per-app PSADT folder, fill
  the identity and detection markers, and it stages the full toolkit tree
  as versioned content with SHA256 hashes over every file — subfolders
  included, so package integrity verification covers the toolkit with no
  `-AllowExtra` weakening. New exported `Test-PsadtLayout` detects v3
  (`Deploy-Application.exe`) vs v4 (`Invoke-AppDeployToolkit.exe`, with a
  `powershell.exe -File` fallback when the launcher exe is absent) and
  builds the deployment type command lines, validated against the
  PSAppDeployToolkit documentation. The stage manifest gains optional
  `InstallCommandLine` / `UninstallCommandLine` fields that
  `New-MECMApplicationFromManifest` honors over the generated .bat
  wrappers — a general mechanism, PSADT is just its first consumer.
  `DeployMode` stays with the toolkit by default (interactive
  close-app/defer when a user session exists); `-DeployMode Silent`
  forces quiet. Toolkit versions are pinned per app in the source folder.
  Prompted by stephannn's PR #2; implemented per-app instead of global
  injection so integrity verification and version pinning survive.

## [1.1.0.5] - 2026-08-14

### Changed

- **Non-admin packaging: local-administrator gates removed from every
  packager.** The `Test-IsAdmin` exit-gate dated from the temp-install
  ARP-derivation era; no shipped packager installs anything during Stage
  anymore (`Find-UninstallEntry` has zero callers, and every msiexec
  reference lives in generated endpoint wrapper content). Stage reads
  installer metadata via COM and writes only user-writable paths; Package
  needs network-share ACLs and CM RBAC, neither of which is local
  elevation. Removed 186 gate blocks and the "Local administrator"
  requirement line across 94 packagers/templates, so packaging runs from
  a standard user account — no admin accounts needed in CM RBAC roles or
  content-share ACLs. `Test-IsAdmin` stays exported for future packagers
  whose Stage provably needs elevation; `Samples/AUTHORING.md` documents
  the new house rule. Idea credit: stephannn (PR #2).

## [1.1.0.4] - 2026-08-14

### Added

- **Invoke-WebRequest fallback for downloads.** `Invoke-DownloadWithRetry`
  keeps in-box curl.exe primary — the Schannel build trusts the Windows
  certificate store and negotiates modern TLS independent of per-machine
  .NET registry state — and falls back to `Invoke-WebRequest` when curl
  fails, covering networks that force a WinINET-configured proxy curl
  cannot see. The fallback sends default credentials to the system proxy
  (Kerberos/NTLM auth), uses `-UseBasicParsing`, suppresses the 5.1
  progress bar that throttles large downloads, and deletes partial files
  between methods and attempts so torn content can never pass integrity
  verification. Scraping/URL-resolution calls are unchanged. Idea credit:
  stephannn (PR #2).

## [1.1.0.3] - 2026-08-14

### Changed

- **dotnet8 / dotnet10both detection accepts the successor patch.** In-place
  runtime upgrades replace `dotnet\host\fxr\<version>`, flipping the prior
  month's app to not-installed and making its still-active deployment
  reinstall over the new runtime — previously managed with supersedence.
  Detection is now `(x86-N AND x64-N) OR (x86-N+1 AND x64-N+1)`, where N+1
  is the packaged version with its patch component incremented
  (`Get-NextPatchVersion`, new exported helper; falls back to
  single-version detection with a warning for non-numeric components such
  as previews). The manifest schema gains `Detection.GroupSizes` — exactly
  two contiguous clause runs; `New-MECMApplicationFromManifest` passes the
  second run to `-GroupDetectionClauses` with an `OR` connector on its
  first clause, and the cmdlet's left-associative expression build
  parenthesizes the first run, yielding the grouped OR without
  supersedence management between consecutive monthly packages.

### Removed

- **`package-dotnet10x64.ps1`.** Obsolete; the dual-bitness
  `package-dotnet10both.ps1` is the only .NET 10 packager.

## [1.1.0.2] - 2026-08-14

### Fixed

- **Version-less applications never picked up new versions.** For packagers
  whose CMName omits the version (by design), the Package phase found the
  existing application + deployment type and returned success without
  touching either — the new version's content was staged and copied but
  never reached MECM. `New-MECMApplicationFromManifest` now compares the
  manifest `SoftwareVersion` against the existing application's: unchanged
  versions remain an idempotent no-op; a changed version replaces the
  deployment type and updates the application's SoftwareVersion (and
  Description when a comment is passed). The new deployment type is created
  under a staging name, the old one is removed, then the new one is renamed
  to the canonical name — ordered that way because a deployed application
  refuses to remove its last deployment type. Renaming uses
  `Set-CMDeploymentType -NewDeploymentTypeName` (validated against vendor
  docs; the parameter is not `-NewName`).

- **specexec packager verifies deployment types server-side.** The run
  reported success as long as no cmdlet threw; a silently incomplete
  application (created by a pre-1.0.0.11 partial run) passed unnoticed.
  The Package phase now logs how many deployment types the existing
  application has when resuming, and after the create loop queries the
  site and prints OK/MISSING per expected deployment type, throwing if
  any of the 6 are absent.

## [1.1.0.1] - 2026-08-14

### Fixed

- **Grid columns no longer collapse when Debug Columns is toggled.** The
  star-sized Application column had no minimum width, so showing the four
  fixed-width debug columns crushed it to a sliver — and the crushed
  width did not recover when they were hidden again. The column now keeps
  at least 160 px, extra width scrolls horizontally, and hiding the debug
  columns restores the layout.

### Added

- **Grid filter box.** A filter above the application grid narrows rows
  by application, vendor, status, or CM name as you type. Bulk selection
  (the checkbox-column header cycle and its updates-only step) acts on
  the visible rows and always clears hidden rows first, so a filtered
  "select all" can never queue hidden apps into a Stage or Package run.

## [1.1.0.0] - 2026-08-13

### Additions

- **Intune Win32 content prep during Package.** MECM Preferences gains a
  "Create .intunewin during Package" option plus a Content Prep
  detected-tool row. When enabled, a successful Package run also produces
  `<app>-<version>.intunewin` from the staged content (setup reference:
  `install.bat`) and stores it beside the network content version folder,
  with a copy beside the local staged version folder. The artifact never
  lands inside a version folder — stage hash verification treats added
  files as integrity failures. Prep failures are surfaced in the log and
  never fail the package run; the MECM application is already created
  when the post-step executes. Settings persist under `Intune` and
  `DetectedTools.IntuneWinAppUtil` in `AppPackager.preferences.json`;
  prefs files without the new keys load with the feature off. Zero
  per-packager changes — the post-step hangs off the shared Package path,
  so every packager gains the capability at once.
- **IntuneWinAppUtil.exe detected tool with download-on-first-use.**
  Detection checks the stored preferences path, the tool cache under
  `%LOCALAPPDATA%\AppPackager\Tools`, and PATH, once per launch. The
  Download button in MECM Preferences fetches the Microsoft Win32 Content
  Prep Tool from Microsoft's repository and keeps the file only after its
  Authenticode signature verifies as Valid and Microsoft-signed; a failed
  verification deletes the download and reports the reason. The
  "Create .intunewin during Package" option stays locked until the tool
  is present. The tool is never redistributed with AppPackager.
- **`AppPackagerCommon` 0.0.12 exports `Install-IntuneWinAppUtil` and
  `New-IntuneWinPackage`** so packager scripts and ad-hoc callers can
  produce `.intunewin` files directly. `New-IntuneWinPackage` refuses an
  output folder equal to the content folder, enforces a bounded runtime
  on the prep tool, and returns the artifact path, size, and SHA-256.

## [1.0.0.12] - 2026-08-13

### Additions

- **Test-collection deployment after content distribution.** MECM
  Preferences gains three controls under the Auto-distribute group:
  "Deploy to test collection after distribution", "Test collection"
  name, and "Create collection if it does not exist". The controls
  unlock only when Auto-distribute is enabled and a DP Group is set —
  the gating lives in the GUI, not the runtime. When enabled, the
  Package phase follows `Start-CMContentDistribution` with
  `New-CMApplicationDeployment` (Install / Available /
  available immediately / default options) to the named collection.
  A missing collection is created as an empty direct-membership device
  collection limited to All Systems when the create option is checked,
  otherwise the deployment is skipped with a warning. Existing
  deployments are treated as success so re-packaging stays idempotent.
  Settings persist under `ContentDistribution` in
  `AppPackager.preferences.json`; prefs files without the new keys load
  with the feature off.

## [1.0.0.11] - 2026-08-12

### Fixes

- **`specexec-mitigations` failed at the deployment-type existence check.**
  The packager called `Test-MECMApplicationHasDeploymentType`, which is
  internal to `AppPackagerCommon` (absent from `FunctionsToExport`), so the
  call failed with "not recognized" on every run. Other packagers were
  unaffected because that helper only runs inside
  `New-MECMApplicationFromManifest`, in module scope. The check now uses
  `Get-CMDeploymentType -ApplicationName -DeploymentTypeName` directly.

## [1.0.0.10] - 2026-08-12

### Additions

- **`package-specexec-mitigations.ps1` — first multi-deployment-type
  packager.** One application, six Script deployment types covering the
  speculative-execution CVE registry mitigations: {Intel HT-on, Intel
  HT-off, AMD} x {standard, Hyper-V host}. Deployment types exist per
  distinct registry payload only — OS class never changes the values, so
  there is no workstation/server split, and the AMD override is
  HT-independent. Requirement rules route each device to exactly one
  deployment type via three global conditions (reused by name when they
  already exist, created otherwise): CPU vendor (WQL,
  `Win32_Processor.Manufacturer`), Hyper-Threading state (Boolean script;
  WQL cannot compare two properties of one instance), and Hyper-V role
  (`Services\vmms` key existence). Detection is per-deployment-type
  `FeatureSettingsOverride` / `FeatureSettingsOverrideMask` DWORD
  comparison, plus `MinVmVersionForCpuBasedMitigations` on the Hyper-V
  variants. Content is self-generated (no vendor download);
  `-GetLatestVersionOnly` reports the pinned `-ContentVersion`. Install
  exits 3010 with `RebootBehavior BasedOnExitCode`. Every CM cmdlet
  parameter was validated against Microsoft Learn documentation — notably
  `New-CMDetectionClauseRegistryKeyValue` accepts `Integer` (not `Int64`)
  and its `-Is64Bit` switch means the *32-bit* registry view, the inverse
  of the global-condition cmdlets' `-Is64Bit` boolean.

## [1.0.0.9] - 2026-08-01

### Fixes

- **Multi-installer wrappers aborted on success codes.** The generated
  `install.ps1` for `msvcruntimes`, `dotnet8`, and `dotnet10both` guarded
  the first installer with `if ($proc1.ExitCode -ne 0) { exit ... }`.
  Windows installers return **3010** for "succeeded, reboot required" and
  **1641** for "succeeded, reboot initiated", so a successful first
  install exited the script before the second architecture was installed,
  and MECM recorded the deployment as failed. Compound detection requires
  both architectures, so the application could never report installed.
  The wrappers now treat `0`, `3010`, and `1641` as success, run both
  installers, and surface a pending-reboot code from either one.

  Impact was uneven. `msvcruntimes` installs x86 first, and x86
  redistributables rarely request a reboot, which is why this stayed
  hidden. `dotnet8` and `dotnet10both` install **x64 first**, so a reboot
  code there skipped the x86 runtime — the one that x86 consumers of
  .NET 8 such as Citrix Workspace, SailPoint, and VMware Tools depend on.

- **`winrar` skipped its license key on a reboot code.** Same guard, same
  cause: a 3010 from the installer bypassed the `rarreg.key` copy. The
  wrapper now continues on reboot codes and returns the installer's real
  exit code instead of a hardcoded `0`.

## [1.0.0.8] - 2026-06-13

### Additions

- **`-VerboseLog` on every packager.** The switch (and the
  `APP_PACKAGER_VERBOSE=1` environment variable equivalent) previously
  existed only on `package-adobereader.ps1`; all packagers, sample
  packagers, and templates now accept it and pass it through to
  `Initialize-Logging`. Every packager's top-level catch now records
  full failure detail via `Write-LogErrorRecord` — exception chain,
  `FullyQualifiedErrorId`, failing `file:line`, failing statement, and
  script stack trace — before the `SCRIPT FAILED` summary line, so any
  packager failure identifies its exact call site in the log.

## [1.0.0.7] - 2026-06-12

### Changes

- **Optimization: removed the per-entry `Get-CMSite` connection probe
  added in v1.0.0.6.** The probe cost a provider round trip on every
  site-drive entry — twice per package run and once per application
  during a batch Check MECM (~90 round trips across the full packager
  set). A successful `Set-Location` onto the site drive is trusted
  again; if entering an existing drive fails, the drive is still torn
  down and rebuilt from the configured provider machine at no
  happy-path cost.
- **Clarification on v1.0.0.6:** the stated rationale for the probe was
  not accurate. The behavior it targeted traces to broken RBAC role
  assignments, not to site-drive connection lifetime — broken RBAC can
  make CM cmdlets fail or return null while the console works normally.
  With correct RBAC, site-drive connections behave as they always did.
  If cmdlets half-work while the console is fine, review RBAC role
  assignments before suspecting the connection.

## [1.0.0.6] - 2026-06-12

### Fixes

- **CMSite drive connections survive ConfigMgr 2509's stale-drive
  behavior.** As of 2509 the site drive's provider connection does not
  survive the session leaving the drive (`Set-Location C:`), and a drive
  auto-mounted by the module's `OnImport` hook may never have had a live
  connection at all: re-entering the drive either fails or succeeds with a
  dead connection where every CM cmdlet throws
  `Key cannot be null. Parameter name: key`. `Connect-CMSite` and the
  GUI's inline Check MECM connect now probe the drive with a cheap
  `Get-CMSite` call after every entry; if the probe fails they leave the
  drive, `Remove-PSDrive` it, recreate it from the configured provider
  machine (falling back to the stale drive's own Root when no provider is
  configured), re-enter, and re-probe — the same sequence as rerunning the
  AdminUI connect script. CM cmdlets continue to run only from the site
  drive and filesystem work only from `C:`; every `C:` detour now goes
  through the reconnect logic on the way back.
  `Get-MecmCurrentVersionByCMName` also restores the caller's original
  location instead of leaving the shell parked on the site drive, where the
  next `C:` operation would silently kill the connection.

## [1.0.0.5] - 2026-06-12

### Additions

- **Verbose failure diagnostics for packager scripts.** `Write-Log` gains a
  `DEBUG` level (always written to the structured log file; echoed to the
  console only when verbose logging is on) and the module exports
  `Write-LogErrorRecord`, which logs the full exception chain,
  `FullyQualifiedErrorId`, the failing `file:line`, the failing statement,
  and the script stack trace from any catch block. Enable verbose mode with
  `Initialize-Logging -VerboseLogging`, the new `-VerboseLog` switch on
  `package-adobereader.ps1`, or `APP_PACKAGER_VERBOSE=1` (inherited by GUI
  child processes). `New-MECMApplicationFromManifest` now tracks which step
  is in flight (`Connect-CMSite`, duplicate check, `New-CMApplication`,
  detection clause creation, `Add-CMScriptDeploymentType`, revision-history
  cleanup) and names it on failure, so opaque ConfigMgr cmdlet errors such
  as `Key cannot be null. Parameter name: key` finally identify their call
  site in the log for every packager.
- **Provider site-code validation in `Connect-CMSite` (verbose mode).**
  After connecting, the module queries `root\sms:SMS_ProviderLocation` on
  the drive's provider machine and warns when the drive name does not match
  a site code the provider actually serves — the canonical cause of CM
  cmdlets failing with `Key cannot be null. Parameter name: key` after an
  apparently successful connect.

### Fixes

- **`Get-MecmCurrentVersionByCMName` honors the Provider Machine
  preference.** The function accepted `-ProviderMachineName` but never used
  it: when the `${SiteCode}:` drive was not mounted (console MRU empty
  because the AdminUI never connected on that workstation), Check MECM
  failed with `Failed to connect to CM site PSDrive`. It now falls back to
  `New-PSDrive -Root <ProviderMachineName>` inline at script scope before
  giving up, and the giving-up message says exactly which preference to
  set. Manifest AppName is also validated as non-empty before
  `Get-CMApplication`/`New-CMApplication` run.

## [1.0.0.4] - 2026-05-27

### Additions

- **Two NVIDIA Graphics Driver packagers (FR #1).** Adds
  `package-nvidia-geforce.ps1` (GeForce Game Ready DCH x64, covers
  current Maxwell+ consumer GTX/RTX cards) and
  `package-nvidia-rtx-enterprise.ps1` (Quadro Certified DCH x64, covers
  current NVIDIA RTX PRO / RTX A-series workstation cards). Both query
  NVIDIA's `AjaxDriverService.php` JSON endpoint with pinned `psid`/`pfid`
  per packager. Pinning answers the lookup-form combo-box ambiguity by
  treating the flagship pfid as a stable "latest driver for current
  family" key, since the DCH installer is unified across the whole
  family. Detection is a single ARP `RegistryKeyValue` on the constant
  NVIDIA Display.Driver uninstall GUID
  (`{B2FE1952-0186-46C3-BAEC-A80AA35AC5B8}_Display.Driver` →
  `DisplayVersion`), so both MECM apps coexist without colliding. Silent
  install: `setup.exe -s -noreboot -clean`. Silent uninstall:
  `setup.exe -uninstall -s -noreboot`.
- **README packager count: 89 → 91.** Table extended with the two new
  NVIDIA entries.

## [1.0.0.3] - 2026-05-13

### Fixes

- **MECM site connection restored for `Get-MecmCurrentVersionByCMName`
  (Check MECM button).** v1.0.0.2 routed the in-script Check MECM call
  through the module-scope `Connect-CMSite`. Module functions run in
  their own isolated session state, so the caller's `${SiteCode}:`
  PSDrive (mounted by an AdminUI connect script in the launching shell)
  was not visible and the function fell through to the explicit-provider
  branch, throwing `'<SiteCode>' Configuration Manager PSDrive is not available
  and no provider machine name is configured` even though the drive was
  usable in the same shell. Restored the inline `Import-Module
  ConfigurationManager` + `Set-Location ${SiteCode}:` pattern at script
  scope. Workstations that pre-mount the drive go straight to
  `Set-Location`; workstations that have not pre-mounted import the
  module so its `OnImport` hook mounts the drive from `HKCU` MRU, then
  `Set-Location` succeeds.

## [1.0.0.2] - 2026-05-08

### Fixes

- **ConfigMgr site connections now match the AdminUI connect prompt.**
  `Connect-CMSite` no longer fails immediately when the `${SiteCode}:`
  drive is absent; it creates the missing `CMSite` PSDrive with
  `New-PSDrive -Root <ProviderMachineName>` and then enters the drive.
  MECM Preferences now stores the provider machine (paste the
  `$ProviderMachineName` value from the AdminUI connect script into
  Options → MECM Preferences → Provider Machine), and package child
  processes receive it via `APP_PACKAGER_CM_PROVIDER`. The deliberate
  `MCM:`/`C:` choreography inside `New-MECMApplicationFromManifest` is
  preserved; child packager launches use the packager folder as their
  working directory and `Get-MecmCurrentVersionByCMName` restores the
  caller's location in `finally`.
- **`Resolve-ConfigurationManagerModulePath` adds preferences-detected
  and known install paths as fallbacks when `SMS_ADMIN_UI_PATH` is
  absent.**
- **Packager history timestamps stay ISO formatted after JSON round-trip.**
  `Read-PackagerHistory` normalizes `ConvertFrom-Json` datetime values
  back to `yyyy-MM-ddTHH:mm:ssZ`, avoiding culture-formatted strings in
  callers.

## [1.0.0.1] - 2026-05-06

### Fixes

- **Package step now binds with empty `-Comment`.** PS 5.1's
  `[Parameter(Mandatory)][string[]]` rejects arrays containing empty-string
  elements with `Cannot bind argument to parameter 'Arguments' because it
  is an empty string.` The Package argsBase always carried `-Comment $Comment`;
  when the GUI user left Comment blank, `$Comment=''` broke the bind. Stage's
  argsBase had no empty literals so it was never affected. `[AllowEmptyString()]`
  on `Set-ProcessStartInfoArgumentList` permits the element through.
- **`Connect-CMSite` survives a missing `SMS_ADMIN_UI_PATH`.** `Join-Path`
  threw on a null env var before the `Import-Module ConfigurationManager`
  fallback could run. Guarded; falls back cleanly when the env var is absent.
- **Streaming child processes time out on idle.** `Invoke-ProcessWithStreaming`
  only enforced `WaitForExit(15s)` after the stdout read loop exited, so a
  child that hung without printing wedged the GUI forever. Added a
  configurable idle timeout (default 30 minutes) that kills a silent child
  and surfaces the kill in stdout.
- **`Save-Preferences` no longer silently drops failures.** The empty `catch`
  around the `packager-preferences.json` write hid every failure; the GUI
  reported success while packagers kept reading stale Company / M365 / SSMS
  values. Now surfaces via `Write-Warning`.
- **`Get-PackagerFolderInfo` reads enough of the file.** `-TotalCount 120`
  cut off packagers where `$AppFolder` / `$BaseDownloadRoot` live past line
  120 (e.g. `package-teamviewerhost.ps1`). Streams until all three vars are
  found, then breaks early.

### Additions

- **`package-dotnet10both.ps1`** — dual-arch .NET 10 Desktop Runtime packager.
  Mirrors `package-dotnet8.ps1` with channel-version 10.0; ships x86 + x64 in
  a single MECM application with compound File detection on `hostfxr.dll`.

## [1.0.0] - 2026-05-02

AppPackager is a MahApps.Metro WPF GUI for the SRL packaging engine.
It discovers per-application packager scripts, queries vendor sources
for the current version, queries MECM for the deployed version,
stages installers + wrappers + detection methods locally, and copies
content to the MECM share + creates the MECM Application + Deployment
Type in one workflow. Extract the zip and run `start-apppackager.ps1`.

### Features

- **Sidebar workflow** — One Click, Check Latest, Check MECM, Stage
  Packages, Package Apps, plus an Options modal. Theme toggle
  bottom-docked on the sidebar.
- **One Click** — iterates the apps you've marked as tracked in One
  Click Settings and runs Check Latest → Stage → Package per the
  chosen action. Cadence-gated so Report-only runs throttle; Stage
  and Stage-and-Package always run. Pre-Stage MECM pre-flight skips
  any tracked app whose version is already in MECM.
- **Background pipeline runspace** — multi-app loops run on a
  background STA runspace with an animated progress overlay so the
  window stays responsive instead of freezing during long downloads
  / extracts / MECM round-trips. Pause / Cancel after current app
  available mid-run.
- **Application grid** — every discovered packager rendered as a
  row: vendor, current MECM version, latest vendor version, status,
  comment field. Persistent history (Last Checked, Latest Version)
  stored in `%LOCALAPPDATA%\AppPackager\app-history.json` so values
  survive across sessions.
- **Options modal** — Discord/VS Code-style left-nav + right pane:
  - **MECM Preferences** — site code, file share root, download
    root, estimated/maximum runtime, Auto-distribute-to-DP toggle +
    DP Group Name, plus read-only detected-tools status (ConfigMgr
    Console + 7-Zip CLI).
  - **Packager Preferences** — M365 ODT settings (channel, deploy
    mode, ExcludeApps), SSMS silent install options, TeamViewer Host
    config, Citrix Workspace App switches. Inline preview buttons
    show the assembled CWA command line and the generated ODT
    `install.xml`.
  - **One Click Settings** — pick which packagers the tracked set
    includes, choose action (Report-only / Stage / Stage and
    Package), toggle Force on launch, set per-app cadence overrides.
  - **Product Filter** — show / hide individual packager scripts in
    the main grid by vendor (checkbox TreeView).
- **Search dialog** — themed search/picker modal for application,
  package, task sequence, software-update group, and collection
  names; replaces interactive `Read-Host` prompts inside the
  packagers.
- **Themed Message dialog** — every confirm / message routes through
  a brand-themed `Show-ThemedMessage` helper; no raw system
  MessageBoxes.
- **Preview dialog** — read-only inspector for generated install.xml
  / CWA command lines / packaged-content manifests.
- **Title-bar drag fallback** — native `WM_NCHITTEST` hook + managed
  `DragMove` for the main window and every modal dialog so the
  title bar drags reliably under any host.
- **MahApps Dark.Steel / Light.Blue themes** with live swap.
- **Window state persistence** — size, position, theme, debug-column
  state all restored across launches.

### Stage / Package safety

- **Schema-v3 stage manifests** — every staged payload + generated
  wrapper has `RelativePath`, `SHA256`, and `Size` recorded.
- **Post-Stage verification** — Stage fails closed if the on-disk
  stage folder does not match the manifest hash list.
- **Post-Package verification** — Package fails closed if the copied
  network content does not match the staged manifest before any
  MECM application creation.
- **Pre-1.0.3 soft landing** — older schema v2 manifests without
  `FileHashes` still read with a WARN and skip byte-level
  verification.
- **MECM existing-app validation** — Packaging fails closed when an
  existing MECM application is missing the expected deployment type,
  rather than treating a partial prior run as success.
- **Operation summaries** — Check Latest, Stage, Package, and One
  Click maintain operation counts and emit operation-specific
  summary labels.

### Vendor source coverage

Built-in packagers cover (alphabetical, abridged): 7-Zip, Adobe
Acrobat Reader, Audacity, Bitwarden Desktop, DBeaver Community,
Draw.io, Everything, Firefox, GIMP, Git for Windows, Google Chrome,
Inkscape, Microsoft Edge, Microsoft Teams (new client), Mozilla
Firefox, Notepad++, Office 365 (Apps / Project / Visio, x64+x86, all
six SKUs), PostgreSQL, Postman (User), Power BI Desktop, PuTTY,
SQL Server Management Studio 22, TeamViewer Host, Visual Studio Code
(User + System), VLC, WinSCP, Wireshark.

### Stack

- PowerShell 5.1 + .NET Framework 4.7.2+
- WPF + MahApps.Metro (vendored DLLs in `Lib\`)
- ConfigurationManager PowerShell module (provided by the MECM
  Console install) — required for Check MECM, Package Apps, and
  One Click with Stage-and-Package
- 7-Zip CLI — optional, required only by packagers that extract
  archived installers (auto-detected at launch via the ARP
  registry)
