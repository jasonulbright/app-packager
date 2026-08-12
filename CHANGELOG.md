# Changelog

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
