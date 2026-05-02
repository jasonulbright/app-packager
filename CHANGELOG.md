# Changelog

## [1.0.0] - 2026-05-02

AppPackager is a MahApps.Metro WPF GUI for the SRL packaging engine.
It discovers per-application packager scripts, queries vendor sources
for the current version, queries MECM for the deployed version,
stages installers + wrappers + detection methods locally, and copies
content to the MECM share + creates the MECM Application + Deployment
Type in one workflow. Ships as a zip + `install.ps1` wrapper; no MSI,
no code signing required.

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
- 7-Zip CLI — optional, required only by packagers that crack
  archived installers (auto-detected at launch via the ARP
  registry)
