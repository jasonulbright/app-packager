# Changelog

## [1.5.1.9] - 2026-09-05

### Fixed

- Install Spectra PDF unattended with its required acceptance switch.

## [1.5.1.8] - 2026-09-04

### Fixed

- Uninstall KDiff3 with its silent switch alone; the in-place switch removed nothing.

## [1.5.1.7] - 2026-09-04

### Added

- Choose Install for System or User per app for installers with a mode switch.

### Changed

- Show the architecture, language and network cells as visible controls in Deployment Conditions.
- Align the detected-tool rows with the input rows in MECM Preferences.
- List the default VPN adapter patterns alphabetically.

### Fixed

- Pass the variant split selection to the Stage phase.
- Enable the Variant split cell only for packagers that declare variants.

## [1.5.1.6] - 2026-09-04

### Fixed

- Detect Gpg4win in the 64-bit registry view.
- Detect KDiff3 in the 32-bit registry view.

## [1.5.1.5] - 2026-09-04

### Changed

- Vendor installer-analysis 1.3.3.0 so drops resolve more NSIS detection keys and registry views.

### Fixed

- Install RStudio for all users and detect it in the 64-bit registry view.

## [1.5.1.4] - 2026-09-04

### Added

- Offer an Install for toggle in the drop preview for installers with a mode switch.
- Swap install arguments, uninstall command, folder, detection key and context together when the mode changes.

## [1.5.1.3] - 2026-09-04

### Added

- Show GitHub API authentication status and quota in Options.

### Changed

- Use a signed-in GitHub CLI token for GitHub API calls.

## [1.5.1.2] - 2026-09-04

### Changed

- Read Wireshark's uninstall key from the installer script instead of installing it during staging.
- Authenticate GitHub API calls with GITHUB_TOKEN or GH_TOKEN when set.

### Fixed

- Stage Slack as an MSIX provisioned for all users.
- Read the Defraggler version from the builds page and the installer resource.
- Accept ShareX's per-architecture setup asset.
- Re-download a cached Android Studio installer whose checksum fails.
- Retry the Windows ADK and WinPE add-on layout when another installer holds the mutex.

## [1.5.1.1] - 2026-09-04

### Added

- Add Apache NetBeans (catalog at 285).

### Changed

- Read Inno Setup drops from the compiled header: AppId key, version, folder and context.
- Vendor SuiteCommon 0.4.3 so PowerShell 7 module paths no longer break child runspaces.

### Fixed

- Stop classifying Inno Setup 6 installers as per-user.

## [1.5.1.0] - 2026-09-04

### Added

- Read NSIS drops from the compiled script: folder, uninstaller, detection key, hive and view.
- Stage per-user installers as user-context deployments with HKCU detection.
- Expand environment variables in uninstall wrappers at run time.
- Show the install context and detection source in the drop preview.

### Changed

- Switch Audacity 4.0 to its MSI installer.

### Fixed

- Show the Icon Pack and Content Prep buttons at every window width.
- Stop the sidebar comment box clipping at the default window height.
- Keep the Application column readable in Deployment Conditions.
- Verify the checksum when installing from a local zip.
- Keep PowerShell 7 module paths out of child runspaces.

## [1.5.0.15] - 2026-09-02

### Added

- Add an Add Installer sidebar button for the drop-to-package intake.

## [1.5.0.14] - 2026-09-02

### Added

- Install from an already downloaded release zip with -ZipPath.

### Changed

- Install from the release zip alone, without downloading a script first.

## [1.5.0.13] - 2026-09-02

### Fixed

- Ship every PowerShell file as pure ASCII.
- Bootstrap through curl.exe in the README one-liner.

## [1.5.0.12] - 2026-09-02

### Fixed

- Publish the bootstrap script as a release asset and download through curl.exe first.

## [1.5.0.11] - 2026-09-02

### Added

- Install the icon pack from a local or UNC zip in Options.

## [1.5.0.10] - 2026-09-02

### Added

- Choose Skip, Overwrite or Fail when a same-version application already exists.
- Prompt during Package and One Click conflicts, with apply to all remaining.

### Changed

- Re-apply Publisher when updating an existing application.

### Fixed

- Tolerate already-removed revisions during revision-history cleanup.

## [1.5.0.9] - 2026-09-02

### Fixed

- Apply the application icon on first and on unchanged-version package runs.

## [1.5.0.8] - 2026-09-02

### Added

- Publish the full icon pack with an icon for every packager.

### Fixed

- Convert the largest frame of multi-frame MSI icons.
- Read the Inkscape MSI link from the release page.
- Use external icons for Thonny and TurboVNC.

## [1.5.0.7] - 2026-09-02

### Added

- Load external icons from a downloadable icon pack, with an Options row to install it.
- Mark an icon source on every packager.

### Fixed

- Retry icon extraction when a freshly written file is still locked.

## [1.5.0.6] - 2026-09-02

### Fixed

- Refresh changed files on the network share instead of skipping existing names.
- Copy nested stage folders on the drop-to-package path.

## [1.5.0.5] - 2026-09-02

### Fixed

- Keep application icons under the 256 KB Configuration Manager limit.

## [1.5.0.4] - 2026-09-02

### Added

- Extract application icons from installers and apply them in MECM and Intune.
- Add a bootstrap installer that downloads, verifies and extracts a release.
- Preserve preferences, window state and logs across updates.
- Check for updates at launch and offer Update now in the sidebar.
- Add an About panel with version, license and release links.

## [1.5.0.3] - 2026-09-02

### Changed

- Disable DBeaver AI through its environment variable instead of editing its configuration file.

## [1.5.0.2] - 2026-09-02

### Fixed

- Report blocked files after a zip extract and offer to unblock them.
- Stop a failed module import from surfacing as an unrelated error.

## [1.5.0.1] - 2026-09-02

### Added

- Add DBeaver install scope and Disable AI options.
- Install DBeaver per user with matching uninstall and detection.

## [1.5.0.0] - 2026-09-01

### Added

- Open a first-run setup wizard for the deployment target and its settings.
- Adapt the sidebar to the deployment target.

### Fixed

- Start Intune-only runs without a ConfigMgr console, site code or share.

## [1.4.0.24] - 2026-09-01

### Added

- Add Yubico PIV Tool, YubiKey Manager CLI, Zeal, Zotero and Zulip Desktop.
- Add Azul Zulu JDK and JRE 8, 11 and 17 (catalog at 284).

## [1.4.0.23] - 2026-09-01

### Added

- Add Synology Drive Client, SmarTTY, Tabular Editor 2, Tailscale, TeamSpeak 3 and TeraCopy.
- Add Bulk Rename Utility, Thonny, TightVNC, TortoiseHg, TurboVNC, Typora and UltiMaker Cura.
- Add UltraVNC, Unity Hub, UrBackup Client, Vagrant, VeraCrypt and Omnissa Horizon Client.
- Add VSCodium, WebStorm, WireGuard, XnView MP and Yubico Authenticator (catalog at 273).

## [1.4.0.22] - 2026-08-31

### Added

- Add PeaZip, PicPick, Pidgin, Proton VPN, PSPad, QGIS, QGIS LTR and Rainmeter.
- Add Rancher Desktop, Raspberry Pi Imager, RenderDoc, Rocket.Chat, RustDesk and RVTools.
- Add Salesforce CLI, ScreenToGif, SharePoint Online Management Shell, Shotcut and Simplenote.
- Add SMath Studio, Softerra LDAP Browser, Stellarium and SyncBackFree (catalog at 249).

## [1.4.0.21] - 2026-08-31

### Added

- Add NetLogo, NETworkManager, Nextcloud, NoMachine, NVDA, Obsidian, ocenaudio and Oh My Posh.
- Add OpenShot, OpenVPN, OpenWebStart, MySQL Connector/NET, OrcaSlicer, ownCloud and Pandoc.
- Add Parallels Client, Password Safe, Path Copy Copy, PDF Studio Viewer and PDF24 Creator.
- Add PDFCreator, PDFgear and PDFsam Basic (catalog at 226).

### Fixed

- Load the compression assembly in MSIX uninstall wrappers.

## [1.4.0.20] - 2026-08-31

### Added

- Add Calibre, Kreya, Krita, Liberica JDK 21, MariaDB Server and Mattermost Desktop.
- Add Azure CLI, Azure PowerShell, Azure Storage Explorer and SQL Server 2022 Express.
- Add Windows Admin Center, Windows ADK, Windows PE add-on, MongoDB Compass and Firefox ESR.
- Add Intune Debug Toolkit, MuseScore Studio 4, Nagstamon, NAPS2 and NetBird (catalog at 203).

## [1.4.0.19] - 2026-08-31

### Added

- Add Chrome Remote Desktop Host, Google Credential Provider, Google Drive, Go, Graphviz and grepWin.
- Add gsudo, HandBrake, HashTools, HeidiSQL, HWMonitor, IAP Desktop and IBM Aspera Connect.
- Add IBM Semeru JDK and JRE 8, 11 and 17, ImageGlass and IrfanView.
- Add Joplin, KDiff3, KeePassXC and KeyStore Explorer (catalog at 183).

## [1.4.0.18] - 2026-08-31

### Added

- Add Clockify, CloudCompare, Cloudflare WARP, CMake, CodeMeter Runtime Kit and Colour Contrast Analyser.
- Add Cryptomator, Cyberduck, DataGrip, DAX Studio, DB Browser for SQLite, DbVisualizer and Defraggler.
- Add Dell Command Update, Devolutions Remote Desktop Manager, DisplayLink Graphics, dnGREP and Draftable Desktop.
- Add Duo Desktop, FreeCAD, GeoGebra Classic, Gephi, GitHub CLI, GoLand and Gpg4win (catalog at 158).

### Changed

- Regenerate the README application table from the packager headers.

## [1.4.0.17] - 2026-08-31

### Fixed

- Point the Spectra PDF packager at the current repository.

## [1.4.0.16] - 2026-08-31

### Added

- Add Agent Ransack, AIMP, AWS CLI v2, AWS Tools for Windows and AWS VPN Client.
- Add Amazon DCV Client, Amazon Redshift ODBC driver, Amazon WorkSpaces, Android Studio and AnyBurn.
- Add Arduino IDE, AWS SAM CLI, AWS Session Manager Plugin, AxCrypt, Azure Functions Core Tools.
- Add Bambu Studio, BleachBit, Blender, Box Drive, Bulk Crap Uninstaller and BurnAware Free.
- Add Calibrite Profiler, Certify The Web, Chef Workstation and Spectra PDF (catalog at 133).
- Verify downloaded installers carry an installer signature before staging.

## [1.4.0.15] - 2026-08-31

### Added

- Add Anaconda, AnyDesk, Brave, CCleaner, Citrix Workspace Current Release, CPU-Z and CutePDF Writer.
- Add Greenshot, Opera, pgAdmin 4, PyCharm, Slack, TreeSize Free and XenCenter.
- Add XenServer VM Tools and Zoom Workplace (catalog at 108).

## [1.4.0.14] - 2026-08-31

### Added

- Add an Intune-only deployment target that stages, builds and publishes without MECM.

## [1.4.0.13] - 2026-08-28

### Added

- Publish to Intune after Package with tenant credentials stored in Options.

## [1.4.0.12] - 2026-08-28

### Added

- Publish .intunewin packages to Intune through Microsoft Graph.

## [1.4.0.11] - 2026-08-28

### Added

- Record active command overrides in the stage manifest.

## [1.4.0.10] - 2026-08-28

### Added

- Edit install and uninstall command overrides per app in Deployment Conditions.

## [1.4.0.9] - 2026-08-28

### Added

- Apply per-app install and uninstall command overrides at package time.

## [1.4.0.8] - 2026-08-27

### Added

- Extend the Network split to the remaining M365 Apps, Project and Visio packagers.

## [1.4.0.7] - 2026-08-27

### Added

- Split M365 Apps x64 into Online and Managed deployment types by network.

## [1.4.0.6] - 2026-08-27

### Added

- Split Firefox into one deployment type per OS language.

## [1.4.0.5] - 2026-08-27

### Changed

- Drop the architecture marker from split application names.

## [1.4.0.4] - 2026-08-27

### Added

- Stage ARM64 variants of Chrome and Firefox with the Architecture split.

## [1.4.0.3] - 2026-08-27

### Added

- Stage a 7-Zip ARM64 deployment type with the Architecture split.

## [1.4.0.2] - 2026-08-27

### Added

- Select variant splits per app in the Deployment Conditions grid.

## [1.4.0.1] - 2026-08-27

### Added

- Support multiple deployment types per application from the stage manifest.

## [1.4.0.0] - 2026-08-17

### Added

- Attach architecture, OS language and network requirement rules per app through Deployment Conditions.

## [1.3.0.0] - 2026-08-17

### Added

- Drop an installer onto the window to analyze, preview and package it.
- Save a dropped installer as a new packager script.

## [1.2.0.1] - 2026-08-16

### Changed

- Restore the saved window size before maximizing on launch.

## [1.2.0.0] - 2026-08-16

### Changed

- Share logging, site connection and window plumbing with the tool suite.
- Stop pipeline teardown from freezing the window.

## [1.1.0.7] - 2026-08-14

### Added

- Choose nested or flat share folder layout in MECM Preferences.

## [1.1.0.6] - 2026-08-14

### Added

- Package existing PSADT folders as versioned content with verified hashes.

## [1.1.0.5] - 2026-08-14

### Changed

- Package from a standard user account; local administrator gates removed.

## [1.1.0.4] - 2026-08-14

### Added

- Fall back to Invoke-WebRequest when curl.exe cannot download.

## [1.1.0.3] - 2026-08-14

### Changed

- Detect .NET 8 and 10 runtimes by the current or successor patch version.

### Removed

- Remove the single-architecture .NET 10 packager.

## [1.1.0.2] - 2026-08-14

### Fixed

- Replace the deployment type when a version-less application gains a new version.
- Verify every deployment type of the speculative-execution mitigations application.

## [1.1.0.1] - 2026-08-14

### Added

- Filter the application grid by application, vendor, status or name.

### Fixed

- Keep grid columns readable when Debug Columns is toggled.

## [1.1.0.0] - 2026-08-13

### Added

- Create .intunewin packages during Package with a downloadable Content Prep tool.

## [1.0.0.12] - 2026-08-13

### Added

- Deploy to a test collection after content distribution.

## [1.0.0.11] - 2026-08-12

### Fixed

- Fix the speculative-execution mitigations packager's deployment-type check.

## [1.0.0.10] - 2026-08-12

### Added

- Add the speculative-execution mitigations packager with six deployment types.

## [1.0.0.9] - 2026-08-01

### Fixed

- Treat reboot-required exit codes as success in multi-installer and WinRAR wrappers.

## [1.0.0.8] - 2026-06-13

### Added

- Accept -VerboseLog on every packager and log full failure detail.

## [1.0.0.7] - 2026-06-12

### Changed

- Remove the per-entry site connection probe.

## [1.0.0.6] - 2026-06-12

### Fixed

- Rebuild a stale ConfigMgr site drive connection automatically.

## [1.0.0.5] - 2026-06-12

### Added

- Add verbose failure diagnostics and provider site-code validation.

### Fixed

- Honor the Provider Machine preference in Check MECM.

## [1.0.0.4] - 2026-05-27

### Added

- Add NVIDIA GeForce and RTX Enterprise driver packagers (91 packagers).

## [1.0.0.3] - 2026-05-13

### Fixed

- Restore the Check MECM site connection when the drive is pre-mounted.

## [1.0.0.2] - 2026-05-08

### Fixed

- Create the ConfigMgr site drive from the Provider Machine preference when absent.
- Various bug fixes.

## [1.0.0.1] - 2026-05-06

### Added

- Add the dual-architecture .NET 10 Desktop Runtime packager.

### Fixed

- Package with an empty comment.
- Time out idle packager processes instead of waiting forever.
- Various bug fixes.

## [1.0.0] - 2026-05-02

### Added

- Discover packager scripts, check vendor and MECM versions, stage and package from one window.
- Run One Click on tracked apps with cadence gating and MECM pre-flight.
- Verify staged and copied content against manifest hashes before creating applications.
- Configure MECM, packager, One Click and product filter settings in Options.
- Ship packagers for 7-Zip, Adobe Acrobat Reader, Firefox, Chrome, Edge, Office 365, Teams and more.
