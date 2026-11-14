# Changelog

## [Unreleased]

### Fixed

- Show a text box and a combo box with a border of 3:1 contrast or more in the light theme. The MahApps border was 1.6:1 on white.
- Show the muted text (notes, hints, the version) with MahApps.Brushes.Gray.MouseOver. The old brush was 2.7:1 on white in the light theme.
- Show the column headers of a disabled grid with the gray text of the other disabled controls.

### Added

- Add Invoke-OptionsContrastSmoke.ps1. It opens every section of Options in the light and in the dark theme and checks the contrast of the text (4.5:1) and of the control borders (3:1), enabled and disabled. Tests\KnownContrastDefects.ps1 lists the known limits that SuiteCommon owns.

## [2026.11.13.0137] - 2026-11-13

## 1 footer that shows the site and the site server

### Fixed

- Restore the weight of the section headers, the log header, the progress title, the dropped file name, the About, Welcome and dialog prompt titles, the Workbench headings and the Not Offered By Discover label to SemiBold. The light theme change had made them bold.
- Restore the normal weight of the About keys, the Detection, Product, Publish To and Install For labels, the Default Destinations label and four Workbench labels.
- Only field titles and the app title are bold.

### Added

- Add FontWeight.Guard.Tests.ps1 and FontWeightGuardHelpers.ps1. They fail when a bold weight is set on a label that is not a field title or an app title, in XAML, in the XAML of the scripts and in code that sets a weight at run time. A label before a panel is not a field title.
- Show the site code and the site server in the footer, at the right end: `Site MCM · cm01.contoso.com`. The footer shows `Site not set` before you set them, and it is hidden while ConfigMgr is not in Systems in use. The text is read-only. Select it to open Options. The footer refreshes after the window loads, after the Options dialog closes, after the Setup window closes and before Check ConfigMgr.
- Add Footer.Guard.Tests.ps1. It checks the footer status element, the shared footer font, the footer calls and the rule that the site code and the site server are editable only in Options and in the Setup window. It has bad-sample tests.
- Add the Installer Drop dialog to the capture tool, as the source of the README image drop-preview.png.

### Changed

- Show the site code box of the Convert ConfigMgr App to Intune dialog as read-only text. It shows the site code from Options. The error text names Options, Connections, ConfigMgr.
- Use the shared footer style of SuiteCommon for the status text, in both themes. The status element has no font setting of its own.
- Refresh the vendored SuiteCommon library to release 2026.11.13.0082.
- Update the README text and the user guide text. Retake the images of the user guide that show the main window with the footer (the dialog images have no footer and stay as they were), and the README images main-dark.png, setup.png, workbench.png and drop-preview.png. The capture tool takes them from the same run.
- Retake the four Chrome CVE images in both themes. The live NVD data now lists 54 CVEs for Chrome 157.0.8103.40, so the text of the CVE page names 54 and `54 of 54`.
- Rewrite the Setup step of the Chrome capture for the rebuilt Setup window. The capture of the Options pages now closes the Options window, so that it does not cover the next images.

## [2026.11.13.0136] - 2026-11-13

## 9 installer option fixes and a new Options window

### Added

- Open Options as a tree of sections with a search box.
- Add SQL instance and SQL port to the site database settings.
- Import your own Microsoft 365 configuration.xml and use it in the packagers.
- Add Test connection for ConfigMgr, WSUS and Intune.
- Add 8 Office Deployment Tool settings to Microsoft 365 Apps.
- Add Adobe Acrobat Reader properties, including EULA, shortcut, Chrome extension and update mode.
- Add Beyond Compare 5 policy switches for remote profiles and saved passwords.
- Add 11 Citrix Workspace App switches and the Enterprise Browser component.
- Add SSMS workloads, languages and package cache options.
- Add the installer options audit with the vendor source of each switch.

### Fixed

- Stop SSMS recommended and optional components from acting without a workload.
- Stop Microsoft 365 packages from forcing shared computer activation and architecture change.
- Remove the retired Bing ExcludeApp ID from Microsoft 365 packages.
- Fix the DBeaver System uninstall arguments and the AI setting text.
- Fix the Citrix App Protection tooltip.
- Mark the TeamViewer Host settings that only a MSI uses.

### Changed

- Use the same Connections fields in Options and in the Setup window.
- Show the WSUS port in one field.
- Use sentence case and bold field titles in Options in both themes.
- Leave shared computer activation and architecture change off in Microsoft 365 packages unless you select them.
- Emit SSMS recommended and optional components only with a workload.

## [2026.11.13.0135] - 2026-11-13

## 7 fixes to the light and dark themes

### Fixed

- Show white text on every blue background in the light theme.
- Show the Site Code and SMS Provider titles in bold 14 pt in both themes.
- Show field labels and headers in bold. Use Title Case for buttons, headers and dialog titles.
- Use one short sentence for each tooltip, in Configuration Manager words.
- Keep raw error text out of the window and the log pane. The log file holds it.
- Show notes in the theme text color, with no reduced opacity.
- Show a visible border on hover for every button.

### Changed

- Use one themed button style in the window, the dialogs and the Workbench.
- Define the shared styles one time. The dialogs and the Workbench take them at run time.
- Write the window log to a daily file in the Logs folder.
- Let the image capture take the lab lock and work in its own folder.

### Added

- Add guard tests for the theme, the fonts, the button styles, Title Case and error text.
- Require Pester 5.1.0 in every test file.

### Known limit

- In the light theme a disabled sidebar button keeps white text. The fix belongs in SuiteCommon.

## [2026.11.12.0134] - 2026-11-12

## 4 images of a Chrome CVE check

### Added

- Add images of the Chrome CVE grid and CVE window in both themes.
- Add the switch -OnlyCveChrome to the image capture command.

### Changed

- Keep files of other sessions in the lab folder when the image capture ends.

## [2026.11.12.0133] - 2026-11-12

## 21 corrected CPE names and a new CVE lookup

### Fixed

- Find the CVEs that NVD lists with one exact version or with no version range.
- Compare versions with pre-release words, patch letters and four parts correctly.
- Correct the CPE name of 21 packagers so that the lookup finds their CVEs.
- Show the time of a CVE lookup in the same form on PowerShell 7.
- Keep a patch of the ConfigMgr version when the latest version has fewer parts.
- Keep a CVE open when the latest version has no version data in the record.
- Replace a damaged KEV catalog copy and write the catalog through a temporary file.

### Changed

- Show the CVEs of the ConfigMgr version, also for a row with no update.
- Show how many CVEs the update fixes in the column **Fixed by update**.
- Show the status `Fixed`, `Open` or `No version` for each CVE in the CVE window.
- Show `Not in NVD` when NVD holds no record under the product name.
- List NVD records with no version apart and do not count them.
- Read all NVD records of a product one time and use them for six hours.
- Use the saved NVD records when NVD does not answer.
- Save each CVE result in its own file. The old cache file is removed.
- Accept more than one CPE name in a packager header. A CVE counts one time.

## [2026.11.11.0132] - 2026-11-11

## 68 test files run in CI on 2 PowerShell versions

### Added

- Add the GitHub Actions workflow `checks.yml`: lint, Pester tests in three groups on Windows PowerShell 5.1 and PowerShell 7, and the window probes.
- Add `Invoke-Checks.ps1`, the command that the `Checks` workflow runs.
- Add a Testing section to the README.

### Fixed

- Fix the CVE dialog probe and the discovery probe. They did not load functions that the app now calls.
- Fix the child bridge probe. It failed when the temp folder path was in 8.3 short form.

## [2026.11.07.0131] - 2026-11-07

## 5 theme fixes and 1 change

### Fixed

- Recolor the Workbench and Add Installer buttons when the theme changes.
- Show the selected item in the Workbench and Options lists in readable colors.
- Show disabled Workbench fields without a light grey block in the dark theme.
- Wrap the Workbench action buttons so the last button is not cut off.
- Dim the title bar of a dialog when it loses focus, like its owner.

### Changed

- Draw Workbench and dialog buttons with a thin light outline.

## [2026.11.07.0130] - 2026-11-07

## 46 images in a user guide

### Fixed

- Show the full header of every column in the grids of the dialogs.
- Widen the Options window so that the application names stay readable.

### Added

- Add the user guide with images of the running window.
- Add one command that regenerates every image of the guide on a lab computer.

## [2026.11.05.0129] - 2026-11-05

## 5 fixes and changes

### Fixed

- Wait for the GitHub Desktop uninstall entry before the install reports success.
- Count CVE results not looked up in the run apart in the monitor summary.
- Log the ConfigMgr query of the CVE monitor.

### Changed

- Show the same progress display for long actions in the main window.
- Cancel Discover unmanaged software from the progress display.

## [2026.11.05.0128] - 2026-11-05

## 12 fixes

### Fixed

- Stop two sessions from overwriting each other's CVE cache.
- Report a Shodan 404 as a lookup error, not an empty list.
- Mark a truncated NVD result as incomplete.
- Reject a CVE response that lacks required fields.
- Harden CVE lookup requests.
- Harden CVE report files.
- Hide a saved CVE count when its recorded product differs from the packager.
- Remove expired CVE reports and use one version comparer in the CVE views.
- Check the observed site version in One Click; keep the site drive on the provider.
- Widen the CVE grid columns so the headers are no longer cut off.
- Try WinRM when DCOM returns no site definition during discovery.
- Report a Discover choice that was not saved or refreshed.

## [2026.11.05.0127] - 2026-11-05

## 1 change

### Changed

- Document the pull request flow for releases.

## [2026.11.05.0126] - 2026-11-05

## 1 addition

### Added

- Convert a ConfigMgr application to Intune from the sidebar of the app.

## [2026.11.05.0125] - 2026-11-05

## 5 additions

### Added

- Convert an existing ConfigMgr application to Intune Win32 apps.
- Make one Intune app for each convertible deployment type of an application.
- Keep ConfigMgr's first-match order between deployment types with Intune requirement rules.
- Convert Or and grouped detection to one Intune detection script.
- List dependencies and supersedence of a converted application in a file.

## [2026.11.04.0124] - 2026-11-04

## 3 fixes

### Fixed

- Detect and uninstall Tableau Desktop from the 64-bit registry.
- Detect and uninstall Tableau Prep Builder from the 64-bit registry.
- Detect and uninstall Tableau Reader from the 64-bit registry.

## [2026.10.29.0123] - 2026-10-29

## 2 fixes

### Fixed

- Keep ConfigMgr discovery result tables intact when query output includes other values.
- Capitalize Discover Unmanaged Software in its button and dialog.

## [2026.10.25.0122] - 2026-10-25

## 29 fixes and changes

### Security

- Fix several security issues.

### Fixed

- Detect 18 packagers by installed version instead of file presence.
- Install Vim, pgAdmin 4 and AIMP silently on a clean computer.
- Remove Anaconda files completely on uninstall.
- Remove the AWS Session Manager plugin on uninstall.
- Accept Android Studio install exit code 1223 as success.
- Find the Vim uninstall entry under its real name.
- Request the x64 build of Zoom Workplace.
- Skip a batch item only when its build is current for that scope.
- Refuse Intune targets in the command-line Package mode.
- Stop a pinned version that differs from the vendor's current release.
- Keep earlier WSUS updates when one approval fails.
- End WSUS child processes that run too long.
- Restore the previous App Packager files when an update fails.
- Keep non-ASCII text in Workbench scripts.
- Keep your own Workbench scripts when generated ones are replaced.
- Keep the right profile when importing a Workbench bundle.
- Require a trusted timestamp when keeping existing signatures.
- List unfixed CVEs in HTML and Markdown reports.
- Quote paths with spaces when starting processes.
- Escape quotes and line breaks in drop-to-package names.
- Report MSIX wrapper failures instead of continuing.
- Send Intune app metadata once, with the final commit.
- Label ARM64 builds correctly for ad-hoc MSI installs.

### Changed

- Store the PostgreSQL and MariaDB administrator password per computer, encrypted.
- Build Options window panels when first shown.
- Load GitHub status in the background.
- Read the packager catalog once per change.
- Update the installer analysis library.
## [2026.10.20.0121] - 2026-10-20

## 3 fixes

### Fixed

- Read only distinct software titles in Discover unmanaged software.
- Stop an active discovery scan with Cancel.
- Check and download the current 64-bit Zoom Workplace release.

## [2026.10.18.0120] - 2026-10-18

## 3 fixes

### Fixed

- Treat a ConfigMgr title with an unrelated word after a dash as another application.
- Read and stage WinSCP versions with two parts.
- Read XenServer VM Tools versions with four parts from the downloads page.

## [2026.10.17.0119] - 2026-10-17

## 287 packagers matched against hardware inventory

### Added

- Add Discover unmanaged software to find hidden apps that clients run without a ConfigMgr application.
- Show found apps in the grid with Enable selected or Enable all.
- Keep checked apps out of later discoveries with Don't ask again for selected.
- Add Discovery SQL and Discovery DB overrides to Options, ConfigMgr Preferences.
- Add the ArpName header tag to every packager.
- Add the optional ArpVersion and ArpArch header tags.

### Fixed

- Check ConfigMgr matches applications by installer method and architecture.
- Check ConfigMgr treats 64-bit and amd64 as x64, and 32-bit as x86.
- Read the SQL server name from the SMS Provider over DCOM, then WinRM.

## [2026.10.14.0118] - 2026-10-14

### Fixed

- Use NVD's CPE filter URL to preserve scoped results on its current search page.
- Search external CVE sites with readable product/vendor names; retain canonical CPE searches for NVD.
- Log Visual Studio Code User's destination path for post-package integrity verification.

## [2026.10.12.0117] - 2026-10-12

## 1 PowerShell 7 feature

### Added

- Publish to WSUS under PowerShell 7 through a Windows PowerShell child process.

## [2026.10.12.0116] - 2026-10-12

## 1 PowerShell 7 fix

### Fixed

- Keep the PowerShell 7 module roots when the app runs under PowerShell 7.

## [2026.10.11.0115] - 2026-10-11

## 2 PowerShell 7 fixes

### Fixed

- Read installer headers without the byte encoding that PowerShell 7 removed.
- Read CVE cache timestamps the same way under Windows PowerShell and PowerShell 7.

## [2026.10.11.0114] - 2026-10-11

## 6 fixes

### Fixed

- Skip One Click when the vendor reports a version older than the one last seen.
- Skip a destination whose published version is newer than the vendor version.
- Count a publish record or refusal only for the site, server or tenant it was written for.
- Publish again when the Workbench profile or the signing policy changed since the last publish.
- Ignore CVE records for the Android, Linux or macOS build of a product.
- Mark a CVE report row whose latest lookup failed or was not cached.

## [2026.10.11.0113] - 2026-10-11

## 7 CVE data fields

### Added

- Show the fix version, published date and modified date per CVE in the window and every report.
- Count the CVEs fixed in the latest release, in an earlier release, or with no fix version.
- Show the newest CVE date and the last NVD update per app.
- Add In latest and Newest CVE columns to the grid; NVD updated sits with the debug columns.

## [2026.10.11.0112] - 2026-10-11

## 1 window fix

### Fixed

- Double-click the title bar to maximize or restore the window.

## [2026.10.11.0111] - 2026-10-11

## 1 suite update notice

### Added

- Show the suite update inside a suite install; Update now opens the suite launcher.

## [2026.10.11.0110] - 2026-10-11

## 4 grid and CVE window fixes

### Fixed

- Show the CVE count on a selected row and widen the CVEs and Status columns.
- Show a preloaded ConfigMgr version with its compare result and CVE count on launch.
- Make the links in the CVE window readable in the dark theme.
- Clip long CVE summaries to six lines; the full text sits in the tooltip.

## [2026.10.11.0109] - 2026-10-11

## 3 fixes

### Fixed

- Run the CVE window's Refresh off the UI thread and show the new result.
- Open the CVE window beside the main window instead of over it.
- Write only https vendor references to the CVE report CSV.

## [2026.10.11.0108] - 2026-10-11

## CVE lookup for 294 packagers

### Added

- Look up the CVEs an update fixes for every checked application.
- Show CVEs, Max CVSS and KEV columns in the grid.
- Open a CVE window with links to six CVE sites per CVE.
- Write the CVE report as HTML, CSV, XLSX, Markdown and JSON.
- Record the ConfigMgr version with its site and show it on the next launch.
- Run the Version Monitor as a batch mode of the main script.
- Post the monitor run summary to a webhook and copy reports to a folder.
- Add the NVD API key and the lookup-after-Check-ConfigMgr preference.
- Add the vendor security advisory page to 94 packagers.
- Add the CPE to RenderDoc, Sysinternals and draw.io.
- Add Open CVE Sources and Copy CVE IDs to the row menu.

### Changed

- Remove the Version Monitor module and its NVD cache.

### Documentation

- Document the CVE lookup, the report formats and the monitor configuration.

## [2026.10.10.0107] - 2026-10-10

## 8 update checks fixed

### Fixed

- Report a vendor release with more version parts than ConfigMgr holds as an update.
- Compare all four version parts when a vendor feed steps back.
- Read the version from the first output line that is a version.
- Record Firefox ESR as 140.17.0 instead of 140.17.0esr.
- Keep the Liberica JDK 21 build suffix as the vendor prints it.
- Run the Windows Admin Center detection key check in Stage only.
- Retry the SMath Studio page fetch twice; stop after 20 seconds.
- Compare a ConfigMgr version that carries a channel suffix or a fifth number.

### Documentation

- Add the Network requests section: which process sends each call and where.
- Match the application table to the packager headers and detection blocks.
- Document the WsusSupport and LocalSource header tags.
- List the One Click module, vendored modules, Samples and Tests in Project Structure.

## [2026.10.07.0106] - 2026-10-07

## 1 shared module version synchronized

### Changed

- Sync the shared SuiteCommon module to 2026.10.07.0049.

## [2026.10.07.0105] - 2026-10-07

## 3 update checks fixed

### Fixed

- Find ConfigMgr titles with the version mid-name or the architecture attached.
- Report a vendor release that adds a fourth version part as an update.
- Skip a vendor release older than the version ConfigMgr or history holds.
## [2026.10.07.0104] - 2026-10-07

## 312 packagers declare WSUS support

### Added

- Show which packagers support WSUS in the One Click plan before a run.
- Suggest WSUS database maintenance when a publish times out.

### Changed

- Check download addresses taken from vendor pages before passing them to curl.
- Escape vendor text placed in generated install and detection scripts.
- Check script signatures again right before sending them to ConfigMgr or Intune.
- Refuse staged content that no longer matches its stage manifest.
- Refuse a network content folder that holds files the manifest does not name.

### Fixed

- Report failure when an installer cannot start instead of reporting success.
- Stage scripts and hooks typed in the Application Workbench.
- Accept the CCleaner exit code returned when no desktop session exists.
- Detect CutePDF Writer, Jabra Direct and Windows Admin Center by stable keys.
- Find the newest Defraggler and LibreOffice builds the vendors actually serve.
- Find the UltraVNC uninstaller through its uninstall entry.
- Read Unity Hub installers from the current vendor feed.
## [2026.10.02.0103] - 2026-10-02

## 312 packagers

### Added

- Package MuleSoft Anypoint Studio from the vendor's latest ZIP to C:\AnypointStudio.
## [2026.10.01.0102] - 2026-10-01

## 1 crash fixed

### Fixed

- Stop the app from closing when Options opens.
## [2026.10.01.0101] - 2026-10-01

## 311 packagers leave published content unchanged on Skip

### Changed

- Package PyCharm 2026 instead of the discontinued PyCharm Community.

### Fixed

- Leave share content unchanged when an existing application is kept.
- Apply the theme to the Content Prep, icon pack and About buttons.
## [2026.09.30.0100] - 2026-09-30

## 27 Adobe Acrobat Reader languages selectable

### Added

- Package the multilingual Adobe Acrobat Reader edition with chosen languages.
- Choose the Adobe Acrobat Reader edition and languages in Packager Preferences.

## [2026.09.30.0099] - 2026-09-30

## 3 destinations planned, run and reported per application

### Added

- Open a One Click plan before the run: one row per tracked application.
- Choose One Click destinations per application: ConfigMgr, WSUS, Intune.
- Show One Click progress per destination in the plan window.
- Write a One Click report per run with identifiers and rollback steps.
- Skip a version a destination already has; Force publishes it again.
- Decide an existing ConfigMgr application by policy instead of a prompt.
- Remove the version from an application name when the vendor puts it there.

### Fixed

- Keep staged applications checked for the publish step.
- Stop One Click from staging unchanged applications again.
- Name every default destination in the Setup log line.
- Widen the selection column.
## [2026.09.30.0098] - 2026-09-30

## 5 screenshots show the three publish destinations

### Changed

- Name the three publish destinations in About.
- Show Setup, the WSUS options, and the WSUS Updates window in the README.
## [2026.09.30.0097] - 2026-09-30

## 3 publish destinations chosen per run

### Added

- Publish to ConfigMgr, Intune, or WSUS from three sidebar buttons.
- Publish installers to WSUS as updates for computers with an older version.
- Manage published WSUS updates: approve, decline, expire, remove.
- Import Microsoft Update Catalog updates by ID or link.
- Choose the systems in use in Setup and Options.
- Build the .intunewin at Stage when the tool is present.
- Add Visual C++ x64 and x86 runtime packagers, RTools, and Slido.
- Run WSUS publishing from the command line.

### Fixed

- Detect ASP.NET Core hosting bundles by their own entry, not a shared runtime key.
- Enable sidebar buttons after Setup or Options saves the missing setting.
- Keep the publish note in the log after a run.
- Follow the SSL choice with the WSUS port without a click.
- Open Setup after the main window has rendered.
- Run version checks and publishes without a site code when ConfigMgr is not in use.
- Find DBeaver and Nextcloud in the console lookup.
## [2026.09.25.0096] - 2026-09-25

## 6 user reports closed

### Fixed

- Clear a packager's checkbox after a successful Stage or Package.
- Keep ConfigMgr and online version results through Options, filters, and grid rebuilds.
- Find RStudio, Git, Thunderbird, Node.js, Webex, and Corretto applications in the console lookup.
- Stop the console lookup from matching a longer application name.
- Validate every Citrix Workspace switch against the 2603.11 and 2607 vendor pages.
- Let the Citrix store policy and auto-update options stay unset.

### Added

- Choose deployment type content fallback and slow-network behavior in Options.

## [2026.09.25.0095] - 2026-09-25

## 12 OpenJDK packagers verified on a device

### Fixed

- Detect the per-user OpenJDK MSI by its java.exe; its Add/Remove entry lands under HKLM.

## [2026.09.25.0094] - 2026-09-25

## 307 packagers

### Added

- Add Microsoft Build of OpenJDK 17: MSI and EXE, per machine and per user.
- Add Microsoft Build of OpenJDK 21: MSI and EXE, per machine and per user.
- Add Microsoft Build of OpenJDK 25: MSI and EXE, per machine and per user.

## [2026.09.25.0093] - 2026-09-25

## 295 packagers

### Added

- Update the shared SuiteCommon module to 2026.09.25.0033.
- Add Signing Suite.
- Add AppPackager Suite for per-user installs.
- Add Jabra Direct.

## [2026.09.21.0092] - 2026-09-21

### Fixed

- Use the site code and provider from the suite launcher when none is saved.

### Changed

- Update the README screenshots.

## [2026.09.21.0091] - 2026-09-21

### Changed

- Use the product name ConfigMgr in the application and the documentation.

### Fixed

- Fix the Package step for ASP.NET Core 10 Server Hosting Bundle.

## [2026.09.18.0090] - 2026-09-18

### Fixed

- Use the icon pack icon for every packager that has one.

## [2026.09.18.0089] - 2026-09-18

### Added

- Add ASP.NET Core 10 Server Hosting Bundle.
- Add Citrix Workspace app LTSR for x86, x64 and ARM64.
- Add Tableau Desktop, Prep Builder and Reader.
- Add a per-packager folder for installers you download yourself.

## [1.6.0.9] - 2026-09-17

### Added

- Add GitHub Desktop for per-user installs.

## [1.6.0.8] - 2026-09-17

### Added

- Add Beyond Compare 5.

### Fixed

- Stop the Application Workbench freezing on packagers created from a dropped installer.
- Apply icons and extra files chosen in the Application Workbench.

## [1.6.0.7] - 2026-09-16

### Fixed

- Keep MECM versions and compare results after saving Options.

## [1.6.0.6] - 2026-09-15

### Fixed

- Fix Check latest failing for Apache NetBeans.

## [1.6.0.5] - 2026-09-15

### Added

- Add an Options setting to include the version in every application name.

### Fixed

- Fix application title choices saved in the Application Workbench never applying.

## [1.6.0.4] - 2026-09-14

### Fixed

- Various bug fixes.

## [1.6.0.3] - 2026-09-14

### Changed

- Include the test suite in source code downloads.

## [1.6.0.2] - 2026-09-13

### Fixed

- Various bug fixes.

## [1.6.0.1] - 2026-09-13

### Changed

Deployment conditions moved to Application Workbench

### Removed
Deployment Conditions options panel

## [1.6.0.0] - 2026-09-12

### Added

- Application Workbench for fine-tuning packagers.
- Script signing.
- Previously gitignored and private test suite.

### Fixed

- Various bug fixes.

### Removed

- BurnAware Free

## [1.5.2.2] - 2026-09-09

### Added/Changed

- Velopack installer support via installer-analysis 1.3.4.0.

### Fixed

- Fixed HKCU detection for user installs.

## [1.5.2.1] - 2026-09-09

### Added

- Per-app versioned or perpetual application titles.
- Overwrite/skip/cancel for existing applications.

### Fixed

- Fixed the task-sequence option breaking user installs.

## [1.5.2.0] - 2026-09-07

### Changed

- Removed the 7-Zip requirement from TeamViewer Host

### Fixed

- Various bug fixes

## [1.5.1.9] - 2026-09-05

### Fixed

- Fixed SpectraPDF install params.

## [1.5.1.8] - 2026-09-04

### Fixed

- Fixed KDiff3 uninstall params.

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
