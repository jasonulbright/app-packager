# Installer options audit

Each switch and property that App Packager writes for seven applications, compared with the vendor documentation. Checked 2026-11-13. A row says Correct only when a vendor page confirms the syntax. A row says UNVERIFIED when no vendor page confirms it.

Status words: Correct, Wrong, Missing, Not applicable, UNVERIFIED.

## Result

| Application | Wrong | Missing | UNVERIFIED | Wrong fixed | Missing added |
|---|---|---|---|---|---|
| Adobe Acrobat Reader | 0 | 6 | 1 | 0 | 4 (EULA_ACCEPT, DISABLEDESKTOPSHORTCUT, ENABLE_CHROMEEXT, DISABLE_BROWSER_INTEGRATION) and UPDATE_MODE, verified on the Adobe page |
| Beyond Compare 5 | 0 | 2 | 2 | 0 | 2 |
| Citrix Workspace App | 1 | 13 | 3 | 1 | 11 |
| DBeaver Community | 2 | 1 | 2 | 2 | 0 |
| Microsoft 365 Apps | 3 | 11 | 4 | 3 | 8 |
| SQL Server Management Studio | 1 | 5 | 1 | 1 | 3 |
| TeamViewer Host | 2 | 5 | 4 | 2 | 0 |
| Total | 9 | 43 | 17 | 9 | 29 |

Skipped Missing items, with the reason, are in the Skipped list of each application. The main reasons: the packager has no step for it (SSMS offline layout, SSMS update channel, Microsoft 365 Version MatchInstalled), the detection rule fixes the value (Citrix INSTALLDIR, DBeaver /D), the setting needs a list editor (Citrix STORE1 and later), or the option applies to the MSI only and the packager installs the EXE (TeamViewer Host). Microsoft 365 AppSettings and the tenant association key are set through your own configuration.xml (Options, Applications, Microsoft 365 Apps).

Wrong items fixed: Citrix App Protection tooltip; DBeaver AI label and tooltip; DBeaver System uninstall arguments in the manifest; Microsoft 365 MigrateArch, SharedComputerLicensing and the Bing ExcludeApp ID; SSMS includeRecommended and includeOptional without a workload; TeamViewer Host settings that the EXE installer ignores (UI text and the desktop shortcut tooltip).

Vendor sources are named in each table.

## 1. Adobe Acrobat Reader

Packager: `Packagers/package-adobereader.ps1`. The install runs `setup.exe /sAll /rs`. MUI properties go to `[Product] CmdLine=` in `setup.ini`. The uninstall runs `msiexec /x {ProductCode} /qn /norestart`.

| Option | Where in repo | Vendor source URL | Status | Finding and exact fix |
|---|---|---|---|---|
| `setup.exe /sAll` | package-adobereader.ps1:524 | .../AdminGuide/cmdline.html | Correct | Adobe lists `/sAll` as silent, no UI. |
| `setup.exe /rs` | package-adobereader.ps1:524 | .../AdminGuide/cmdline.html | Correct | Adobe lists `/rs` as "do not reboot". `/rps` is not needed. |
| `CmdLine=` in `[Product]` section of setup.ini | package-adobereader.ps1:196-226, 485 | .../AdminGuide/cmdline.html | Correct | Adobe says `[Product] CmdLine` holds space-separated Adobe properties and MSI switches. `[Startup] CmdLine` holds setup.exe switches. The repo writes to `[Product]`. |
| `LANG_LIST="a,b"` | package-adobereader.ps1:183 | .../AdminGuide/properties.html | Correct | Domain: one locale code, a comma list, or `ALL`. Default: not stated. MUI installer only. A multi-language list needs a MUI serial or trial per Adobe. |
| `LANG_LIST=ALL` (UI "All languages") | package-adobereader.ps1:164, 182 | .../AdminGuide/properties.html | Correct | Adobe documents `ALL` upper case. The repo maps stored `All` to `ALL`. Property values are case-sensitive per Adobe. |
| `SUPPRESSLANGSELECTION=1` | package-adobereader.ps1:183 | .../AdminGuide/properties.html | Correct | Domain: `1` or `0`. Default: `1`. Version 11 and later. MUI only. The repo value is the same as the default. Harmless. |
| Edition English / MUI (installer file name suffix `_en_US` / `_MUI`) | package-adobereader.ps1:147, 171-175 | not a switch | Not applicable | Chooses the installer file, not a switch. No vendor switch to check. |
| MUI locale list (27 codes) | package-adobereader.ps1:135, start-apppackager.ps1:6704 | none found | UNVERIFIED | Adobe pages list no locale code table. The repo says it copied the codes from the MUI `setup.ini`. Re-check by reading `Languages=` in the staged `setup.ini` of each MUI release. |
| Uninstall `msiexec /x {code} /qn /norestart` | package-adobereader.ps1:531 | .../AdminGuide/cmdline.html | Correct | Adobe shows `msiexec /x {Product Code} /qn`. Braces required. `/norestart` is a standard msiexec switch. |
| `EULA_ACCEPT=YES` | none | .../AdminGuide/properties.html | Missing | Syntax: `EULA_ACCEPT=YES`. Domain: `YES`. Default: not set (EULA shows at first launch). Version 9.0 and later. Tooltip: "Accepts the Adobe license agreement for all users. The EULA dialog does not show at first start." |
| `DISABLEDESKTOPSHORTCUT=1` | none | .../AdminGuide/properties.html | Missing | Syntax: `DISABLEDESKTOPSHORTCUT=1`. Domain for Reader: `1`. Default: not set (Adobe creates the shortcut). Applies at install and at every update. Tooltip: "Does not create the Reader shortcut on the desktop. Use it on shared and kiosk devices." |
| `UPDATE_MODE=n` | none | .../AdminGuide/properties.html | Missing | Syntax: `UPDATE_MODE=3`. Domain: `0`, `2`, `3`, `4`. Meaning: 0 manual, 2 download and let the user choose, 3 download and install automatically, 4 notify only. Default: not stated. Reader only. The page text says "If updates have not been disabled". Re-read the meaning column before coding. Tooltip: "Sets how Reader gets updates. Use 0 when ConfigMgr patches Reader." |
| `ENABLE_CHROMEEXT=0` | none | .../AdminGuide/properties.html | Missing | Syntax: `ENABLE_CHROMEEXT=0`. Domain: `1`, `0`. Default: the Continuous track installs the Chrome extension. Version 12.x (Feb 2017) and later. Tooltip: "Blocks the Adobe Chrome extension at install. Use it when a policy controls browser extensions." |
| `DISABLE_BROWSER_INTEGRATION=YES` | none | .../AdminGuide/properties.html | Missing | Syntax: `DISABLE_BROWSER_INTEGRATION=YES`. Domain: `YES`. Default: not set. The property controls the PDF viewer ActiveX control for IE. Adobe suggests it for Terminal Servers. Tooltip: "Does not install the Reader browser plug-in. Use it on RDS hosts." |
| MSI verbose log | none | .../AdminGuide/cmdline.html | Missing | Syntax in `[Product] CmdLine`: `/L*v "C:\Windows\Temp\AcroRdr_install.log"`. The Adobe page shows `/L*v` only in msiexec examples. setup.exe passes `[Product] CmdLine` to msiexec. Confirm on the lab. Tooltip: "Writes a verbose MSI log. Use it to find a failed install on a client." |

Skipped (reason)
- `/sPB`: shows a progress bar. A ConfigMgr install runs with no UI.
- `/rps`: restarts the device with no warning. Unsafe.
- `/sl "LANG_ID"`: single-language choice for the MUI installer. `LANG_LIST` already covers it.
- `SUPPRESSLANGSELECTION=0`: shows a language dialog to each new user. Not silent.
- `DISABLE_ARM_SERVICE_INSTALL=1`: Windows 7 and Vista only.
- `INSTALLDIR`: the detection path in the repo is fixed to the default folder.
- `REMOVE_PREVIOUS`: default is `YES` already.
- `ISX_SERIALNUMBER`, `PRESERIALIZATIONFILEPATH`, `ROAMLICENSING`, `ROAMIDENTITY`: licensing for Acrobat. Reader has no serial.
- `DISABLE_DISTILLER`, `DISABLE_PDFMAKER`, `REMOVE_PREVIOUS_READER`: Acrobat only.
- `DISABLE_ASIAN_FONTS=YES`: documented, useful only on size-limited images. Optional.
- `IGNOREVCRT64`, `IGNORE_INTEROP_REQS`, `ENABLE_OPTIMIZATION`, `ENABLE_CACHE_FILES`, `DISABLE_CACHE`: packaging internals or slow the install.
- `OWNERSHIP_STATE`, `LEAVE_PDFOWNERSHIP`, `IW_DEFAULT_VERB`: Adobe marks them deprecated on Windows 10. Use ConfigMgr file-type association policy.
- `ADD_THUMBNAILPREVIEW`: Adobe gives no value domain. UNVERIFIED. Skipped.
- `SOURCELIST`: network install sources. A ConfigMgr package has none.

---

## 2. Beyond Compare 5

Packager: `Packagers/package-beyondcompare5.ps1`. Install: `/VERYSILENT /NORESTART /ALLUSERS /DISABLEUPDATES /SUPPRESSMSGBOXES` (line 120). Uninstall: `/VERYSILENT /NORESTART /SUPPRESSMSGBOXES` (line 121). The license file `BC5Key.txt` sits next to the installer (lines 118, 290-291).

| Option | Where in repo | Vendor source URL | Status | Finding and exact fix |
|---|---|---|---|---|
| `/VERYSILENT` | package-beyondcompare5.ps1:120 | https://scootersoftware.com/kb/netsetup | Correct | Inno Setup. No wizard and no progress window. |
| `/NORESTART` | package-beyondcompare5.ps1:120-121 | https://scootersoftware.com/kb/netsetup | Correct | Without it `/VERYSILENT` reboots with no prompt. |
| `/ALLUSERS` | package-beyondcompare5.ps1:120 | https://scootersoftware.com/kb/netsetup | Correct | BC5 switch. Installs for all users. |
| `/DISABLEUPDATES` | package-beyondcompare5.ps1:120 | https://scootersoftware.com/kb/dl5_winalternate | Correct | Disables the "Check For Updates" action. Alias `/DisableCheckForUpdates` exists from BC 5.2. |
| `/SUPPRESSMSGBOXES` | package-beyondcompare5.ps1:120-121 | https://scootersoftware.com/kb/netsetup | Correct | Works only with `/SILENT` or `/VERYSILENT`. The repo uses `/VERYSILENT`. |
| Uninstall `/VERYSILENT /NORESTART /SUPPRESSMSGBOXES` | package-beyondcompare5.ps1:121 | https://scootersoftware.com/kb/netsetup | Correct | The page lists `/VERYSILENT /SUPPRESSMSGBOXES /NORESTART` for uninstall. |
| `BC5Key.txt` beside the installer | package-beyondcompare5.ps1:118, 290-291, 393 | https://scootersoftware.com/kb/netsetup | Correct | The page says the setup copies the key file to the user system and registers BC. File name is exact. The packager copies the file and never reads it. |
| Where the setup copies BC5Key.txt when it runs as SYSTEM with `/ALLUSERS` | none | https://scootersoftware.com/kb/netsetup | UNVERIFIED | The page says "to their system". It does not state the target folder for an all-users install. Check on the lab: run the install as SYSTEM, then list `C:\ProgramData\Scooter Software\Beyond Compare 5` and `%APPDATA%`. |
| `/DisableRemoteProfiles` | none | https://scootersoftware.com/kb/netsetup | Missing | Syntax: `/DisableRemoteProfiles`. No value. Default: profiles on. BC 5.2 and later. Disables FTP, SFTP, WebDAV, Dropbox, Amazon S3 and OneDrive profiles. Tooltip: "Turns off remote profiles in Beyond Compare. Use it when policy blocks file transfer to outside hosts." |
| `/DisableSavedPasswords` | none | https://scootersoftware.com/kb/netsetup | Missing | Syntax: `/DisableSavedPasswords`. No value. Default: passwords can be saved. BC 5.2 and later. Passwords stay in memory only. The save-password boxes are hidden. Tooltip: "Stops Beyond Compare from saving passwords on disk. Users type the password each session." |
| `/LOG="path"` | none | https://jrsoftware.org/ishelp/topic_setupcmdline.htm | UNVERIFIED | Inno Setup documents `/LOG="filename"` (fixed path, overwrites). The Scooter pages do not list it. Test on the lab before use. Tooltip: "Writes the setup log to a file. Use it to find a failed install." |

Skipped (reason)
- `/CURRENTUSER`: per-user install. The packager is machine-wide and its detection is machine-wide.
- `/DIR="path"`: the repo detection path is fixed.
- `/GROUP`, `/NOICONS`: cosmetic. `/NOICONS` also removes the Start menu entry users need.
- `/SP-`: the Scooter page lists it for BC2 to BC4 only.
- `/SILENT`: shows a progress window. `/VERYSILENT` is the right switch.
- `/SAVEINF`, `/LOADINF`: authoring aids, not deployment options.
- `/32`: BC 4.1.4 era switch. BC5 packager is x64.
- `/LANG`, `/TASKS`, `/COMPONENTS`, `/MERGETASKS`: Inno generic. The Scooter page says nothing about BC5 task or component names. Skipped.
- MSI: Scooter says "MSI installers were removed for BC5".

---

## 3. Citrix Workspace app (Current and LTSR)

Packagers: `Packagers/package-citrixworkspace-cr.ps1` (stream `Current`) and `package-citrixworkspace-ltsr-x64.ps1`, `-x86.ps1`, `-arm64.ps1` (stream `LTSR`). The function `ConvertTo-CwaInstallArguments` is the same in all four files. In the LTSR files the line numbers are 41 higher than in `-cr.ps1` for lines 164-275 (x64: `/CleanInstall` line 210, `ADDLOCAL` line 314). The Current packager reads doc set 2610. The LTSR packagers read doc set 2607 LTSR (catalog LTSR version 26.7.1.4).

Config file: `Packagers/citrix-workspace-switches.json` (not in the repo until saved from the GUI).

| Option | Where in repo | Vendor source URL | Status | Finding and exact fix |
|---|---|---|---|---|
| `/silent` | package-citrixworkspace-cr.ps1:164 | 2610 page, "Common parameters" | Correct | "Disables installation dialogs and prompts". |
| `/silent /uninstall` | package-citrixworkspace-cr.ps1:331 | 2610 page, "Uninstall using the command-line interface" | Correct | Vendor line: `CitrixWorkspaceApp.exe /silent /uninstall`. |
| `/CleanInstall` (UI "Clean Install") | package-citrixworkspace-cr.ps1:169; UI start-apppackager.ps1:6853 | 2610 page, install overview | Correct | No value. Removes leftover files and registry values, then installs. Citrix recommends it for upgrades. Since 2503 it works with App Protection. |
| `/includeSSON` | package-citrixworkspace-cr.ps1:171-173 | 2610 page, "Authentication parameters" | Correct | Admin install, fresh install only. Since 2503 SSON installs dormant even without the switch. GPO "Enable pass-through authentication" turns it on. A reboot is needed after the GPO change. |
| `ENABLE_SSON=Yes\|No` | package-citrixworkspace-cr.ps1:174 | 2610 page, "Enable single sign-on" | Correct | Domain: `Yes`, `No`. Default: `Yes`. Takes effect with `/includeSSON`. The repo emits it only with `/includeSSON`. |
| `startAppProtection` | package-citrixworkspace-cr.ps1:177 | 2610 page, "Start App Protection" | Correct | Replaces `includeAppProtection` (deprecated in 2212). Needs admin install. |
| UI tooltip "Installs anti-keylogging and anti-screen capture protection" | start-apppackager.ps1:6855 | 2610 page, "Start App Protection" | Wrong | Since 2212 the installer always installs App Protection. `startAppProtection` starts the component. Fix the tooltip to: "Starts the App Protection component. The installer already installs it. Needs an admin install." |
| `ENABLEPRELAUNCH=True\|False` | package-citrixworkspace-cr.ps1:178-180 | 2610 page, "Control session pre-launch" | Correct | Domain: `True`, `False`. Default: not stated. |
| `SELFSERVICEMODE=True\|False` | package-citrixworkspace-cr.ps1:181-183 | 2610 page, "Control access to the self-service" | Correct | Domain: `True`, `False`. Default: not stated. `False` hides the self-service UI. |
| `/InstallMSTeamsPlugin=Y\|N` (Current) | package-citrixworkspace-cr.ps1:190-192 | 2610 page, "Install Third Party plug-ins" | Correct | Domain: `Y`, `N`. Default on since 2508. `Y` re-adds the plug-in on upgrade. The plug-in needs an admin install. |
| MS Teams plug-in on LTSR (repo ignores it) | package-citrixworkspace-ltsr-x64.ps1:237-239 | 2607 LTSR page | UNVERIFIED | The 2607 LTSR page has no `InstallMSTeamsPlugin`. The 2402 LTSR page lists ADDLOCAL key `MTOPBootStrapperInstaller` for the Teams plug-in. No evidence for 2607. The repo behavior (ignore and warn) is safe. Re-check when 2607 LTSR adds a switch. |
| `Installzoomplugin=N` (Current) | package-citrixworkspace-cr.ps1:193 | 2610 page, "Zoom VDI Plugin Management" | Correct | Domain: `N`. Default on since 2511. The repo emits it only when the box is cleared. |
| `ADDONS=ZoomVDIPlugin` (LTSR) | package-citrixworkspace-cr.ps1:199 (LTSR files: 240) | 2607 LTSR page, "ADDONS" | Correct | Keys: `ZoomVDIPlugin`, `WebexVDIPlugin`. Example: `ADDONS=ZoomVDIPlugin,WebexVDIPlugin`. The 2402 LTSR page has no ADDONS key. |
| `ADDONS=WebexVDIPlugin` | package-citrixworkspace-cr.ps1:202-203 | 2610 page and 2607 LTSR page | Correct | Not installed by default. The repo joins keys with a comma and no space. |
| `InstallEPAClient=N` | package-citrixworkspace-cr.ps1:201 | 2610 page, "Exclude EPA Client" | Correct | Domain: `N`. Valid for a fresh install or an upgrade from a build without EPA. |
| UI label "EPA Client (default on 2508+)" | start-apppackager.ps1:6872 | 2610 page, "Exclude EPA Client" | UNVERIFIED | The vendor page does not state the version that made EPA default. Remove "2508+" or confirm in the Citrix What's New page. |
| `/InstallUberAgent`, `/SkipUberAgentUpgrade` | package-citrixworkspace-cr.ps1:205-207 | 2610 page and 2607 LTSR page, "UberAgent parameters" | Correct | Both streams document it. UberAgent 7.5.1 and later is x64 only. Command line only. Warn for the x86 and ARM64 packagers. |
| `/InstallSRAgent` (Current only) | package-citrixworkspace-cr.ps1:213-216 | 2610 page, "Session Recording add-on support" | Correct | The 2607 LTSR page does not list it. The repo ignores it on LTSR and warns. The x86 installer needs x64 .NET Desktop Runtime 8.0.15 or later (8.x). |
| `AutoUpdateCheck=auto\|manual\|disabled` | package-citrixworkspace-cr.ps1:219-225 | 2610 page, "Auto-update parameters" | Correct | Default `auto`. This parameter must be set for `AutoUpdateStream`, `DeferUpdateCount` and `AURolloutPriority` to work. |
| `EnableCEIP=True\|False` | package-citrixworkspace-cr.ps1:226-228 | 2610 page, "Enable CEIP" | Correct | Default `True`. |
| `EnableTracing=true\|false` | package-citrixworkspace-cr.ps1:229-231 | 2610 page, "Enable always-on tracing" | Correct | Default `true`. Vendor examples use lower case. |
| `ALLOWADDSTORE=S\|A\|N` | package-citrixworkspace-cr.ps1:236-241 | 2610 page, "Store configuration parameters" | Correct | Default `S` (HTTPS only). `A` allows HTTP and HTTPS. `A` does not apply to a per-user install. `N` blocks user stores. |
| `ALLOWSAVEPWD=S\|N\|A` | package-citrixworkspace-cr.ps1:236-241 | 2610 page, "Save the store credentials locally" | Correct | Default `S`. Applies only to stores that use the Citrix Workspace app protocol. |
| `STORE0="name;url;on;description"` | package-citrixworkspace-cr.ps1:244-265 | 2610 page, "Examples of store configuration" | Correct | The repo builds `Name;Url;On;Name` and quotes the value when it has a space. Vendor examples quote the value. The repo rejects `;` and `"` in the name. Add a warning when the StoreFront URL has no `/discovery`: the vendor page says `discovery` is mandatory for pass-through authentication. A Citrix Gateway URL does not need it. |
| `ADDLOCAL=...` (9 keys, opt-in) | package-citrixworkspace-cr.ps1:268-275 | 2610 page, "Install one or more of the specific components" | Correct | Key `SelfService`: the vendor table says `SELFSERVICE`, the vendor examples say `SelfService`. The repo matches the examples. Mandatory keys `ReceiverInside`, `ICA_Client`, `AM` install anyway. No space after a comma (see CTX335534). Omitting `DesktopViewer` can freeze the Connection Center windows. An upgrade ignores exclusions for components already present. |
| ADDLOCAL key `CitrixEnterpriseBrowser` | none (list at package-citrixworkspace-cr.ps1:271) | 2607 LTSR page and 2610 page, ADDLOCAL table | Missing | Syntax: add `CitrixEnterpriseBrowser` to `ADDLOCAL=`. Default: installed when no `ADDLOCAL`. Current 2511 and later do not carry it. LTSR 2607 does. Tooltip: "Installs the Citrix Enterprise Browser. Clear it to leave the browser out of an LTSR install." |
| `/forceinstall` | none | 2610 page, "Common parameters" | Missing | Syntax: `/forceinstall`. No value. Default: off. Cleans old configuration after a failed install or an unsupported-version upgrade. It replaces `rcu`. Tooltip: "Cleans old Citrix data when an install or upgrade fails. Use it as a retry package." |
| `STORE1`, `STORE2`, ... | none | 2610 page, "To configure multiple stores" | Missing | Syntax: `STORE1="name;url;on;description"`. Same format as `STORE0`. Default: none. Tooltip: "Adds a second store. Use it for a backup or a Citrix Gateway store." |
| `INSTALLDIR=path` | none | 2610 page, "Specify custom installation directory" | Missing | Syntax: `INSTALLDIR="C:\Apps\Citrix"`. Domain: folder path. Default: `C:\Program Files\Citrix`. Vendor note: Program Files is protected by the OS, so a custom folder needs the right permissions. Tooltip: "Sets the install folder. Leave empty for the default." |
| `AutoUpdateStream=LTSR\|Current` | none | 2610 page, "Select the version for update" | Missing | Syntax: `AutoUpdateStream=LTSR`. Domain: `LTSR`, `Current`. Default: not stated. Works only when `AutoUpdateCheck` is also set. Tooltip: "Selects the update channel. Use LTSR on devices that stay on a long-term build." |
| `AURolloutPriority=Auto\|Fast\|Medium\|Slow` | none | 2610 page, "Set rollout priority" | Missing | Syntax: `AURolloutPriority=Slow`. Domain: `Auto`, `Fast`, `Medium`, `Slow`. Default: `Auto`. Works only when `AutoUpdateCheck` is also set. Tooltip: "Sets when in the Citrix delivery period a device gets an update." |
| `ALLOW_BIDIRCONTENTREDIRECTION=0\|1` | none | 2610 page, "HDX features parameters" | Missing | Syntax: `ALLOW_BIDIRCONTENTREDIRECTION=1`. Domain: `0`, `1`. Default: `0`. Tooltip: "Allows content redirection between the client and the host in both directions." |
| `FORCE_LAA=0\|1` | none | 2610 page, "Set local app access" | Missing | Syntax: `FORCE_LAA=1`. Domain: `0`, `1`. Default: `0`. Needs an admin install. Tooltip: "Installs the Local App Access component." |
| `ALLOW_CLIENTHOSTEDAPPSURL=0\|1` | none | 2610 page, "Set URL redirection feature" | Missing | Syntax: `ALLOW_CLIENTHOSTEDAPPSURL=1`. Domain: `0`, `1`. Default: `0`. Tooltip: "Turns on URL redirection on the device." |
| `LEGACYFTAICONS=True\|False` | none | 2610 page, "Display icons for documents or files" | Missing | Syntax: `LEGACYFTAICONS=True`. Domain: `True`, `False`. Default: `False`. `True` hides file-type icons of subscribed apps. Tooltip: "Hides file-type icons of subscribed apps." |
| `DisableSetting=0\|1\|2\|3` | none | 2610 page, "Hide Shortcuts and Reconnect option" | Missing | Syntax: `DisableSetting=3`. Domain: 0 show both, 1 show Reconnect only, 2 show Shortcuts only, 3 hide both. Default: `0`. Tooltip: "Hides the Shortcuts and Reconnect options in Advanced Preferences." |
| `CLIENT_NAME`, `ENABLE_DYNAMIC_CLIENT_NAME` | none | 2610 page, "Specify the name to identify the user device" | Missing | Syntax: `ENABLE_DYNAMIC_CLIENT_NAME=No CLIENT_NAME=name`. `ENABLE_DYNAMIC_CLIENT_NAME` domain: `Yes`, `No`. Default `Yes` (client name follows the computer name). `CLIENT_NAME` default: `%COMPUTERNAME%`. Set `CLIENT_NAME` only with `No`. Tooltip: "Sets a fixed client device name. Leave empty to use the computer name." |
| `/SkipSRAgentUpgrade` | none (SR agent at package-citrixworkspace-cr.ps1:213-216) | 2610 page, "Session Recording add-on support" | Missing | Syntax: `/InstallSRAgent /SkipSRAgentUpgrade`. No value. Default: the installer upgrades an existing agent. Current only. Tooltip: "Keeps the Session Recording agent that is already on the device. Use it when another package manages the agent." |
| `/noreboot` | none | 2402 LTSR page only | UNVERIFIED | The 2402 LTSR page lists `noreboot` (no slash) as "suppresses reboot prompts". The 2607 LTSR and 2610 pages do not list it. Do not add it until Citrix documents it for the stream in use. |
| `/InstallEmbeddedBrowser=N` and `InstallEmbeddedBrowser=N` | none | 2402 LTSR page | Not applicable | Documented for 2402 LTSR only. The 2607 LTSR and 2610 pages exclude the browser through the `ADDLOCAL` list. |
| `/DisableAppProtection`, `/EXCLUDE`, `LOGFILE` | none | 2607 LTSR, 2610 pages | Not applicable | The user list names them. Neither current page documents them. Do not emit them. |

Skipped (reason)
- `includeAppProtection`: deprecated in 2212. Replaced by `startAppProtection`.
- `rcu`: deprecated in 1909. Replaced by `/forceinstall`.
- `IncludeCitrixCasting`: Citrix says it cannot be installed in LTSR 2402 and later, even with the switch.
- `DeferUpdateCount`: Citrix says it does not apply from 2207.
- `STARTMENUDIR`, `DESKTOPDIR`: documented. `DESKTOPDIR` also needs the registry key `PutShortcutsOnDesktop=True`, which is not a property. Low value. Skipped.
- `/uninstall` without `/silent`: shows UI.
- `? / help`: prints a help dialog. Not silent.
- Per-user install switch: the Citrix page documents none. A non-admin run makes a per-user install. Do not mix with an `/allusers`-style expectation.
- `ADDLOCAL` keys `MTOPBootStrapperInstaller` and `EPAClient`: listed only in the 2402 LTSR table. Use `InstallEPAClient=N` for EPA.
- `AppProtection` as an `ADDLOCAL` key: Citrix says it is invalid since 2212.

Notes for the Citrix owner
- The Citrix page says to type all parameters on one line. The repo passes an argument array to `Start-Process`. Both are equal.
- A value with a space needs quotes. The repo adds them for `STORE0` only. Add the same rule for `INSTALLDIR` and `CLIENT_NAME` when they ship.
- Every added switch needs a Pester case in `Tests/CitrixWorkspace.Tests.ps1`, which already enforces that the four `ConvertTo-CwaInstallArguments` copies are identical (line 233-237 of that test).

---

## 4. DBeaver Community

Packager: `Packagers/package-dbeaver.ps1`. System scope: `/allusers /S` (line 245). User scope: `/S /currentuser` (line 235). Uninstall script: `Uninstall.exe` with `/allusers /S` or `/currentuser /S` (lines 247-248, 237-238). DisableAI: sets the environment variable `DBEAVER_AI_DISABLED=true` (lines 259-271).

| Option | Where in repo | Vendor source URL | Status | Finding and exact fix |
|---|---|---|---|---|
| `/S` | package-dbeaver.ps1:235, 245 | https://dbeaver.com/docs/dbeaver/Windows-Silent-Install/ | Correct | Silent mode. Case-sensitive. Must be paired with `/allusers` or `/currentuser`. |
| `/allusers` (InstallScope System) | package-dbeaver.ps1:245-248 | https://dbeaver.com/docs/dbeaver/Windows-Silent-Install/ | Correct | Case-insensitive. Needs an administrator. Default folder `C:\Program Files\DBeaver`. |
| `/currentuser` (InstallScope User) | package-dbeaver.ps1:235-238 | https://dbeaver.com/docs/dbeaver/Windows-Silent-Install/ | Correct | Case-insensitive. Installs for the current user only. |
| Uninstaller file name `Uninstall.exe` | package-dbeaver.ps1:241, 250, 283 | https://dbeaver.com/docs/dbeaver/Windows-Silent-Install/ | Correct | The vendor page says `Uninstaller.exe`. The installer of the current release writes `Uninstall.exe` . The vendor page is wrong. Keep the repo name. |
| Uninstall args `/allusers /S` and `/currentuser /S` in the generated script | package-dbeaver.ps1:248, 238 | https://dbeaver.com/docs/dbeaver/Windows-Silent-Install/ | Correct | The vendor page says the three switches also apply to the uninstaller. |
| Manifest `UninstallArgs` for System scope = `/S` | package-dbeaver.ps1:247, 315 | https://dbeaver.com/docs/dbeaver/Windows-Silent-Install/ | Wrong | `/S` needs `/allusers` or `/currentuser`. The script file at line 248 is right, but the manifest value is not. Fix: `$uninstallArgs = "/allusers /S"` at line 247. |
| UI label and tooltip for DisableAI describe `-Dai.disabled=true` in dbeaver.ini | start-apppackager.ps1:6688 | https://dbeaver.com/docs/dbeaver/Disable-AI-assistance/ | Wrong | The packager sets `DBEAVER_AI_DISABLED=true` (package-dbeaver.ps1:259-271), not an ini line. The tooltip says "Appends ... to the installed dbeaver.ini ... re-applied after every reinstall". Fix the label to `Disable AI features (DBEAVER_AI_DISABLED)` and the tooltip to: "Sets the environment variable DBEAVER_AI_DISABLED=true after install. Machine scope for a System install, user scope for a User install. A user cannot turn AI on again in Preferences." Also fix the comment and the .SYNOPSIS text at package-dbeaver.ps1:63 (`Appends -Dai.disabled=true ...`). |
| `DBEAVER_AI_DISABLED=true` env var, Machine scope | package-dbeaver.ps1:265-268 | https://dbeaver.com/docs/dbeaver/Disable-AI-assistance/ | Correct | Vendor text: "Use the system variable by setting `DBEAVER_AI_DISABLED` to `true`". A running DBeaver needs a restart to read it. |
| `DBEAVER_AI_DISABLED` at User scope for a User install | package-dbeaver.ps1:265 | https://dbeaver.com/docs/dbeaver/Disable-AI-assistance/ | UNVERIFIED | The vendor page says "system variable" and states no scope. A user variable is read by the user process, so it should work. No vendor text confirms it. Test on the lab with a User install. |
| Edition limit of the AI switch for Community | package-dbeaver.ps1:259 | https://dbeaver.com/docs/dbeaver/Disable-AI-assistance/ | UNVERIFIED | The page is in the Community doc tree and states no edition limit. An older DBeaver wiki says the permanent disable is for paid editions. Check on the lab that a CE build honors the variable. |
| `-Dai.disabled=true` in dbeaver.ini (alternative) | not used | https://dbeaver.com/docs/dbeaver/Disable-AI-assistance/ | Not applicable | Vendor alternative. The repo notes that the installer replaces dbeaver.ini at every run, so the variable is the better choice. Keep. |
| `/D=path` | none | https://dbeaver.com/docs/dbeaver/Windows-Silent-Install/ | Missing | Syntax: `/D=C:\Apps\DBeaver`. Domain: folder path. Default: `C:\Program Files\DBeaver` (all users) or `%LOCALAPPDATA%\DBeaver` (current user). Rules: must be the last parameter, case-sensitive, no quotes. The detection path and the uninstall path in the packager are fixed. A custom folder needs both changed. Tooltip: "Sets the install folder. Leave empty for the default. The detection rule follows this folder." |

Skipped (reason)
- Exit codes `666660` (bad parameter), `666661`, `666662`, `666663`, `666666` (cannot start elevated instance): documented. They are results, not switches. The wrapper passes the code through. No change needed.
- `/uninstall`, component selection, `/NOSHORTCUT`, desktop shortcut options: the vendor page documents none.
- Run without `/allusers` or `/currentuser`: the installer prompts. Not silent.
- `/allusers` for a SYSTEM install through ConfigMgr: works because SYSTEM is an administrator. No change.

---


## 5. Microsoft 365 Apps / Office Deployment Tool

Primary source (A): https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/office-deployment-tool-configuration-options
Other sources: (B) https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/overview-office-deployment-tool, (C) https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/upgrade-from-msi-version, (D) https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/change-bitness, (E) https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/overview-deploying-languages-microsoft-365-apps, (F) https://learn.microsoft.com/en-us/microsoft-365-apps/updates/manage-microsoft-365-apps-updates-configuration-manager, (G) https://learn.microsoft.com/en-us/microsoft-365-apps/licensing-activation/overview-shared-computer-activation, (H) https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/microsoft-search-bing, (I) https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/deploy-microsoft-365-apps-remote-desktop-services, (J) https://learn.microsoft.com/en-us/microsoft-365-apps/licensing-activation/overview-licensing-activation-microsoft-365-apps

### 1a. What the code emits

| Option | Where in repo | Vendor source URL | Status | Finding and exact fix |
|---|---|---|---|---|
| `Add OfficeClientEdition="32"/"64"` | psm1:5180; m365x64:108 (`$Architecture = "64"`) | A (OfficeClientEdition) | Correct | Values 32 and 64 are the only documented values. |
| `Add Channel="MonthlyEnterprise"/"Current"` | psm1:5167-5168, 5181; m365x64:78-80, 99-106 | A (Channel attribute) | Correct | Both values are documented. Learn says the default for a new install is `Current`, so the explicit attribute is right. Learn also lists BetaChannel, CurrentPreview, SemiAnnualPreview, SemiAnnual (see Skipped). |
| Channel GUID to CDN URL map (`55336b82-...`, `492350f6-...`) | m365x64:99-102, 109 | none found | UNVERIFIED | No Learn page read in this audit lists the CDN GUIDs. Verify against the Office CDN channel table before release. |
| `Add Version="16.0.x"` (omitted when `Latest`) | psm1:5183-5185 | A (Version attribute) | Correct | Optional. Learn says to set `Channel` together with `Version`. The code does. |
| `Add SourcePath="..."` | psm1:5177-5179; m365x64:269 (download), none on install | A (SourcePath) | Correct | Omitted on install, so ODT looks in the folder of setup.exe. This matches the documented default. |
| `Add OfficeMgmtCOM="TRUE"` | psm1:5182 | A (OfficeMgmtCOM); F | Correct | Documented values are `True`/`False`. Value is hard-coded. Learn: with `OfficeMgmtCOM` and `Updates Enabled` both true, only ConfigMgr delivers updates. The comment at m365x64:242 ("WSUS/CDN handles updates") is right only for the ConfigMgr path. Case (`TRUE` vs `True`) is not stated; the repo test expects `TRUE`. See Missing M2 for a toggle. |
| `Add MigrateArch="TRUE"` | psm1:5186 | A (MigrateArch); D | Wrong | Learn default is False. With TRUE and a different `OfficeClientEdition`, ODT changes the installed bitness. D shows `MigrateArch='TRUE'` with `OfficeClientEdition="32"` reverts 64-bit to 32-bit. The x86 packages (`package-m365apps-x86.ps1`, `...project-x86`, `...visio-x86`) therefore downgrade existing x64 Office. Fix: emit `MigrateArch` only when an explicit "Migrate architecture" option is on; default off (omit the attribute). |
| `Product ID="O365ProPlusRetail"` etc. | psm1:5192; m365x64:107 | A (Product ID) | Correct | IDs come from the Product IDs list that A links. |
| `Language ID="en-us"` (fixed) | psm1:5193 | A (Language); E | Correct | Syntax is right. The language is fixed. See Missing M1. |
| `ExcludeApp ID=...` list (Access, Excel, Groove, Lync, OneDrive, OneNote, Outlook, OutlookForWindows, PowerPoint, Publisher, Teams, Word) | psm1:5194-5197; start:6606-6618 | A (ExcludeApp) | Correct | These 12 IDs match the documented allowed values exactly. `Groove` = OneDrive and `Lync` = Skype for Business is documented. |
| `ExcludeApp ID="Bing"` | psm1:5173 (default), start:234, 6619, 428 | A (list has no Bing); H | Wrong | `Bing` is not in the documented ID list. H says the Microsoft Search in Bing background service is not available for new deployments of Version 2504 and later, and the product retired 2025-03-31. The ID has no documented effect. Fix: remove `Bing` from the default list at psm1:5173 and start:234, remove the checkbox at start:6619, and drop it from `$validExcludes` at start:425-428. |
| `ExcludeApp` applied to Visio and Project products | psm1:5191-5198 (loop runs for every `ProductIds` entry); `package-m365visio-*.ps1:225`, `package-m365project-*.ps1:225` | A (ExcludeApp ID list) | UNVERIFIED | The ID list is for Microsoft 365 Apps products. Learn names OneDrive for Visio. No page states that `Word`, `Access` and similar IDs are accepted or ignored under `VisioProRetail` or `ProjectProRetail`. Test one Visio package. Safe fix: write `ExcludeApp` only under the O365 suite product, and `OneDrive`/`Groove` only under Visio and Project. |
| `Property SharedComputerLicensing Value="1"` | psm1:5203 | A (SharedComputerLicensing); G; I; J | Wrong | Default is 0. Learn: set 1 only for shared computers, for example RDS. J and G: user-based licensing is the default for PCs that one user owns. The code forces 1 on every package. Fix: emit it only when a "Shared computer activation" option is on; default off (omit the line). See Missing M3. |
| `Property FORCEAPPSHUTDOWN Value="TRUE"` | psm1:5204 | A (FORCEAPPSHUTDOWN); C | Correct | Values TRUE/FALSE, default FALSE. TRUE closes blocking apps and can lose data. C says RemoveMSI can fail when apps are open and Display Level is None, so TRUE is a valid ConfigMgr choice. It is hard-coded. See Missing M5. |
| `Property DeviceBasedLicensing Value="0"` | psm1:5205 | A (DeviceBasedLicensing) | Not applicable | `0` is the default. The line has no effect. Fix: emit the line only when the option is `1` (see Missing M4). |
| `Property PinIconsToTaskbar Value="FALSE"` | psm1:5206 | A (PinIconsToTaskBar) | Not applicable | Learn: this property applies to Windows 7 and 8.1 only, and the icons are not pinned when the install runs as SYSTEM (ConfigMgr). Both OS versions are out of support. Fix: remove the line. |
| `Property SCLCacheOverride Value="0"` | psm1:5207 | A (SCLCacheOverride) | Not applicable | `0` is the default. It matters only with shared computer activation. Fix: remove the line, or emit it with Missing M3. |
| `RemoveMSI` | psm1:5208 | A (RemoveMSI); C | Correct | Learn recommends it. It removes MSI Office 2007 to 2016, Visio and Project. It does not remove Click-to-Run versions. See Missing M7. |
| `AppSettings / Setup Name="Company"` | psm1:5209-5213 | A (AppSettings shows only `<User .../>`); OCT overview https://learn.microsoft.com/en-us/microsoft-365-apps/admin-center/overview-office-customization-tool | UNVERIFIED | A and B document `AppSettings` with `<User Key= Name= Value= Type= App= Id= />` only. The OCT page says the organization name is set in the General section, but no page shows the `<Setup Name="Company" .../>` XML. Compare with an export from https://config.office.com before release. |
| `Display Level="None" AcceptEULA="TRUE"` | psm1:5214; m365x64:317 | A (Display) | Correct | Level None or Full (default Full). AcceptEULA TRUE or FALSE (default FALSE). Both are right for ConfigMgr. |
| `Logging Level="Standard" Path="%programdata%\Appdeploy\Office2016"` | psm1:5215 | A and B (no Logging element on either page) | UNVERIFIED | The element is not in the current Learn ODT pages. B says logs go to `%temp%`. Third-party pages show `Level="Standard"` and `Path=` as the sample-file form. Treat as accepted but undocumented. The path text `Office2016` is stale. |
| `Updates` element | not emitted | A (Updates); F | Correct | Default Enabled is TRUE, so omission is safe. F recommends an explicit `Updates Enabled="True"` with `OfficeMgmtCOM`. See Missing M8. |
| `/configure install.xml`, `/download download.xml` | m365x64:276, 327, 334 | B (modes) | Correct | `/download` and `/configure` are the documented modes. |
| `Remove` uninstall XML (`Product ID`, Display None) | m365x64:312-319 | A (Remove) | Correct | Documented form. `<Remove All="TRUE"/>` also exists (default FALSE) and removes Visio and Project too, so the per-product form is the safer choice. |
| Deploy mode Managed / Online | m365x64:80-81, 207-253 | B (download and install steps) | Correct | Managed = `/download` with `SourcePath` and pinned `Version`. Online = no `SourcePath`, no `Version`, CDN install. Both are documented patterns. |
| Company name -> AppSettings only | start:6565 | see AppSettings row | UNVERIFIED | Tooltip says "embedded in Office deployment XML and other packager configs". Counted under the AppSettings row. |

### 1b. Missing options (documented, silent-safe, useful for ConfigMgr)

| Option | Where in repo | Vendor source URL | Status | Finding and exact fix |
|---|---|---|---|---|
| M1 Languages: list, `MatchOS`, `Fallback`, `AllowCdnFallback` | psm1:5193 (fixed `en-us`) | A (Language, Fallback, AllowCdnFallback); E | Missing | Syntax: `<Add ... AllowCdnFallback="TRUE">` and `<Language ID="en-us" />` then `<Language ID="MatchOS" Fallback="en-us" />`. Domain: culture codes from E (for example `de-de`, `ja-jp`), plus `MatchOS`; `AllowCdnFallback` True/False, default False. The first language sets the shell UI culture. Download XML needs explicit IDs: E says `MatchInstalled` cannot be used with `/download`, and A says `MatchOS` needs the language pack in the local source or a CDN fallback. Tooltip: "Installs these Office display languages. The first language sets the shell UI language." |
| M2 Update source: `OfficeMgmtCOM` toggle | psm1:5182 | A (OfficeMgmtCOM); F | Missing | Syntax: `OfficeMgmtCOM="True"` or `"False"` on `Add`. Default False in ODT; the repo hard-codes True. Domain: True/False. Tooltip: "True lets ConfigMgr software updates patch Office. False lets Office update from the CDN." |
| M3 Shared computer activation, `SCLCacheOverride`, `SCLCacheOverrideDirectory` | psm1:5203, 5207 | A; G; I | Missing | Syntax: `<Property Name="SharedComputerLicensing" Value="1" />`. For a roaming token add `<Property Name="SCLCacheOverride" Value="1" />` and `<Property Name="SCLCacheOverrideDirectory" Value="\\server\share\%username%" />`. Domain: 0 or 1; default 0. Needs a plan that supports it (G: Microsoft 365 Apps for enterprise plans, or Business Premium). Tooltip: "Turns on shared computer activation for RDS and VDI hosts. Leave it off for normal PCs." |
| M4 Device-based licensing | psm1:5205 | A (DeviceBasedLicensing); J | Missing | Syntax: `<Property Name="DeviceBasedLicensing" Value="1" />`. Domain 0 or 1; default 0. J: the device must be synced to Microsoft Entra ID. Tooltip: "Licenses Office to the device. Use it for shared devices where users do not sign in with Entra accounts." |
| M5 `FORCEAPPSHUTDOWN` as a choice | psm1:5204 | A (FORCEAPPSHUTDOWN); C | Missing | Syntax: `<Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />`. Domain TRUE/FALSE; default FALSE. Tooltip: "TRUE closes Office apps that block setup. Users can lose unsaved data." |
| M6 `MigrateArch` as a choice | psm1:5186 | A (MigrateArch); D | Missing | Syntax: `MigrateArch="TRUE"` on `Add`. Domain TRUE/FALSE; default FALSE. D: needs Version 1902 or later installed, the target bitness files in the source, and `AllowCdnFallback="TRUE"` is advised. Tooltip: "Changes an installed Office to the architecture of this package. Use it for 32-bit to 64-bit upgrades only." |
| M7 `RemoveMSI` options: toggle, `IgnoreProduct`, `MatchPreviousMSI`, `MSICondition` | psm1:5208 | A (RemoveMSI, MSICondition); C | Missing | Syntax keep one product: `<RemoveMSI><IgnoreProduct ID="VisPro" /></RemoveMSI>`. IDs in C: `PrjStd`, `PrjPro`, `VisStd`, `VisPro`, `PrjStdR`, `PrjProR`, `VisStdR`, `VisProR`. Keep old languages: `<Language ID="MatchPreviousMSI" />` (needs `RemoveMSI` and one explicit language). Install only if an MSI product exists: `<Product ID="ProjectProRetail" MSICondition="PrjStd,PrjStdR">`. Tooltip: "Removes MSI Office before the install. IgnoreProduct keeps the listed MSI product." |
| M8 `Updates` element | not emitted | A (Updates: Enabled, UpdatePath, TargetVersion, Deadline, Channel); F | Missing | Syntax: `<Updates Enabled="TRUE" />`. Domain TRUE/FALSE; default TRUE. F recommends TRUE with `OfficeMgmtCOM`. `UpdatePath` (network or HTTP path, default CDN) and `Channel` (same values as `Add`, default Current) are useful for the Online mode without ConfigMgr updates. Tooltip: "TRUE lets Office check for updates. ConfigMgr stays the update source when OfficeMgmtCOM is True." |
| M9 `Version="MatchInstalled"` for Visio and Project add-on packages | psm1:5183 | A (Version) | Missing | Syntax: `<Add Version="MatchInstalled" Channel="MonthlyEnterprise" OfficeClientEdition="64">`. `Channel` is required with it. A: it adds Visio or Project at the version already installed, with no forced update. Also `<Language ID="MatchInstalled" TargetProduct="O365ProPlusRetail" />` (not valid with `/download`, per E). Tooltip: "Installs Visio or Project at the Office version already on the device." |
| M10 `AppSettings` user preferences | psm1:5209-5213 | A and B (AppSettings) | Missing | Syntax (from A): `<AppSettings><User Key="software\microsoft\office\16.0\excel\security" Name="vbawarnings" Value="3" Type="REG_DWORD" App="excel16" Id="L_VBAWarningsPolicy" /></AppSettings>`. B: preferences apply to all current and new users of the device. Build values with https://config.office.com. Tooltip: "Writes Office default settings for every user on the device." |
| M11 `TenantAssociationKey` | none | A (TenantAssociationKey) | Missing | Syntax: `<Property Name="TenantAssociationKey" Value="<TAK>" />`. Needs ODT 16.0.19929.20062 or later. Value: the key from Setup in the Microsoft 365 Apps admin center (alphanumeric, over 300 characters). Registers the device in Inventory. Tooltip: "Registers the device in Microsoft 365 Apps admin center Inventory." Treat the key as a secret; do not log it. |

### 1c. Skipped (reason)

- `Display Level="Full"`: shows UI. Not valid for a required ConfigMgr deployment run as SYSTEM.
- `AUTOACTIVATE`: A says do not set it for Microsoft 365 Apps.
- `Add Channel` values BetaChannel, CurrentPreview, SemiAnnualPreview: not supported or test-only (A). `SemiAnnual`: the ConfigMgr page F announces that Semi-Annual Enterprise Channel merges with Monthly Enterprise Channel from July 2026.
- `PerpetualVL2019`, `PerpetualVL2021`, `PerpetualVL2024`: apply to LTSC and volume products, not to the retail products this tool packages.
- `Add DownloadPath`: needs a second source share and a pinned version. The packager downloads from the CDN directly.
- `PACKAGEGUID`: App-V only.
- `Remove All="TRUE"`: removes Visio and Project too. Unsafe as a default.
- `PinIconsToTaskbar`: see 1a.

---

## 6. SQL Server Management Studio 22

Primary source (S1): https://learn.microsoft.com/en-us/ssms/install/command-line-parameters
Other sources: (S2) https://learn.microsoft.com/en-us/visualstudio/install/use-command-line-parameters-to-install-visual-studio, (S3) https://learn.microsoft.com/en-us/ssms/install/workload-component-ids, (S4) https://learn.microsoft.com/en-us/ssms/install/create-offline, (S5) https://learn.microsoft.com/en-us/ssms/install/update

Code: `Packagers\package-ssms.ps1` (`Get-SsmsInstallArgs` lines 150-169, `Get-SsmsUninstallArgs` 171-179). UI: `start-apppackager.ps1:6652-6672` and defaults at 235-243, 431-442, 7087-7108.

### 2a. What the code emits

| Option | Where in repo | Vendor source URL | Status | Finding and exact fix |
|---|---|---|---|---|
| Bootstrapper URL `https://aka.ms/ssms/22/release/vs_SSMS.exe` | package-ssms.ps1:63 | S1 (bootstrapper table) | Correct | Same URL as the vendor table (Release channel). |
| `--quiet` (UIMode Quiet) | package-ssms.ps1:154; start:6652-6658 | S1 | Correct | Documented. No UI. Standard users cannot use it. |
| `--passive` (UIMode Passive) | package-ssms.ps1:154 | S1 | Correct | S1: `--passive` overrides `--quiet`, so the code emits one or the other. Correct. Passive shows UI and is not suited to a SYSTEM install; the tooltip says so. |
| `--norestart` | package-ssms.ps1:155 | S1 | Correct | S1: must pair with `--quiet` or `--passive`. It does. Setup can return 3010 or 1641 (see note). |
| `--downloadThenInstall` | package-ssms.ps1:156; start:6667 | S1; S2 | Correct | Documented and mutually exclusive with `--installWhileDownloading` (the default). The code never emits both. |
| `--noUpdateInstaller` | package-ssms.ps1:157; start:6668 | S1; S2 | UNVERIFIED | S1: "prevents the installer from updating itself when `quiet` is specified" and the command fails with a nonzero exit code if an installer update is needed. The code allows it with `--passive`; the docs do not say what happens then. Fix: show the option only for UIMode Quiet, and warn that setup can fail with it. |
| `--includeRecommended` | package-ssms.ps1:158; start:6669 | S1 (parameter row) | Wrong | S1: it adds recommended components "for any workloads that are installed", and the workloads come from `--add` or `--allWorkloads`. The code never emits `--add`, so the flag has no workload to act on. It also sets the "add recommended components on update" setting. Fix: emit it only together with at least one `--add` (see Missing S-1). Change the tooltip text "for selected SSMS workloads" to name the control that selects them. |
| `--includeOptional` | package-ssms.ps1:159; start:6670 | S1 | Wrong | Same finding as `--includeRecommended`. Counted once. |
| `--removeOos true` | package-ssms.ps1:160; start:6671 | S1 | Correct | Syntax `--removeOos true` (two tokens) is correct. It applies to one operation. For a persistent setting S1 gives `modifySettings --removeOos true`. |
| `--force` | package-ssms.ps1:161, 177; start:6672 | S1 | Correct | Documented: closes SSMS even in use, can lose work. Without it, setup returns 1003 (SSMS in use) or 8006 (SSMS processes running). |
| `--installPath <dir>` | package-ssms.ps1:163-166; start:6665 | S1 | Correct | Optional on install, required for `modify`, `update`, `uninstall`. |
| Default install path `C:\Program Files\Microsoft SQL Server Management Studio 22\Release` | package-ssms.ps1:70 | S1 (modifySettings example) | Correct | The vendor example uses this exact path. The doc does not state it as a default in prose. Detection file `Common7\IDE\Ssms.exe` (line 273) is UNVERIFIED on Learn pages read here, but is not part of the option audit. |
| `uninstall --installPath <dir> --quiet --norestart [--force]` | package-ssms.ps1:174-178 | S1 (commands table) | Correct | `uninstall` is a documented command and goes first. The wrapper runs `%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\setup.exe` (lines 224-226). S1: start that installer from a different folder than its own. The wrapper sets no working directory; ConfigMgr uses the content folder, which is a different folder. |

### 2b. Missing options

| Option | Where in repo | Vendor source URL | Status | Finding and exact fix |
|---|---|---|---|---|
| S-1 Workloads and components | not offered | S1 (`--add`, `--allWorkloads`, `--all`); S3 | Missing | Syntax: repeat the switch per item, `--add Microsoft.SqlServer.Workload.SSMS.BI --add Microsoft.SqlServer.Workload.SSMS.AI`. Append `;includeRecommended` or `;includeOptional` to one ID for per-workload control. Workload IDs in S3: `Microsoft.SqlServer.Workload.SSMS.AI`, `.BI`, `.CodeTools`, `.DevOps`, `.HybridAndMigration`. Default: the core editor only. Tooltip: "Adds SSMS workloads. Each workload is one --add switch. Without it, setup installs the core editor only." |
| S-2 Product languages | not offered | S1 (`--addProductLang`, language list) | Missing | Syntax: repeat `--addProductLang de-de --addProductLang fr-fr`. Domain: `cs-cz`, `de-de`, `en-us`, `es-es`, `fr-fr`, `it-it`, `ja-jp`, `ko-kr`, `pl-pl`, `pt-br`, `ru-ru`, `tr-tr`, `zh-cn`, `zh-tw`. Default: the language of the machine locale. Tooltip: "Installs these SSMS display languages. Default is the OS language." |
| S-3 Offline layout install | not offered (package-ssms.ps1 file header says "unless using a layout in a future packager") | S1 (`--layout`, `--noWeb`); S4 | Missing | Stage: `vs_SSMS.exe --layout C:\SSMS_Layout --add <id>` (without `--add` it downloads everything). Install from the distribution point content: `<layout>\vs_SSMS.exe --noWeb --quiet --norestart --add <id>`. S1: `--noWeb` stops package downloads; an internet-connected client can still check for updates. S4: install the certificates first, keep the path under 80 characters, and install the Visual C++ v14 redistributable. S4: future updates expect the layout at the same path. Tooltip: "Installs SSMS from content on the distribution point. Clients need no internet to get packages." |
| S-4 Turn off in-product updates | not offered | S1 (`modifySettings`, `--newChannelUri`) | Missing | Syntax after install: `"%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\setup.exe" modifySettings --installPath "<dir>" --newChannelUri C:\doesNotExist.chman --quiet`. S1 states a nonexistent file path stops updates. Default: SSMS checks the Microsoft channel. Pair with the packager's update cadence. Tooltip: "Stops SSMS from updating itself. ConfigMgr supersedence delivers new versions." |
| S-5 Package cache | not offered | S1 (`--cache`, `--nocache`) | Missing | Syntax: `--nocache`. S1: deletes packages after install; default policy is to cache them for repair. Domain: `--cache` or `--nocache`. Tooltip: "Deletes the setup package cache after install. This saves disk space. Repair downloads the packages again." |

### 2c. Notes

- Exit codes from S1 for ConfigMgr return codes: 0 success; 3010 success, restart needed; 1641 success, restart started; 1003 SSMS in use; 8006 SSMS processes running; 1618 another install running; 5007 computer does not meet requirements; 8010 OS not supported. Repo return-code mapping is outside this audit.
- Persistent out-of-support removal: `modifySettings --removeOos true` (S1). Not offered.

### 2d. Skipped (reason)

- `--productId`, `--channelId`, `--nickname`: select or label an instance. A single SSMS 22 Release install needs none. `--productKey` is not in the SSMS page (S1); it is a Visual Studio parameter (S2).
- `--config <vsconfig>`: additive list of workloads. `--add` (S-1) gives the same result with less file handling.
- `--all`, `--allWorkloads`: large install; `--add` is more controlled.
- `--wait`: S1 says it works for layout operations on the bootstrapper. The wrapper already waits (`Start-Process -Wait`).
- `--installerOnly`, `update`, `updateall`, `repair`, `modify`, `export`, `removeChannel`: not install options.
- `--installChannelUri`, `--installCatalogUri`, `--keepLayoutVersion`, `--locale`, `--path cache|shared`: layout or per-machine tuning with no clear ConfigMgr gain; `--path shared` is settable once only (S1).
- `--passive` for production: shows UI.

---

## 7. TeamViewer Host

Vendor sources:
- (T1) Deploy TeamViewer Host or full client, 9/10: https://www.teamviewer.com/en/global/support/knowledge-base/teamviewer-remote/deployment/mass-deployment-user-guide/deploy-teamviewer-host-or-full-client-9-10/
- (T2) Deploy TeamViewer Host or full client, 9/9: https://www.teamviewer.com/en/global/support/knowledge-base/teamviewer-tensor-classic/deployment/deploy-teamviewer-host-or-full-client-9-9/
- (T3) Assignment options (legacy): https://www.teamviewer.com/en/global/support/knowledge-base/teamviewer-classic/deployment/mass-deployment-on-windows-user-guide-legacy/assignment-options-5-6-legacy/
- (T4) Recommended scripts (legacy): https://www.teamviewer.com/en/global/support/knowledge-base/teamviewer-classic/deployment/mass-deployment-on-windows-user-guide-legacy/recommended-scripts-3-6-legacy/
- (T5) Create your module: https://support.teamviewer.com/en/support/solutions/articles/75000125490-create-your-module-3-10

All vendor pages cover the MSI. No official TeamViewer page read in this audit documents the `TeamViewer_Host_Setup_x64.exe` installer or an `/S` switch. Third-party deployment guides (ManageEngine, silentinstall.info) list `/S` for the Host EXE; those are not vendor sources.

Code: `Packagers\package-teamviewerhost.ps1` (EXE install, lines 119-123, 265-280, 303-304); `start-apppackager.ps1` (`Read-TvHostConfig` 3774-3798, `Save-TvHostConfig` 3800-3805, UI 6805-6826, 7172-7175).

Key fact: the packager installs the EXE with `/S` only. It never reads the four UI values. It logs warnings and ignores them (package-teamviewerhost.ps1:165-195). The four settings in the UI change nothing in the install command.

### 3a. What the code emits or stores

| Option | Where in repo | Vendor source URL | Status | Finding and exact fix |
|---|---|---|---|---|
| EXE installer `TeamViewer_Host_Setup_x64.exe /S` | package-teamviewerhost.ps1:122-123, 268, 303 | none (vendor docs are MSI only) | UNVERIFIED | No TeamViewer page read here documents `/S` for the EXE. Only third-party pages do. T1 and T2 document `msiexec /i ... /qn` only. Test the EXE on the lab and record the exit code. |
| Claim "MSI URL returns 404 since April 2026; MSI only for authenticated Management Console users" | package-teamviewerhost.ps1:31-36 | T1 (MSI download is not described; the page links a "Download the .msi installer" article); T5 | UNVERIFIED | Not checked: the 404 itself. T5 and a TeamViewer community post say the custom Host and its download link come from the Management Console, which agrees with the claim. Check the URL by hand before relying on the comment. |
| Uninstall `%ProgramFiles%\TeamViewer\uninstall.exe /S` | package-teamviewerhost.ps1:276-279, 304 | none | UNVERIFIED | The path and `/S` are not in a vendor page. Third-party lists use `TeamViewer_Host_Setup_x64.exe /S` for removal. Check on the lab. With an MSI install, use `msiexec /x {ProductCode} /qn`. |
| Detection `%ProgramFiles%\TeamViewer\TeamViewer.exe` version >= staged | package-teamviewerhost.ps1:290, 306-314 | none | UNVERIFIED | Install folder not stated in the vendor pages read. Confirm on the lab with the staged build. |
| `ApiToken` -> `APITOKEN=<token>` | start:3776, 3790, 6805-6810, 7172; package-teamviewerhost.ps1:183 | T3, T4 (legacy MSI pages) | Wrong | Property name is correct for the MSI and documented as a legacy method. T3 limits it to TeamViewer 14 or newer, Corporate and Tensor licenses. Current pages (T1, T2) use `ASSIGNMENTID` and do not list `APITOKEN`. In this repo the value is never passed to any installer, so the field is inert. Fix: either pass it on an MSI command line, or mark the field "MSI only, not used" in the UI. The token is stored as plain text in `Packagers\teamviewer-host-config.json` (start:3771, 3803-3804); keep that file out of git and out of logs. T3 and T4 do not describe token permissions. |
| `CustomConfigId` -> `CUSTOMCONFIGID=<id>` | start:3777, 6812-6817 | T1, T2, T4 | Correct (name) | Name correct. T1: "only applies to the TeamViewer Host". Ignored by the EXE path (package-teamviewerhost.ps1:184). Tooltip text is fine. |
| `AssignmentOptions` -> `ASSIGNMENTOPTIONS="..."` | start:3778, 6819-6824 | T3, T4 | Correct (name and flags) | T3 documents `--alias [<name>]`, `--grant-easy-access`, `--reassign`, `--group <name>`, `--group-id <gNNN>`, `--proxy`, `--proxy-user`, `--proxy-pw`, `--proxy-pw-base64`. The tooltip flags (`--grant-easy-access`, `--alias %COMPUTERNAME%`, `--reassign`, `--group <name>`) match. T3: `--reassign` resets an existing assignment. Inert in this repo (EXE). Legacy method; the current method is `ASSIGNMENTID` (Missing T-1). |
| `RemoveDesktopShortcut` -> `REMOVE=f.DesktopShortcut` | start:3779, 6826, 7175; package-teamviewerhost.ps1:186 | T1, T2 | Wrong | T1 and T2 give the full command as `msiexec /i "...msi" /qn ADDLOCAL=ALL REMOVE=f.DesktopShortcut`. The tooltip (start:6826) says it only adds `REMOVE=f.DesktopShortcut` to msiexec, and says the install "does not place" a shortcut. Fix: tooltip and any future command line must include `ADDLOCAL=ALL`. To clear a shortcut left by an older `DESKTOPSHORTCUTS=0` install, T1 gives `msiexec /i "...msi" /qn REMOVE=f.DesktopShortcut`. The UI default (`$true`, start:3779) has no effect on the EXE. |
| Tooltip says options pass "to msiexec" | start:6824, 6826 | T1, T2 | Wrong | The packager runs no msiexec (package-teamviewerhost.ps1:268). Counted under the rows above; fix the tooltip text together with them. |

### 3b. Missing options (MSI deployments, per vendor docs)

These options work only when the packager uses the MSI. If the EXE stays, mark them "MSI only".

| Option | Where in repo | Vendor source URL | Status | Finding and exact fix |
|---|---|---|---|---|
| T-1 `ASSIGNMENTID` | not offered | T1, T2 | Missing | Syntax: `msiexec /i "TeamViewer_Host.msi" /qn CUSTOMCONFIGID=<id> ASSIGNMENTID=<id>`. Value: an assignment ID from a valid Management Console configuration (T1: invalid ID gives error 400). Default: none. This is the current assignment method. Tooltip: "Assigns the Host to your TeamViewer account at install. Use the ID from the Management Console." |
| T-2 Post-install assignment command | not offered | T1, T2 | Missing | Syntax: `TeamViewer.exe assignment --id <ASSIGNMENTID> --retries=3 --timeout=120 [--device-alias <name>]`. `--retries`: retry on temporary errors, 1 second apart. `--timeout`: seconds for all tries. `--device-alias` with no value uses the TeamViewer ID; without the parameter it uses `%computername%` (T2). Needs local administrator rights (T1 error 407) and TeamViewer installed first (error 3). Tooltip: "Runs account assignment after install. Use it when the MSI cannot assign during setup." |
| T-3 `SETTINGSFILE` | not offered | T1, T2, T4 | Missing | Syntax: `SETTINGSFILE="C:\path\settings.tvopt"`. Value: absolute local path or network share to an exported `.tvopt` file. T1: do not combine with `IMPORTREGFILE`. Tooltip: "Imports a TeamViewer options file at install. Use an absolute path." |
| T-4 `INSTALLVPN` | not offered | T1 | Missing | Syntax: `INSTALLVPN=1`. T1 shows only the value `1`; the default is not stated. Tooltip: "Installs the TeamViewer VPN driver." Include only if the site uses TeamViewer VPN. |
| T-5 `CLASSICUI` | not offered | T1 | Missing | Syntax: `CLASSICUI=1`. T1 shows only the value `1`. Tooltip: "Installs the Host with the classic user interface." Use it only if the module needs the old UI. |

### 3c. Skipped (reason)

- `ENABLEOUTLOOKPLUGIN=true|false` (default true) and `INSTALLSECURITYKEYREDIRECTION=0|1` (default 0): T1 lists them for the general installer. This audit found no statement that they apply to the Host. UNVERIFIED for Host; skip.
- `--proxy-pw` inside `ASSIGNMENTOPTIONS` (T3): puts a password on the command line and in ConfigMgr logs. `--proxy` and `--proxy-user` are acceptable.
- `--monitoring-policyid`, `--patchmanagement-policyid` (T3): work only with `TeamViewer_Full.msi`, not the Host.
- `IMPORTREGFILE`: not combinable with `SETTINGSFILE` (T1). Use `SETTINGSFILE`.
- `DESKTOPSHORTCUTS=0`: T1 mentions it only as a past method. Use `REMOVE=f.DesktopShortcut`.
- Rollout configuration with assignment: T1 says a second assignment pop-up can appear. Leave the rollout configuration unset for silent deployment.
