# Options dialog

The Options dialog holds every setting of AppPackager. It has a tree of sections on the left and the section on the right.

## What it does

**OK** saves the changes of every section that you opened, in one step. **Cancel** discards them all. The tree has seven top-level nodes: **Connections**, **Packaging**, **Applications**, **Updates and CVEs**, **Security**, **Tools** and **About**. The window title names the section that is open.

The search box above the tree finds a setting by name. Type a name, for example `sql port`. The list of matches replaces the tree. Press Enter to open the first match, or select another one. The window opens the section and puts the cursor in the field.

![The search box lists the matches](../images/options-search.png)

A value that cannot be stored stops the save and opens its field. A port outside 1 to 65535 is one example.

The Setup window and the Connections sections use the same fields. A value that you type in Setup is the value that Options shows.

## When to use it

Use it after the Setup window, when a site, a share or a destination changes, and to set the options that packagers read when they stage.

## Requirements

None.

## Steps

Select **Options**. The ConfigMgr section opens.

### Connections

#### Systems in use

Select the systems that AppPackager publishes to. The sidebar disables the button of a system that is not selected.

![The Systems in use section](../images/options-systems-in-use.png)

#### ConfigMgr

**Site code**, **Site server / SMS Provider**, the site database (**SQL server**, **SQL instance**, **SQL port**, **Database name**) and **Content share UNC**. Discover connects to `server\instance,port`. An empty database field takes its value from the site definition, and an empty database name means `CM_` and the site code. A name `server\instance,port` in the SQL server field shows split into the three fields. **Console status** shows whether the console is installed on this computer.

![The ConfigMgr section](../images/options-configmgr.png)

#### WSUS

The server, the **Port**, **Use SSL**, the signing certificate, the classification, the approval group and **Decline superseded**. The port is 8530 for HTTP and 8531 for HTTPS, and it follows **Use SSL** until you type a custom port. A port outside 1 to 65535 stops the save. See [Publish to Intune and Publish to WSUS](publish-intune-wsus.md).

![The WSUS section](../images/options-wsus.png)

#### Intune

**Tenant ID**, **Client ID** and **Client secret** of the app registration. The secret is stored with DPAPI for the current Windows user. An empty box keeps the saved secret.

![The Intune section](../images/options-intune.png)

#### Test connection

One button for each system. ConfigMgr reads the SMS Provider and opens the site database. WSUS connects to the server. Intune requests a Graph token. Each test uses the values that are in the fields, before you select OK.

![The Test connection section](../images/options-test-connection.png)

### Packaging

#### Folders

**Download root**, the **Packagers folder** and **Content layout**.

![The Folders section](../images/options-folders.png)

#### Deployment type defaults

The estimated and maximum runtime, **Fallback distribution points** and **Neighbor/default DP** of every deployment type that Publish to ConfigMgr creates.

![The Deployment type defaults section](../images/options-deployment-type-defaults.png)

#### Application title

**Include version** and **Remove version**. Without a selection, a packager uses its own name.

![The Application title section](../images/options-application-title.png)

#### After Package

**Distribute content**, **DP group**, **Test deployment**, **Test collection** and **Create collection**. The test deployment needs Distribute content and a DP group.

![The After Package section](../images/options-after-package.png)

### Applications

#### Adobe Acrobat Reader

The edition, the languages of the multilingual edition, and the Adobe properties `EULA_ACCEPT`, `DISABLEDESKTOPSHORTCUT`, `ENABLE_CHROMEEXT`, `DISABLE_BROWSER_INTEGRATION` and `UPDATE_MODE`.

![The Adobe Acrobat Reader section](../images/options-adobe-acrobat-reader.png)

#### Beyond Compare 5

The **License key file** and the policy switches `/DisableRemoteProfiles` and `/DisableSavedPasswords`.

![The Beyond Compare 5 section](../images/options-beyond-compare-5.png)

#### Citrix Workspace App

The **Preview** button shows the command line of the Current Release and LTSR packagers. Six sub-sections set the switches.

![The Citrix Workspace App section](../images/options-citrix-workspace-app.png)

#### Store

The store name and URL (`STORE0`).

![The Store section](../images/options-store.png)

#### Installation

`/CleanInstall`, single sign-on, App Protection, session pre-launch, self-service mode, `/forceinstall`, content and URL redirection, Local App Access, file-type icons, Advanced Preferences and the client name.

![The Installation section](../images/options-installation.png)

#### Plugins and add-ons

The Teams, Zoom and WebEx VDI plug-ins, uberAgent, the EPA client and the Session Recording agent.

![The Plugins and add-ons section](../images/options-plugins-and-add-ons.png)

#### Update and telemetry

Auto-update, update stream, rollout priority, CEIP and always-on tracing. The update stream and the rollout priority work only when Auto-update is set.

![The Update and telemetry section](../images/options-update-and-telemetry.png)

#### Store policy

`ALLOWADDSTORE` and `ALLOWSAVEPWD`.

![The Store policy section](../images/options-store-policy.png)

#### Components ADDLOCAL

The components of `ADDLOCAL`. With **Customize** cleared, the packager omits `ADDLOCAL`.

![The Components ADDLOCAL section](../images/options-components-addlocal.png)

#### DBeaver Community

The install scope and the environment variable `DBEAVER_AI_DISABLED`.

![The DBeaver Community section](../images/options-dbeaver-community.png)

#### Microsoft 365 Apps

Links to the Office Deployment Tool download and the configuration.xml reference, **Import** and **Use my own configuration.xml**, and the generated settings. See the steps below.

![The Microsoft 365 Apps section](../images/options-microsoft-365-apps.png)

#### SQL Server Management Studio

The UI mode, the install path, the workloads, the languages and the other switches of the SSMS 22 setup.

![The SQL Server Management Studio section](../images/options-sql-server-management-studio.png)

#### TeamViewer Host

The four settings that apply to a TeamViewer Host MSI only. The packager installs the EXE, which ignores them.

![The TeamViewer Host section](../images/options-teamviewer-host.png)

#### Local installer sources

One folder for each packager whose download needs a sign-in. The caption names the packager file.

![The Local installer sources section](../images/options-local-installer-sources.png)

### Updates and CVEs

#### One Click

The action on an update, the default destinations, the policy for an existing ConfigMgr application, and for each application the Track box, its destinations and its cadence. See [One Click](one-click.md).

![The One Click section](../images/options-one-click.png)

#### Product filter

Shows or hides single packagers in the main grid, grouped by vendor. **Not offered by Discover** counts the applications that Discover leaves out, and **Clear List** empties the list.

![The Product filter section](../images/options-product-filter.png)

#### CVE lookup

The **NVD API key**, the **Cache time** and **Automatic lookup**. See [Check CVEs](cve.md).

![The CVE lookup section](../images/options-cve-lookup.png)

### Security

#### Script signing

The Authenticode signing of the scripts that AppPackager stages: which kinds are signed, which signatures are required, the certificate store and the certificate, and the timestamp server. **Test signing configuration** signs and verifies a temporary file. A required signature that fails stops the build. There is no unsigned fallback.

![The Script signing section](../images/options-script-signing.png)

#### GitHub API

Shows how AppPackager signs in to the GitHub API.

![The GitHub API section](../images/options-github-api.png)

### Tools

#### Tools

The detected **Console**, **Content Prep Tool**, **7-Zip** and **Icon pack**. The values are read-only.

![The Tools section](../images/options-tools.png)

### About

#### About

The name, the installed version, the license, the link to the repository, the time of the last update check and the latest known release. **Update Now** shows here when a newer release exists.

![The About section](../images/options-about.png)

### Use your own Microsoft 365 configuration.xml

1. Open **Applications**, then **Microsoft 365 Apps**.
2. Select **Import** and choose your configuration.xml. The tool checks that the file is well-formed XML with a `Configuration` root and an `Add` element that has a `Product` with an `ID`. A file that fails the check is refused with the reason. The tool does not send the file anywhere.
3. Select **Use my own configuration.xml**. The Company name, Channel, Deploy mode, Languages, Excluded apps and install behavior settings are disabled. A line says why.
4. Select **OK**. The tool stores a copy in `Packagers\m365-configuration.user.xml`. The section shows the name and the size of the file. **Remove** deletes the copy.

The Microsoft 365 Apps, Project and Visio packagers use the file. They still set `OfficeClientEdition`, `Version`, `SourcePath` and a silent `Display`, keep the channel of the file, and keep their own detection. A file with a `SourcePath` stages in the Managed mode. A file without one stages in the Online mode. **Preview** shows the result.

### Light theme

The light theme gives the same dialog.

![The Options dialog in the light theme](../images/light-options.png)

![The ConfigMgr section in the light theme](../images/light-options-configmgr.png)

![The Microsoft 365 Apps section in the light theme](../images/light-options-microsoft-365-apps.png)

![The Product filter section in the light theme](../images/light-options-product-filter.png)

### Without the console

Without the console, the ConfigMgr section shows `Not detected` for it. The buttons that need the console are disabled.

![The Options dialog without the console](../images/console-missing-options.png)

## Options and settings of this page

| Section | What it sets |
|---|---|
| Connections | Systems, ConfigMgr, WSUS, Intune and the connection tests. |
| Packaging | Folders, deployment type defaults, the title rule and the steps after Package. |
| Applications | The install options of seven applications and the local installer sources. |
| Updates and CVEs | The tracked set and its destinations, the product filter and the CVE lookup. |
| Security | Authenticode signing of staged scripts and the GitHub API. |
| Tools | The detected tools. |
| About | Version and update check. |

[Settings](../reference/settings.md) lists every key and the section that edits it. [Installer options audit](../reference/installer-options-audit.md) lists the vendor source of each switch.

## Limits

- The settings are saved in `AppPackager.preferences.json` in the install folder. Some packager values are copied to `Packagers\packager-preferences.json`.
- The Citrix Workspace App and TeamViewer Host settings are saved in their own files in `Packagers`.
- A scheduled version monitor has its own configuration file. See [Files and logs](../reference/files-and-logs.md).

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| **Console** shows `Not detected`. | The console is not installed, or its module is not found. | Install the console, or set `SMS_ADMIN_UI_PATH` to its `bin\i386` folder. Start the tool again. |
| **GitHub API** shows `Anonymous`. | No token is set. | Set `GITHUB_TOKEN` or `GH_TOKEN`, or run `gh auth login`. |
| A signing test fails. | No certificate is chosen, or the key needs a PIN. | Choose a certificate by thumbprint. |
| The save stops with a message about a port. | The port is not a whole number from 1 to 65535. | Correct the port. |
| The import of a configuration.xml fails. | The file is not well-formed XML, or it has no `Add` element with a `Product`. | Correct the file. The message names the reason. |

## Links

- [Getting started](../getting-started.md)
- [Settings](../reference/settings.md)
