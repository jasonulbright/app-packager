# Sites and site data

This page explains how the tool keeps the data of each ConfigMgr site apart, and what it copies on the first start.

## What it does

A site is one ConfigMgr environment, named by its site code. The tool keeps the data of each site in its own folder. The site code `AB1` and the site code `CD2` never share a setting or a result. The data that does not describe a site stays in one place for all sites.

The site folder is `%LOCALAPPDATA%\AppPackagerData\sites\<site code>`. It is outside the install folder, so an update of the tool never touches it. App Packager Desktop and the suite tools use the same folder names.

## When to use it

Use it when you manage more than one site, for example a production site and a test site, and you want each one to keep its own provider, share, distribution point group and check results.

## Requirements

- A site code of three letters or digits, for example `AB1`. The codes `AUX`, `CON`, `ENV`, `NUL`, `PRN` and `SMS` are not allowed, because Windows and ConfigMgr reserve them as folder names.
- The Setup window has finished. Until then the tool has no site.

## Steps

### Switch to another site

1. Select **Options**. The ConfigMgr section opens.
2. Open the list **Site code**. It shows the sites that have a folder. Pick one, or type a new site code.
3. Select **OK**.

The tool saves the changes that you made in the window to the site that was active. Then it opens the site that you picked. A site that is new starts with empty site fields: enter its **Site server / SMS Provider** and the other values. The footer shows the new site. The tool does not switch while a run, a Check or a Discover scan is active.

A second environment that reuses a site code is written `AB1~2`. Type it in the list to open or create it.

### Start in a site without Options

The tool opens the site in this order: the `-SiteCode` value, the variable `SUITE_CM_SITECODE` that the suite launcher sets, the variable `APP_PACKAGER_SITE`, the file `active-site.json` in `%LOCALAPPDATA%\AppPackagerData`, and the site code that the preferences file saved. A value that is not a valid site code is skipped.

## Options and settings of this page

| Data | Where it is kept | Scope |
|---|---|---|
| Site code, **Site server / SMS Provider**, SQL server, instance, port and database name of Discover, **Content share UNC** | `sites\<site>\settings.json` | Site |
| The DP group and the test collection (site-wide and for each application), the WSUS server | `sites\<site>\settings.json` | Site |
| The ConfigMgr version of each application, and when Check ConfigMgr read it | `sites\<site>\state\app-state.json` | Site |
| `CheckConfigMgrResult.csv` and `CheckConfigMgrResult.json` | `sites\<site>\results\` | Site |
| The identity of the site: its code, a label, the providers and the last use | `sites\<site>\site.json` | Site |
| The list of the files that the first start copied | `sites\<site>\migration.json` | Site |
| The site that the next start opens | `active-site.json` in `%LOCALAPPDATA%\AppPackagerData` | All sites |
| The latest version, the last check, the CVE summary and the other history of each application | `%LOCALAPPDATA%\AppPackager\app-history.json` | All sites |
| The packaging options, the per-application options, the One Click settings, the theme, the logs and the CVE data | The preferences file and the folders that they had before | All sites |
| The GitHub token, the Intune client secret, the NVD key and the TeamViewer API token | The preferences file and the packager files (DPAPI for the current Windows user, as before) | All sites. Never copied into a site folder. |

The preferences file `AppPackager.preferences.json` keeps a copy of the values of the active site, so that a packager that runs alone reads them. The site folder is the store. When you switch, the tool writes the values of the new site into the file.

The variable `APP_PACKAGER_DATA_ROOT` moves the data root. The folder is then `<value>\Users\<Windows SID>\sites\<site code>`.

## The first start

The first start of this version copies the files of the single-site layout into the folder of the site code that the preferences file saved. The copy never moves a file. The tool compares the SHA-256 of each copy with its source. A copy that does not match stops the run, writes the reason to the log, and the tool keeps using the old files. The old files stay where they are, with the same content. A report of every item is in the log.

The copy takes these files: the site keys of the preferences file, the ConfigMgr state of each application from `app-history.json`, and the two result files. It copies no secret. A site key that the copy leaves out, for example the SQL server of Discover, is added from the preferences file. The preferences file records the site that the copy ran for, in the key `SiteDataMigratedFrom`, so the copy runs once.

If you started a different site from the suite launcher, the tool still copies the old files into the folder of the site code that they belong to, and opens the other site with empty site fields.

## Limits

- Only the folder of the active site changes when you save. Another site is never written.
- The Version Monitor keeps its own `monitor-config.json` beside the script. Its site code and server are copied into the site folder, and the monitor still reads the file.
- The old files are kept for at least one release. Delete them by hand when you no longer need them.
- App Packager Desktop writes the same two result files in the same site folder. A file holds the result of the last tool that wrote it.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| The site code is three letters or digits | The value has another length or a reserved name. | Type three letters or digits, for example `AB1`. |
| A run is active. Switch the site when it has finished. | A run, a Check or a Discover scan is active. | Wait for it to finish, then switch. |
| A new site shows empty fields | The site has no saved settings. | Enter the provider and the other values. Select **OK**. |
| The log says that a site data folder is not used | A copy did not match its source. | The tool works with the old files. Fix the reason in the log, then start the tool again. |
