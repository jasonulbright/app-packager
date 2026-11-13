# Settings

This page lists what AppPackager saves and where.

## What it does

The tool saves its settings in JSON files. It saves the Intune client secret and the NVD API key in protected form (`ClientSecretProtected`, `NvdApiKeyProtected`). Only the same Windows account on the same computer can read them back. It saves no other password.

## When to use it

Use it to reset the tool, to move the settings to another computer or to check what the tool keeps.

## Requirements

None.

## Steps

### Reset the tool

Close the tool. Delete `AppPackager.preferences.json` and `AppPackager.windowstate.json` in the install folder. The next start shows the Setup window.

## Options and settings of this page

| File | Content |
|---|---|
| `AppPackager.preferences.json` | The settings of the Options dialog: systems, site, share, download root, One Click settings, the product filter, signing, WSUS and the lists of discovered applications. |
| `AppPackager.windowstate.json` | The window size, position, theme and the debug columns. |
| `AppPackager.workbench.windowstate.json` | The size and position of the Workbench window. |
| `Packagers\packager-preferences.json` | The values that packagers read at stage time. The tool copies them from the Options dialog. |
| `Packagers\m365-configuration.user.xml` | The copy of the configuration.xml that you imported in Options, Applications, Microsoft 365 Apps. It stays on this computer. An update keeps it. |
| `Packagers\citrix-workspace-switches.json`, `Packagers\teamviewer-host-config.json` | The Citrix Workspace and TeamViewer Host settings. |
| `%LOCALAPPDATA%\AppPackager\app-history.json` | The latest version, the version on the site, the last check and the CVE summary of each application. |

## Where each setting is edited

Every key of `AppPackager.preferences.json`, `Packagers\citrix-workspace-switches.json` and `Packagers\teamviewer-host-config.json` has a field in the Options dialog, except the keys in the table below.

| Key | Why Options has no field |
|---|---|
| `FirstRunCompleted` | The Setup window writes it. The box **Don't show this again** and **Save** set it. |
| `DiscoveryEnabled` | The Discover unmanaged software window writes it when you enable an application. |
| `DeploymentConditions.Apps` | The conditions dialog of the main window edits it for each application. |
| `CommandOverrides.Apps` | The command override dialog of the main window edits it for each application. |
| `WorkbenchDataRoot` | File only. It sets the data folder of the Workbench. The variable `APP_PACKAGER_WORKBENCH_ROOT` overrides it. |
| `Installation.EnableSSON` (Citrix Workspace App) | It follows **Single sign-on** in Options, Applications, Citrix Workspace App, Installation. |
| `DetectedTools` | The tool detects the console, 7-Zip and the Content Prep Tool at start. The Tools section shows them read-only. |

The keys `DiscoverySqlInstance`, `DiscoverySqlPort`, `M365Options`, `M365CustomConfig`, `BeyondCompareOptions` and the new keys of `SSMSInstallOptions` and `AdobeReaderInstallOptions` are optional. A file without them loads with their defaults, and every other key keeps its name and its meaning. A name `server\instance,port` in `DiscoverySqlServer` still works, and Options shows it split.

The environment variables `GITHUB_TOKEN`, `GH_TOKEN` and `NVD_API_KEY` supply tokens without a file.

## Limits

- An in-place update keeps these files. See the update section of the README.
- The tool does not read settings from the registry.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| The Setup window shows again. | The preferences file is missing. | Answer it again, or restore the file. |

## Links

- [Files and logs](files-and-logs.md)
- [Options dialog](../core/options.md)
