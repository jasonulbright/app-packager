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
| `AppPackager.preferences.json` | The settings of the Options dialog: systems, site, share, download root, One Click settings, the Product Filter, signing, WSUS and the lists of discovered applications. |
| `AppPackager.windowstate.json` | The window size, position, theme and the debug columns. |
| `AppPackager.workbench.windowstate.json` | The size and position of the Workbench window. |
| `Packagers\packager-preferences.json` | The values that packagers read at stage time. The tool copies them from the Options dialog. |
| `Packagers\citrix-workspace-switches.json`, `Packagers\teamviewer-host-config.json` | The Citrix Workspace and TeamViewer Host settings. |
| `%LOCALAPPDATA%\AppPackager\app-history.json` | The latest version, the version on the site, the last check and the CVE summary of each application. |

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
