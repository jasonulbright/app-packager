# Command line

This page shows how to run a packager or a build without the window.

## What it does

Each packager script runs by itself. `Invoke-AppPackagerBuild.ps1` runs one application through a Workbench profile, as the window does.

## When to use it

Use it for a scheduled build, for a test of one packager, or to see what the window runs.

## Requirements

- Windows PowerShell 5.1 and the same settings as the window.
- For the Package phase: the console, the site code and the share.

## Steps

Change to the install folder (by default %LOCALAPPDATA%\AppPackager) or to the repository folder before you run the commands.

### Run a packager

```powershell
# Stage only: download, read the metadata, write the scripts and the manifest
.\Packagers\package-chrome.ps1 -StageOnly

# Package only: read the manifest, copy to the share, create the application
.\Packagers\package-chrome.ps1 -PackageOnly -SiteCode "MCM" -Comment "Initial deployment" -FileServerPath "\\fileserver\sccm$"

# Check the latest version only
.\Packagers\package-chrome.ps1 -GetLatestVersionOnly
```

### Run a build

```powershell
# Stage 7-Zip through the "Managed" profile
.\Invoke-AppPackagerBuild.ps1 -Application catalog:package-7zip -Profile Managed -Stage

# Stage and package, with run overrides for this build only
.\Invoke-AppPackagerBuild.ps1 -Application package-git.ps1 -Profile default -Target MECM -Stage -Package -EstimatedMinutes 10 -MaximumMinutes 25
```

### Convert a ConfigMgr application

```powershell
.\Convert-CMApplicationToIntune.ps1 -ApplicationName 'Contoso Viewer' -SiteCode MCM -OutputRoot D:\Converted
```

This guide did not run these commands. It ran the same actions through the window.

## Options and settings of this page

| Parameter of `Invoke-AppPackagerBuild.ps1` | Meaning |
|---|---|
| `-Application` | `catalog:package-7zip`, `custom:<script>`, `byo:<id>` or a packager file name. |
| `-Profile` | A profile name or id. `default` when omitted. |
| `-Target` | `ContentOnly`, `MECM`, `MECMAndIntune`, `IntuneOnly`, `MECMAndWSUS` or `WSUSOnly`. |
| `-Stage`, `-Package` | One or both phases. At least one is required. |
| `-Version` | Package this version instead of the latest. |
| `-DownloadRoot` | The local staging root. |
| `-EstimatedMinutes`, `-MaximumMinutes` | Run overrides. They are never written to the profile. |
| `-SiteCode`, `-ProviderMachineName`, `-FileServerPath`, `-Comment` | The ConfigMgr values. |
| `-WsusServer`, `-WsusPort`, `-WsusUseSsl`, `-WsusApprovalGroup` | The WSUS values. |

## Limits

- A WSUS target publishes the manifest that the same run staged. `MECMAndWSUS` needs `-Stage` and `-Package` together.
- A run override outranks the saved profile for this build only.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| The build does nothing. | Neither `-Stage` nor `-Package` is set. | Set at least one. |

## Links

- [Application Workbench](../core/workbench.md)
- [Network](network.md)
