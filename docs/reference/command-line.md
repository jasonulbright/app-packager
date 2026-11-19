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
.\Invoke-AppPackagerBuild.ps1 -Application package-git.ps1 -Profile default -Target ConfigMgr -Stage -Package -EstimatedMinutes 10 -MaximumMinutes 25
```

### Supersedence

```powershell
# Publish with the version rule Exact and Replace and uninstall, for this run only
.\Invoke-AppPackagerBuild.ps1 -Application catalog:package-7zip -Target ConfigMgr -Package -VersionRule Exact -Supersedence ReplaceAndUninstall

# Record the value that you read from the installed product, so that the Exact guard can allow Exact for this version
.\Invoke-AppPackagerBuild.ps1 -Application catalog:package-7zip -Version 24.08 -RecordInstalledValue 24.08

# Package with the supersedence of the older version applied (it is skipped without the answer)
.\Invoke-AppPackagerBuild.ps1 -Application catalog:package-7zip -Package -Supersedence Replace -SupersedenceAnswer Yes
```

The version rule Exact writes `IsEquals` to the version clause only when a value read from the installed product equals the expected value. A publish that the Exact guard blocks creates nothing.

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
| `-Target` | One of the target names in the Target names section. The default is `ConfigMgr`. |
| `-Stage`, `-Package` | One or both phases. At least one is required. |
| `-Version` | Package this version instead of the latest. |
| `-DownloadRoot` | The local staging root. |
| `-EstimatedMinutes`, `-MaximumMinutes` | Run overrides. They are never written to the profile. |
| `-SiteCode`, `-ProviderMachineName`, `-FileServerPath`, `-Comment` | The ConfigMgr values. |
| `-WsusServer`, `-WsusPort`, `-WsusUseSsl`, `-WsusApprovalGroup` | The WSUS values. |
| `-VersionRule` | `Minimum` or `Exact`. It outranks the title override and the site default of Options, Packaging, Supersedence. |
| `-Supersedence` | `Off`, `Replace` or `ReplaceAndUninstall`. `ReplaceAndUninstall` needs `-VersionRule Exact`. |
| `-RecordInstalledValue`, `-AcceptExactBlock` | Run alone. They store the value read from the installed product, or the acceptance of a block of the Exact guard, for `-Application` and `-Version`. |
| `-SupersedenceAnswer` | `Yes` applies the supersedence of a title whose supersedence is on: the older version (same title and vendor, the highest version below this one) gets one relation from the new application. `Skip` is the default and packages without it. A run that nobody is asked in never applies it by itself. |

## Target names

A target names the destinations of a publish. ConfigMgr, Intune and WSUS are independent, and each one is optional. A target that does not name ConfigMgr needs no site code, no content share and no Configuration Manager console. The default value of `-Target` is `ConfigMgr`. You can choose any other target. A value that is not a target name stops the run and lists the valid names. It never selects a destination. The old names `MECM`, `MECMAndIntune` and `MECMAndWSUS` still load.

| Target | ConfigMgr | Intune | WSUS |
|---|---|---|---|
| `ContentOnly` | No | No | No |
| `ConfigMgr` | Yes | No | No |
| `IntuneOnly` | No | Yes | No |
| `WSUSOnly` | No | No | Yes |
| `ConfigMgrAndIntune` | Yes | Yes | No |
| `ConfigMgrAndWSUS` | Yes | No | Yes |
| `IntuneAndWSUS` | No | Yes | Yes |
| `ConfigMgrAndIntuneAndWSUS` | Yes | Yes | Yes |

`-Package` cannot publish to Intune. Use `-Stage` for a target that names Intune, and publish to Intune in the application window.

## Limits

- A WSUS target publishes the manifest that the same run staged. A target that names both ConfigMgr and WSUS needs `-Stage` and `-Package` together.
- A run override outranks the saved profile for this build only.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| The build does nothing. | Neither `-Stage` nor `-Package` is set. | Set at least one. |

## Links

- [Application Workbench](../core/workbench.md)
- [Network](network.md)
