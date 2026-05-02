# Packager templates

Skeleton packagers for new installer formats. Copy a template into
`Packagers/` (drop the `.template` suffix), rename to
`package-<appname>.ps1`, fill in the `TODO-*` markers, and the main
grid picks it up automatically on next launch.

These files live in a subfolder so they don't pollute the grid as
"Not runnable" rows. They are NOT runnable as-is — each has
`throw "TODO: ..."` guards in every phase.

## Available templates

| File | Format | MECM deployment path | Status |
|---|---|---|---|
| `package-msix.ps1.template` | MSIX / APPX / MSIXBUNDLE | Script (Add-AppxProvisionedPackage) | Framework ready; native Add-CMWindowsAppxDeploymentType path TODO |
| `package-intunewin.ps1.template` | Intunewin (Win32) | Script (delegates to inner MSI/EXE) | Skeleton only; .intunewin crack routine not implemented |
| `package-squirrel.ps1.template` | Squirrel self-update installers | Script (per-user via Active Setup) | Skeleton only; user-context hand-off pattern documented |
| `package-psadt.ps1.template` | PSADT v3 + v4 toolkits | Script (Deploy-Application.exe / Invoke-AppDeployToolkit.ps1) | Skeleton only; wrap-a-wrap shape |
| `package-chocolatey.ps1.template` | Chocolatey / NuGet .nupkg | Script (choco install or inlined chocolateyInstall.ps1) | Skeleton only; relies on choco.exe on targets |

## Forking a template

1. Copy to the parent directory with a concrete name:

   ```powershell
   Copy-Item .\Templates\package-msix.ps1.template .\package-newteams.ps1
   ```

2. Edit the header (Vendor / App / CMName / VendorUrl).
3. Fill the `TODO` blocks (GetLatestVersionOnly, StageOnly, PackageOnly).
4. Re-launch the GUI; the new packager appears in the grid.

## House rules

- Every packager keeps the Script deployment-type shape
  (install.bat → install.ps1 calling the native installer) unless the
  format makes that impossible. MSIX goes through the Script path via
  Add-AppxProvisionedPackage to stay uniform with MSI / EXE packagers.
- Detection clauses prefer RegistryKeyValue (ARP) > File > Script, in
  that order. Script detection is a last resort.
- Every packager supports the three-phase CLI surface
  (`-GetLatestVersionOnly`, `-StageOnly`, `-PackageOnly`) plus a
  default "no phase" usage hint.
