# Tests

Test scaffolding for app-packager. End users do not need to run anything in this folder. This document is for developers extending the codebase.

## Smoke harness

`Invoke-PackagerSmoke.ps1` runs five offline checks against every `Packagers/package-*.ps1`:

1. PowerShell parse (no syntax errors).
2. Standard metadata block populated (`Vendor`, `App`).
3. Required parameter contract present (`-StageOnly`, `-PackageOnly`, `-GetLatestVersionOnly`, etc.).
4. The `GetLatestVersionOnly` branch marker is present.
5. Signed launcher policy: the deployment wrappers generated with `APP_PACKAGER_SIGNING` carrying `SignDeployment` on contain no execution-policy or encoded-command argument, and the packager authors no launcher string of its own that does. The wrappers are generated once in a child process through `Write-ContentWrappers` and `New-DeploymentLauncherCommand -Signed $true`; nothing is signed and no certificate is needed.

A separate guard runs once per sweep: every shipped `.ps1`/`.psm1`/`.psd1` outside `Tests`, `Logs` and `Icons` must be pure ASCII.

Run it before adding a new packager or editing an existing one.

```powershell
.\Tests\Invoke-PackagerSmoke.ps1

# Live vendor-version checks, six packagers at a time (about three minutes for the catalog)
.\Tests\Invoke-PackagerSmoke.ps1 -IncludeLatest -LatestTimeoutSec 90 -ThrottleLimit 6

# Full stage sweep: download, wrappers, manifest for every packager, reusing cached installers
.\Tests\Invoke-PackagerSmoke.ps1 -IncludeStage -StageTimeoutSec 1800 -ThrottleLimit 4 -DownloadRoot C:\temp\ap

# One packager end to end
.\Tests\Invoke-PackagerSmoke.ps1 -Packager package-git.ps1 -IncludeLatest -IncludeStage -DownloadRoot C:\temp\ap
```

The offline checks prove a script still parses and keeps the GUI contract; only the live modes prove a packager still works. Run -IncludeLatest after any batch of packager edits and on a schedule, and -IncludeStage before a release that touches download or wrapper code. 90 packagers query the GitHub REST API, which allows 60 unauthenticated calls per hour; set GITHUB_TOKEN (or GH_TOKEN) before a live run so they authenticate (5000 per hour).

Stage-sweep caveats:

- The six `package-m365*.ps1` packagers default to Managed mode, which downloads the full Office content through the Office Deployment Tool (gigabytes per product over a background transfer). Pass them in `-SkipStage` during a sweep and validate them alone, or with `-M365DeployMode Online`, which stages only the bootstrapper.
- `package-windowsadk.ps1` and `package-windowspeaddon.ps1` build their layouts through Windows Installer and cannot run at the same time; they retry on exit 1618, but a sweep with `-ThrottleLimit` above 1 should skip one of them or run them separately.
- A timed-out packager is killed with its process tree, so a stalled download does not leave curl.exe or setup.exe behind.

A green offline run reports `295 script(s), 1475 check(s), 1475 passed, 0 failed, 0 skipped`.

## Catalog matrix

`-Matrix` writes one level-1 inventory row per packager to `Tests/out/catalog-matrix.csv` (or `-MatrixPath`). `Tests/out` is gitignored, so the default output never reaches a commit.

```powershell
.\Tests\Invoke-PackagerSmoke.ps1 -Matrix
.\Tests\Invoke-PackagerSmoke.ps1 -Matrix -MatrixPath C:\temp\ap-matrix\catalog-matrix.csv
```

Columns: packager id, vendor and app, default detector type, the distinct clause types and operators across every authored `Detection` block, compound shape (`Compound`, `Connector`, `GroupSizes`), how many detection blocks the script authors, variant and install-mode capability, launcher source, whether any generated content carries an execution-policy argument, the Intune native-mapping verdict with its finding codes, and the result of each offline check.

The detector shape is read out of the packager source by AST, including the case where the script builds its `Detection` hashtable into a variable first; an offline sweep cannot run a real Stage. The Intune verdict comes from `Get-IntuneCompatibilityFindings` on a representative manifest built from that block: `Unsupported` for any Blocking finding, `Review` for any Review finding, `Ready` otherwise. It reports the base detection only - deployment-type variants are a separate column, because every variant set is an Intune blocker on its own.

## Pester

The Pester suite covers the shared modules (`AppPackagerCommon`, `AppPackagerWorkbench`, `AppPackagerSigning`), the packaging workflow, user detection, script signing, update preservation, and a thin wrapper around the offline smoke harness.

```powershell
Import-Module Pester -RequiredVersion 5.7.1 -Force
Invoke-Pester -Path `
    .\Packagers\AppPackagerCommon.Tests.ps1, .\Packagers\AppPackagerWorkbench.Tests.ps1, .\Packagers\AppPackagerSigning.Tests.ps1, `
    .\Tests\PackagerSmoke.Tests.ps1, .\Tests\PackageWorkflow.Tests.ps1, .\Tests\UserDetection.Tests.ps1, `
    .\Tests\SigningCombinations.Tests.ps1, .\Tests\UpdatePreservation.Tests.ps1
```

`Invoke-FullRegression.ps1` runs that set on both hosts and prints the counts.

### Certificate rule

No test writes to a certificate trust store in any scope. `SigningCombinations.Tests.ps1` creates a throwaway code-signing certificate in `Cert:\CurrentUser\My`, removes it in `AfterAll`, and asserts signature intactness (the hash and the signature block agree), never host trust: the build host is not required to trust the signer, and trusting it there would prove nothing about an endpoint. Adding a certificate to `Root` or `TrustedPublisher` raises an interactive Windows confirmation, which a test must never do.

## Full regression

`Invoke-FullRegression.ps1` runs everything that needs no network, no lab client and no elevation, then prints one summary table and exits non-zero on any failure:

1. Offline packager smoke with the catalog matrix.
2. The combined Pester set under `powershell.exe` 5.1.
3. The same set under `pwsh` 7 (skipped, not passed, when pwsh is absent).
4. Every `Invoke-*Smoke.ps1` UI probe under `powershell.exe -NoProfile -STA`.

```powershell
.\Tests\Invoke-FullRegression.ps1
.\Tests\Invoke-FullRegression.ps1 -SkipPwsh7 -MatrixPath C:\temp\ap-matrix\catalog-matrix.csv
```

`Invoke-AdobeExtractSmoke.ps1` is skipped by default: it needs a local copy of the vendor installer and re-runs itself elevated. Pass `-SkipProbe @()` to include it interactively.

## Files

- `Invoke-PackagerSmoke.ps1` - offline smoke harness, runnable directly.
- `Invoke-OnExistingSmoke.ps1` - existing-application overwrite decision and the GUI's conflict-marker parse.
- `Invoke-FirstRunCallbackSmoke.ps1` - first-run Setup save/skip/close handlers driven through .NET events in a script scope.
- `Invoke-ConflictDialogSmoke.ps1` - existing-application dialog buttons return Skip, Overwrite and Cancel.
- `Invoke-ScopeProbeSmoke.ps1` - launch-scope block exposes script functions to closures; run it with `& <path>`, not `-File`.
- `Invoke-IntuneUploadReview.ps1` - block-blob upload against a loopback receiver, byte-for-byte.
- `Invoke-AdobeExtractSmoke.ps1` - Adobe enterprise installer extraction by tar.exe and by the installer's own switches; needs a local copy of the installer and re-runs itself elevated.
- `Invoke-TeamViewerHostVersionSmoke.ps1` - TeamViewer Host version read from the version resource fixed block when the string table is empty.
- `Invoke-WorkbenchSmoke.ps1` - Application Workbench window build, population, validation and unsaved-switch paths.
- `Invoke-SigningOptionsSmoke.ps1` - the Options script-signing panel, its certificate picker and its test button.
- `Invoke-TitleOptionsSmoke.ps1` - stored title-mode choices reach the background context map.
- `Invoke-FullRegression.ps1` - runs every offline stage on both hosts and prints one summary table.
- `PackagerSmoke.Tests.ps1` - Pester wrapper around the smoke harness.
- `PackageWorkflow.Tests.ps1` - title policy, package conflict preflight, profile precedence, legacy migration, per-profile stage isolation, build selection, run overrides, One Click freshness, and a CLI stage of an offline fixture packager compared against a GUI-equivalent run snapshot.
- `SigningCombinations.Tests.ps1` - all eight signing switch combinations through `Write-StageManifest`, strict-requirement refusals, post-sign mutation detection, non-ASCII content and a stage path containing a space.
- `UpdatePreservation.Tests.ps1` - what an install-root replacement keeps, and that the workbench data root outside the install folder is untouched.
- `UserDetection.Tests.ps1` - the shipped per-user detection rules.
- `..\Packagers\AppPackagerCommon.Tests.ps1` - unit tests for the shared module.
- `..\Packagers\AppPackagerWorkbench.Tests.ps1` - unit tests for the build model.
- `..\Packagers\AppPackagerSigning.Tests.ps1` - unit tests for the signing service.
