# Tests

Test scaffolding for app-packager. End users do not need to run anything in this folder. This document is for developers extending the codebase.

## Smoke harness

`Invoke-PackagerSmoke.ps1` runs four offline checks against every `Packagers/package-*.ps1`:

1. PowerShell parse (no syntax errors).
2. Required parameter contract present (`-StageOnly`, `-PackageOnly`, `-GetLatestVersionOnly`, etc.).
3. Standard metadata block populated.
4. Wrapper-content generation produces ASCII install/uninstall scripts.

Run it before adding a new packager or editing an existing one.

```powershell
.\Tests\Invoke-PackagerSmoke.ps1

# Optional live vendor-version checks with child-process timeouts
.\Tests\Invoke-PackagerSmoke.ps1 -IncludeLatest -LatestTimeoutSec 90
```

A green run reports `89 script(s), 356 check(s), 356 passed, 0 failed, 0 skipped`.

## Pester

The Pester suite covers shared module behavior in `Packagers/AppPackagerCommon.psm1` and a thin Pester wrapper around the offline smoke harness.

```powershell
Import-Module Pester -RequiredVersion 5.7.1 -Force
Invoke-Pester -Path .\Packagers\AppPackagerCommon.Tests.ps1, .\Tests\PackagerSmoke.Tests.ps1
```

A green run reports `Tests Passed: 111, Failed: 0`.

## Files

- `Invoke-PackagerSmoke.ps1` - offline smoke harness, runnable directly.
- `PackagerSmoke.Tests.ps1` - Pester wrapper around the smoke harness.
- `..\Packagers\AppPackagerCommon.Tests.ps1` - unit tests for the shared module.
