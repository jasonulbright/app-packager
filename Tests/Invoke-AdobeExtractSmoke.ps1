#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$ScriptPath,
    [string]$InstallerPath = (Join-Path $env:TEMP 'AdobeReader.exe'),
    [string]$WorkRoot = ([System.IO.Path]::GetTempPath().TrimEnd('\')),
    [switch]$Elevated,
    [switch]$NoElevate
)

$ErrorActionPreference = 'Stop'
if (-not $ScriptPath) { $ScriptPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\Packagers\package-adobereader.ps1')).Path }
if (-not (Test-Path -LiteralPath $InstallerPath)) { throw "Installer payload not found: $InstallerPath" }
if (-not (Test-Path -LiteralPath $WorkRoot)) { throw "Work root not found: $WorkRoot" }

function Test-SmokeElevated {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    return (New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function ConvertTo-SingleQuoted {
    param([string]$Value)
    return "'" + ($Value -replace "'", "''") + "'"
}

# The enterprise installer's manifest asks for administrator, so CreateProcess
# returns ERROR_ELEVATION_REQUIRED from a filtered token and the -sfx path
# cannot start at all. Re-run the whole file elevated and relay its output.
if (-not $Elevated -and -not $NoElevate -and -not (Test-SmokeElevated)) {
    $relayLog = Join-Path $WorkRoot 'adobe-extract-smoke.elevated.log'
    Remove-Item -LiteralPath $relayLog -Force -ErrorAction SilentlyContinue
    $inner = "& {0} -Elevated -ScriptPath {1} -InstallerPath {2} -WorkRoot {3} *> {4}" -f `
        (ConvertTo-SingleQuoted $PSCommandPath), (ConvertTo-SingleQuoted $ScriptPath),
        (ConvertTo-SingleQuoted $InstallerPath), (ConvertTo-SingleQuoted $WorkRoot),
        (ConvertTo-SingleQuoted $relayLog)

    Write-Output "Session is not elevated; re-running elevated."
    $child = Start-Process -FilePath 'powershell.exe' -Verb RunAs -WindowStyle Hidden -Wait -PassThru `
        -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', $inner)
    $childExit = $child.ExitCode
    $child.Dispose()

    if (Test-Path -LiteralPath $relayLog) {
        Get-Content -LiteralPath $relayLog | ForEach-Object { Write-Output $_ }
        Remove-Item -LiteralPath $relayLog -Force -ErrorAction SilentlyContinue
    }
    if ($childExit -ne 0) { throw "Elevated run failed with exit code $childExit." }
    exit 0
}

# Lifts Expand-AdobeInstaller out of the packager and runs both extraction
# paths against the real enterprise installer. Nothing is installed; each
# extraction folder is about 900 MB and is removed after its checks.
$tokens = $null; $parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors.Message -join '; ') }
$fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Expand-AdobeInstaller' }, $false)
if (-not $fn) { throw "Cannot locate Expand-AdobeInstaller in $ScriptPath." }

function Write-Log {
    param(
        [Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')][string]$Level = 'INFO',
        [switch]$Quiet
    )
    if (-not $Quiet) { Write-Host ("      [{0,-5}] {1}" -f $Level, $Message) }
}

. ([scriptblock]::Create($fn.Extent.Text))

$RequiredNames = @('abcpy.ini', 'AcroRead.msi', 'Data1.cab', 'setup.exe', 'setup.ini')
$script:Failed = 0
$script:Inventory = @{}

function Test-Assert {
    param([Parameter(Mandatory)][string]$Label, [bool]$Condition, [string]$Detail = '')

    if ($Condition) { Write-Output "PASS: $Label" }
    else {
        $script:Failed++
        Write-Output ("FAIL: $Label" + $(if ($Detail) { " - $Detail" } else { '' }))
    }
}

function Invoke-ExtractScenario {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$ExpectedMethod,
        [hashtable]$ExtraArgs = @{}
    )

    $destination = Join-Path $WorkRoot ("adobe-extract-" + $Name)
    if (Test-Path -LiteralPath $destination) { Remove-Item -LiteralPath $destination -Recurse -Force }
    New-Item -ItemType Directory -Path $destination -Force | Out-Null

    Write-Output ""
    Write-Output "--- scenario: $Name (expect '$ExpectedMethod') ---"
    Write-Output "    destination: $destination"

    try {
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $method = Expand-AdobeInstaller -InstallerPath $InstallerPath -Destination $destination @ExtraArgs
        $watch.Stop()

        $files = @(Get-ChildItem -LiteralPath $destination -File | Sort-Object Name)
        $script:Inventory[$Name] = @($files | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Length = $_.Length } })
        foreach ($file in $files) { Write-Output ("    {0,-32} {1,12:N0}" -f $file.Name, $file.Length) }
        Write-Output ("    elapsed: {0:N1}s" -f $watch.Elapsed.TotalSeconds)

        Test-Assert -Label "$Name returns '$ExpectedMethod'" -Condition ($method -eq $ExpectedMethod) -Detail "got '$method'"
        Test-Assert -Label "$Name extracts 6 files" -Condition ($files.Count -eq 6) -Detail "got $($files.Count)"
        foreach ($required in $RequiredNames) {
            Test-Assert -Label "$Name extracts $required" -Condition ([bool]($files | Where-Object { $_.Name -eq $required }))
        }
        $patches = @($files | Where-Object { $_.Extension -eq '.msp' })
        Test-Assert -Label "$Name extracts one .msp" -Condition ($patches.Count -eq 1) -Detail "got $($patches.Count)"
        Test-Assert -Label "$Name extracts no empty files" -Condition (-not ($files | Where-Object { $_.Length -eq 0 }))
    }
    finally {
        Remove-Item -LiteralPath $destination -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Output "Packager     : $ScriptPath"
Write-Output "Installer    : $InstallerPath"
Write-Output ("Installer MB : {0:N1}" -f ((Get-Item -LiteralPath $InstallerPath).Length / 1MB))
Write-Output ("Elevated     : {0}" -f (Test-SmokeElevated))

Invoke-ExtractScenario -Name 'tar' -ExpectedMethod 'tar'
Invoke-ExtractScenario -Name 'sfx' -ExpectedMethod 'sfx' -ExtraArgs @{ TarPath = 'C:\nonexistent\tar.exe' }

Write-Output ""
Write-Output "--- comparison ---"
$left  = $script:Inventory['tar']
$right = $script:Inventory['sfx']
if (-not $left -or -not $right) {
    Test-Assert -Label 'both scenarios produced an inventory' -Condition $false -Detail 'one scenario did not complete'
}
else {
    $leftNames  = ($left  | ForEach-Object { $_.Name }) -join ', '
    $rightNames = ($right | ForEach-Object { $_.Name }) -join ', '
    Test-Assert -Label 'file names match across methods' -Condition ($leftNames -eq $rightNames) -Detail "tar: $leftNames / sfx: $rightNames"

    $mismatched = @()
    foreach ($entry in $left) {
        $other = $right | Where-Object { $_.Name -eq $entry.Name } | Select-Object -First 1
        if (-not $other -or $other.Length -ne $entry.Length) {
            $mismatched += ("{0} (tar {1} / sfx {2})" -f $entry.Name, $entry.Length, $(if ($other) { $other.Length } else { 'missing' }))
        }
    }
    Test-Assert -Label 'file sizes match across methods' -Condition ($mismatched.Count -eq 0) -Detail ($mismatched -join '; ')
}

Write-Output ""
if ($script:Failed) { throw "$script:Failed check(s) failed." }
Write-Output "All Adobe extraction checks passed."
