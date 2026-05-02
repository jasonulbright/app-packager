#Requires -Version 5.1

<#
.SYNOPSIS
    Fast smoke checks for every AppPackager package-*.ps1 script.

.DESCRIPTION
    Defaults to offline checks that are safe to run on any developer machine:
    PowerShell parse, required GUI metadata, standard parameter contract, and
    the GetLatestVersionOnly code path marker.

    Optional live checks can run each packager's GetLatestVersionOnly or
    StageOnly mode with child-process timeouts. Those modes may touch vendor
    endpoints or download installers, so they are opt-in.

.EXAMPLE
    .\Tests\Invoke-PackagerSmoke.ps1

.EXAMPLE
    .\Tests\Invoke-PackagerSmoke.ps1 -IncludeLatest -LatestTimeoutSec 90
#>

[CmdletBinding()]
param(
    [string]$PackagersRoot,

    [string[]]$Packager,

    [switch]$IncludeLatest,

    [int]$LatestTimeoutSec = 60,

    [switch]$IncludeStage,

    [int]$StageTimeoutSec = 900,

    [string]$DownloadRoot = (Join-Path ([System.IO.Path]::GetTempPath()) 'AppPackagerSmoke'),

    [string]$SiteCode = 'MCM',

    [string[]]$SkipLatest = @(),

    [string[]]$SkipStage = @(),

    [switch]$Json
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($PackagersRoot)) {
    $PackagersRoot = Join-Path $PSScriptRoot '..\Packagers'
}

function New-SmokeResult {
    param(
        [Parameter(Mandatory)][string]$Script,
        [Parameter(Mandatory)][string]$Check,
        [Parameter(Mandatory)][ValidateSet('Pass','Fail','Skip')][string]$Status,
        [string]$Detail = ''
    )

    [pscustomobject]@{
        Script = $Script
        Check  = $Check
        Status = $Status
        Detail = $Detail
    }
}

function Get-PackagerHeaderMetadata {
    param([Parameter(Mandatory)][string]$Path)

    $meta = [ordered]@{
        Vendor            = $null
        App               = $null
        CMName            = $null
        VendorUrl         = $null
        CPE               = $null
        ReleaseNotesUrl   = $null
        DownloadPageUrl   = $null
        UpdateCadenceDays = $null
        Description       = $null
    }

    foreach ($line in (Get-Content -LiteralPath $Path -TotalCount 80 -ErrorAction Stop)) {
        if (-not $meta.Vendor          -and $line -match '^\s*(?:#\s*)?Vendor\s*:\s*(.+?)\s*$')          { $meta.Vendor          = $Matches[1].Trim(); continue }
        if (-not $meta.App             -and $line -match '^\s*(?:#\s*)?App\s*:\s*(.+?)\s*$')             { $meta.App             = $Matches[1].Trim(); continue }
        if (-not $meta.CMName          -and $line -match '^\s*(?:#\s*)?CMName\s*:\s*(.+?)\s*$')          { $meta.CMName          = $Matches[1].Trim(); continue }
        if (-not $meta.VendorUrl       -and $line -match '^\s*(?:#\s*)?VendorUrl\s*:\s*(.+?)\s*$')       { $meta.VendorUrl       = $Matches[1].Trim(); continue }
        if (-not $meta.CPE             -and $line -match '^\s*(?:#\s*)?CPE\s*:\s*(.+?)\s*$')             { $meta.CPE             = $Matches[1].Trim(); continue }
        if (-not $meta.ReleaseNotesUrl -and $line -match '^\s*(?:#\s*)?ReleaseNotesUrl\s*:\s*(.+?)\s*$') { $meta.ReleaseNotesUrl = $Matches[1].Trim(); continue }
        if (-not $meta.DownloadPageUrl -and $line -match '^\s*(?:#\s*)?DownloadPageUrl\s*:\s*(.+?)\s*$') { $meta.DownloadPageUrl = $Matches[1].Trim(); continue }
        if ($null -eq $meta.UpdateCadenceDays -and $line -match '^\s*(?:#\s*)?UpdateCadenceDays\s*:\s*(\d+)\s*$') {
            $meta.UpdateCadenceDays = [int]$Matches[1]
            continue
        }
        if (-not $meta.Description -and $line -match '^\s*(?:#\s*)?Description\s*:\s*(.+?)\s*$') {
            $meta.Description = $Matches[1].Trim()
            continue
        }
    }

    [pscustomobject]$meta
}

function ConvertTo-CommandLineArgument {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) { return '""' }
    if ($Value -notmatch '[\s"]') { return $Value }

    $escaped = $Value -replace '(\\*)"', '$1$1\"'
    $escaped = $escaped -replace '(\\+)$', '$1$1'
    return '"' + $escaped + '"'
}

function Invoke-ChildProcess {
    param(
        [Parameter(Mandatory)][string]$FileName,
        [Parameter(Mandatory)][string[]]$ArgumentList,
        [Parameter(Mandatory)][int]$TimeoutSec
    )

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FileName
    $psi.Arguments = (($ArgumentList | ForEach-Object { ConvertTo-CommandLineArgument $_ }) -join ' ')
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true

    $stdout = New-Object System.Text.StringBuilder
    $stderr = New-Object System.Text.StringBuilder
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi

    $outHandler = [System.Diagnostics.DataReceivedEventHandler]{
        param($sender, $eventArgs)
        if ($null -ne $eventArgs.Data) { [void]$stdout.AppendLine($eventArgs.Data) }
    }
    $errHandler = [System.Diagnostics.DataReceivedEventHandler]{
        param($sender, $eventArgs)
        if ($null -ne $eventArgs.Data) { [void]$stderr.AppendLine($eventArgs.Data) }
    }

    $timedOut = $false

    try {
        $proc.add_OutputDataReceived($outHandler)
        $proc.add_ErrorDataReceived($errHandler)

        [void]$proc.Start()
        $proc.BeginOutputReadLine()
        $proc.BeginErrorReadLine()

        if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
            $timedOut = $true
            try { $proc.Kill() } catch { }
        }
        else {
            $proc.WaitForExit()
        }

        [pscustomobject]@{
            ExitCode = if ($timedOut) { $null } else { $proc.ExitCode }
            TimedOut = $timedOut
            StdOut   = $stdout.ToString()
            StdErr   = $stderr.ToString()
        }
    }
    finally {
        $proc.remove_OutputDataReceived($outHandler)
        $proc.remove_ErrorDataReceived($errHandler)
        $proc.Dispose()
    }
}

function Test-PackagerSyntaxAndContract {
    param([Parameter(Mandatory)][System.IO.FileInfo]$File)

    $results = New-Object System.Collections.Generic.List[object]
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($File.FullName, [ref]$tokens, [ref]$parseErrors)

    if ($parseErrors.Count -gt 0) {
        $detail = (($parseErrors | ForEach-Object { $_.Message }) -join '; ')
        $results.Add((New-SmokeResult -Script $File.Name -Check 'Parse' -Status Fail -Detail $detail))
    }
    else {
        $results.Add((New-SmokeResult -Script $File.Name -Check 'Parse' -Status Pass -Detail 'PowerShell parser accepted script.'))
    }

    $meta = Get-PackagerHeaderMetadata -Path $File.FullName
    $missingMeta = @()
    if ([string]::IsNullOrWhiteSpace($meta.Vendor)) { $missingMeta += 'Vendor' }
    if ([string]::IsNullOrWhiteSpace($meta.App))    { $missingMeta += 'App' }

    if ($missingMeta.Count -gt 0) {
        $results.Add((New-SmokeResult -Script $File.Name -Check 'Metadata' -Status Fail -Detail ("Missing: {0}" -f ($missingMeta -join ', '))))
    }
    else {
        $results.Add((New-SmokeResult -Script $File.Name -Check 'Metadata' -Status Pass -Detail ("{0} / {1}" -f $meta.Vendor, $meta.App)))
    }

    $paramNames = @()
    if ($ast.ParamBlock) {
        $paramNames = @($ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
    }

    $requiredParams = @(
        'SiteCode',
        'Comment',
        'FileServerPath',
        'DownloadRoot',
        'EstimatedRuntimeMins',
        'MaximumRuntimeMins',
        'LogPath',
        'GetLatestVersionOnly',
        'StageOnly',
        'PackageOnly'
    )

    $missingParams = @($requiredParams | Where-Object { $paramNames -notcontains $_ })
    if ($missingParams.Count -gt 0) {
        $results.Add((New-SmokeResult -Script $File.Name -Check 'Parameter contract' -Status Fail -Detail ("Missing: {0}" -f ($missingParams -join ', '))))
    }
    else {
        $results.Add((New-SmokeResult -Script $File.Name -Check 'Parameter contract' -Status Pass -Detail 'Standard GUI parameters present.'))
    }

    $source = Get-Content -LiteralPath $File.FullName -Raw -ErrorAction Stop
    if ($source -notmatch 'if\s*\(\s*\$GetLatestVersionOnly\s*\)') {
        $results.Add((New-SmokeResult -Script $File.Name -Check 'Latest mode marker' -Status Fail -Detail 'No GetLatestVersionOnly branch marker found.'))
    }
    else {
        $results.Add((New-SmokeResult -Script $File.Name -Check 'Latest mode marker' -Status Pass -Detail 'GetLatestVersionOnly marker present.'))
    }

    $results
}

function Invoke-PackagerLatestSmoke {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [Parameter(Mandatory)][string]$PowerShellExe
    )

    if ($SkipLatest -contains $File.Name) {
        return New-SmokeResult -Script $File.Name -Check 'Latest live' -Status Skip -Detail 'Skipped by -SkipLatest.'
    }

    $args = @(
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        $File.FullName,
        '-SiteCode',
        $SiteCode,
        '-DownloadRoot',
        $DownloadRoot,
        '-GetLatestVersionOnly'
    )

    $result = Invoke-ChildProcess -FileName $PowerShellExe -ArgumentList $args -TimeoutSec $LatestTimeoutSec
    if ($result.TimedOut) {
        return New-SmokeResult -Script $File.Name -Check 'Latest live' -Status Fail -Detail ("Timed out after {0}s." -f $LatestTimeoutSec)
    }
    if ($result.ExitCode -ne 0) {
        $detail = ($result.StdErr.Trim(), $result.StdOut.Trim() | Where-Object { $_ }) -join ' '
        return New-SmokeResult -Script $File.Name -Check 'Latest live' -Status Fail -Detail ("Exit {0}. {1}" -f $result.ExitCode, $detail).Trim()
    }

    $version = @($result.StdOut -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Last 1)
    if ($version.Count -eq 0 -or $version[0] -notmatch '^[0-9][0-9A-Za-z.+_-]*$') {
        return New-SmokeResult -Script $File.Name -Check 'Latest live' -Status Fail -Detail ("Unexpected version output: {0}" -f $result.StdOut.Trim())
    }

    New-SmokeResult -Script $File.Name -Check 'Latest live' -Status Pass -Detail $version[0]
}

function Invoke-PackagerStageSmoke {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [Parameter(Mandatory)][string]$PowerShellExe
    )

    if ($SkipStage -contains $File.Name) {
        return New-SmokeResult -Script $File.Name -Check 'Stage live' -Status Skip -Detail 'Skipped by -SkipStage.'
    }

    $args = @(
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        $File.FullName,
        '-SiteCode',
        $SiteCode,
        '-DownloadRoot',
        $DownloadRoot,
        '-StageOnly'
    )

    $result = Invoke-ChildProcess -FileName $PowerShellExe -ArgumentList $args -TimeoutSec $StageTimeoutSec
    if ($result.TimedOut) {
        return New-SmokeResult -Script $File.Name -Check 'Stage live' -Status Fail -Detail ("Timed out after {0}s." -f $StageTimeoutSec)
    }
    if ($result.ExitCode -ne 0) {
        $detail = ($result.StdErr.Trim(), $result.StdOut.Trim() | Where-Object { $_ }) -join ' '
        return New-SmokeResult -Script $File.Name -Check 'Stage live' -Status Fail -Detail ("Exit {0}. {1}" -f $result.ExitCode, $detail).Trim()
    }

    New-SmokeResult -Script $File.Name -Check 'Stage live' -Status Pass -Detail 'StageOnly completed.'
}

$resolvedRoot = Resolve-Path -LiteralPath $PackagersRoot -ErrorAction Stop
$files = @(Get-ChildItem -LiteralPath $resolvedRoot -Filter 'package-*.ps1' -File | Sort-Object Name)
if ($Packager -and $Packager.Count -gt 0) {
    $wanted = @($Packager | ForEach-Object { $_.ToLowerInvariant() })
    $files = @($files | Where-Object {
        $wanted -contains $_.Name.ToLowerInvariant() -or
        $wanted -contains $_.BaseName.ToLowerInvariant()
    })
}

if ($files.Count -eq 0) {
    throw "No packager scripts matched under '$resolvedRoot'."
}

$allResults = New-Object System.Collections.Generic.List[object]
$powershellExe = (Get-Command powershell.exe -ErrorAction Stop).Source

foreach ($file in $files) {
    foreach ($result in (Test-PackagerSyntaxAndContract -File $file)) {
        $allResults.Add($result)
    }

    if ($IncludeLatest) {
        $allResults.Add((Invoke-PackagerLatestSmoke -File $file -PowerShellExe $powershellExe))
    }

    if ($IncludeStage) {
        $allResults.Add((Invoke-PackagerStageSmoke -File $file -PowerShellExe $powershellExe))
    }
}

$resultArray = @($allResults.ToArray())

$summary = [pscustomobject]@{
    Packagers = $files.Count
    Checks    = $resultArray.Count
    Passed    = @($resultArray | Where-Object { $_.Status -eq 'Pass' }).Count
    Failed    = @($resultArray | Where-Object { $_.Status -eq 'Fail' }).Count
    Skipped   = @($resultArray | Where-Object { $_.Status -eq 'Skip' }).Count
    Results   = $resultArray
}

if ($Json) {
    $summary | ConvertTo-Json -Depth 5
}
else {
    Write-Host ("Packager smoke: {0} script(s), {1} check(s), {2} passed, {3} failed, {4} skipped" -f $summary.Packagers, $summary.Checks, $summary.Passed, $summary.Failed, $summary.Skipped)
    $failed = @($resultArray | Where-Object { $_.Status -eq 'Fail' })
    if ($failed.Count -gt 0) {
        Write-Host ''
        Write-Host 'Failures:'
        foreach ($failure in $failed) {
            Write-Host ("  [{0}] {1} - {2}: {3}" -f $failure.Status, $failure.Script, $failure.Check, $failure.Detail)
        }
    }
}

if ($summary.Failed -gt 0) {
    exit 1
}
