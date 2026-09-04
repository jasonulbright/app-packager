#Requires -Version 5.1

<#
.SYNOPSIS
    Downloads, verifies, and installs or updates AppPackager from its GitHub releases.

.DESCRIPTION
    Resolves a release from the jasonulbright/app-packager releases API, downloads
    the AppPackager-<version>.zip asset, verifies its SHA-256 against the release's
    checksums.txt, and extracts it into the install folder.

    Downloads go through curl.exe first with Invoke-WebRequest as the fallback, and extraction through
    Expand-Archive, so no extracted file carries a Mark-of-the-Web: module imports in packager child
    processes fail non-terminating when the block is present, surfacing later as unrelated
    unknown-command errors.

    On an update, user state kept inside the application folder (every *.json plus
    the Logs folder - the same set .gitignore excludes from the release zip) is moved
    aside, the folder is replaced, and the state is restored.

    Bootstrap path: download AppPackager.zip from the latest release with curl.exe and run this script from inside it with -ZipPath (see README).

.PARAMETER InstallPath
    Target folder. Defaults to %LOCALAPPDATA%\AppPackager.

.PARAMETER Version
    Release version to install, e.g. "1.5.0.4". Defaults to the latest release.

.PARAMETER Force
    Replace the folder even when it holds no recognizable AppPackager install.

.EXAMPLE
    .\install.ps1

.EXAMPLE
    .\install.ps1 -InstallPath 'D:\Tools\AppPackager' -Version 1.5.0.3

.NOTES
    ScriptName : install.ps1
    Purpose    : Bootstrap install / update for AppPackager
    Owner      : CM Engineering
    Version    : 1.5.1.6
#>

[CmdletBinding()]
param(
    [string]$InstallPath = (Join-Path $env:LOCALAPPDATA 'AppPackager'),
    [string]$Version,
    [string]$ZipPath,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$script:Repo      = 'jasonulbright/app-packager'
$script:UserAgent = 'AppPackager-Installer'

function Write-Step {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Message)
    Write-Host $Message
}

function Get-ReleaseMetadata {
    param([string]$Version)

    $uri = if ([string]::IsNullOrWhiteSpace($Version)) {
        "https://api.github.com/repos/$script:Repo/releases/latest"
    } else {
        "https://api.github.com/repos/$script:Repo/releases/tags/v$($Version.TrimStart('v'))"
    }

    $json = Invoke-InstallerDownloadText -Url $uri
    $json | ConvertFrom-Json
}

function Invoke-InstallerDownloadText {
    # curl.exe first: proxies that break Invoke-WebRequest's chain (SSL
    # inspection) commonly pass curl, which the packagers already rely on.
    param([Parameter(Mandatory)][string]$Url)

    $curl = Get-Command curl.exe -ErrorAction SilentlyContinue
    if ($curl) {
        $text = (& $curl.Source -L --fail --silent --show-error -A $script:UserAgent $Url) -join "`n"
        if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($text)) { return $text }
    }
    Invoke-RestMethod -Uri $Url -Headers @{ 'User-Agent' = $script:UserAgent } -UseBasicParsing
}

function Invoke-InstallerDownloadFile {
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$OutFile
    )

    $curl = Get-Command curl.exe -ErrorAction SilentlyContinue
    if ($curl) {
        & $curl.Source -L --fail --silent --show-error -A $script:UserAgent -o $OutFile $Url
        if ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $OutFile)) { return }
    }
    Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -Headers @{ 'User-Agent' = $script:UserAgent }
}

function Get-ReleaseAssetUrl {
    param(
        [Parameter(Mandatory)]$Release,
        [Parameter(Mandatory)][string]$Name
    )

    $asset = @($Release.assets) | Where-Object { $_.name -eq $Name } | Select-Object -First 1
    if (-not $asset) {
        throw ("Release {0} has no asset named {1}." -f $Release.tag_name, $Name)
    }
    $asset.browser_download_url
}

function Get-ChecksumForFile {
    # checksums.txt is sha256sum output: "<hash> *<filename>" or "<hash>  <filename>".
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$ChecksumText,
        [Parameter(Mandatory)][string]$FileName
    )

    foreach ($line in ($ChecksumText -split "`r?`n")) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $m = [regex]::Match($line.Trim(), '^([0-9a-fA-F]{64})\s+\*?(.+)$')
        if (-not $m.Success) { continue }
        if ($m.Groups[2].Value.Trim() -eq $FileName) { return $m.Groups[1].Value.ToLowerInvariant() }
    }
    return $null
}

function Get-PreservedStateFile {
    # Mirrors the .gitignore rules that keep these files out of the release zip,
    # so nothing restored here can be clobbered by an extracted file.
    param([Parameter(Mandatory)][string]$Root)

    if (-not (Test-Path -LiteralPath $Root)) { return @() }

    $items = @(Get-ChildItem -LiteralPath $Root -Recurse -File -Filter '*.json' -ErrorAction SilentlyContinue)
    $logs  = Join-Path $Root 'Logs'
    if (Test-Path -LiteralPath $logs) {
        $items += @(Get-ChildItem -LiteralPath $logs -Recurse -File -ErrorAction SilentlyContinue)
    }

    $seen = @{}
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($item in $items) {
        $relative = $item.FullName.Substring($Root.Length).TrimStart('\')
        if ($seen.ContainsKey($relative)) { continue }
        $seen[$relative] = $true
        [void]$result.Add($relative)
    }
    $result.ToArray()
}

function Test-AppPackagerFolder {
    param([Parameter(Mandatory)][string]$Path)
    Test-Path -LiteralPath (Join-Path $Path 'start-apppackager.ps1')
}

function Invoke-Install {
    param(
        [Parameter(Mandatory)][string]$InstallPath,
        [string]$Version,
        [string]$ZipPath,
        [switch]$Force
    )

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $work = Join-Path ([IO.Path]::GetTempPath()) ("apinstall-" + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $work -Force | Out-Null

    try {
        if (-not [string]::IsNullOrWhiteSpace($ZipPath)) {
            # Pre-downloaded zip: content filters that block script and text
            # downloads still allow archives, so the zip may be the only file
            # that can reach the machine. The checksum is verified only when a
            # checksums.txt sits beside the zip.
            if (-not (Test-Path -LiteralPath $ZipPath -PathType Leaf)) {
                throw ("Zip not found: {0}" -f $ZipPath)
            }
            Write-Step ("Installing from local zip    : {0}" -f $ZipPath)
            $zipName = [IO.Path]::GetFileName($ZipPath)
            $localZip = Join-Path $work $zipName
            Copy-Item -LiteralPath $ZipPath -Destination $localZip -Force
            Unblock-File -LiteralPath $localZip -ErrorAction SilentlyContinue
            $archivePath = $localZip

            $sumsBeside = Join-Path (Split-Path -Parent $ZipPath) 'checksums.txt'
            if (Test-Path -LiteralPath $sumsBeside -PathType Leaf) {
                $expected = Get-ChecksumForFile -ChecksumText (Get-Content -LiteralPath $sumsBeside -Raw) -FileName $zipName
                $actual = (Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash.ToLowerInvariant()
                if ($expected -and $actual -ne $expected) {
                    throw ("Checksum mismatch for {0}. Expected {1}, got {2}." -f $zipName, $expected, $actual)
                }
                if ($expected) { Write-Step ("Checksum OK ({0})" -f $expected.Substring(0, 16)) }
            }
            else {
                Write-Step 'No checksums.txt beside the zip; installing unverified.'
            }
        }
        else {
            Write-Step 'Resolving release...'
            $release     = Get-ReleaseMetadata -Version $Version
            $tag         = $release.tag_name
            $realVersion = $tag.TrimStart('v')
            $zipName     = "AppPackager-$realVersion.zip"
            Write-Step ("Release {0} ({1})" -f $tag, $zipName)

            $zipUrl = Get-ReleaseAssetUrl -Release $release -Name $zipName
            $sumUrl = Get-ReleaseAssetUrl -Release $release -Name 'checksums.txt'

            $archivePath = Join-Path $work $zipName
            Write-Step 'Downloading package...'
            Invoke-InstallerDownloadFile -Url $zipUrl -OutFile $archivePath

            Write-Step 'Verifying checksum...'
            $sumPath = Join-Path $work 'checksums.txt'
            Invoke-InstallerDownloadFile -Url $sumUrl -OutFile $sumPath
            $expected = Get-ChecksumForFile -ChecksumText (Get-Content -LiteralPath $sumPath -Raw) -FileName $zipName
            if (-not $expected) { throw ("checksums.txt lists no SHA-256 for {0}." -f $zipName) }

            $actual = (Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash.ToLowerInvariant()
            if ($actual -ne $expected) {
                throw ("Checksum mismatch for {0}. Expected {1}, got {2}." -f $zipName, $expected, $actual)
            }
            Write-Step ("Checksum OK ({0})" -f $expected.Substring(0, 16))
        }

        $stage = Join-Path $work 'stage'
        Expand-Archive -LiteralPath $archivePath -DestinationPath $stage -Force

        # A local zip carries no release tag; the extracted script header
        # names the version being installed.
        if ([string]::IsNullOrWhiteSpace($realVersion)) {
            $realVersion = ''
            $entry = Join-Path $stage 'start-apppackager.ps1'
            if (Test-Path -LiteralPath $entry) {
                foreach ($headerLine in (Get-Content -LiteralPath $entry -TotalCount 80)) {
                    if ($headerLine -match '^\s*Version\s*:\s*([0-9][0-9\.]*[0-9])\s*$') { $realVersion = $Matches[1]; break }
                }
            }
        }

        $existing = Test-Path -LiteralPath $InstallPath
        $backup   = Join-Path $work 'state'
        $preserved = @()

        if ($existing) {
            if (-not (Test-AppPackagerFolder -Path $InstallPath) -and -not $Force) {
                $anything = @(Get-ChildItem -LiteralPath $InstallPath -Force -ErrorAction SilentlyContinue)
                if ($anything.Count -gt 0) {
                    throw ("{0} exists, is not empty, and holds no AppPackager install. Re-run with -Force to replace it." -f $InstallPath)
                }
            }

            $preserved = @(Get-PreservedStateFile -Root $InstallPath)
            if ($preserved.Count -gt 0) {
                Write-Step ("Preserving {0} user-state file(s)..." -f $preserved.Count)
                foreach ($relative in $preserved) {
                    $dest = Join-Path $backup $relative
                    New-Item -ItemType Directory -Path (Split-Path -Parent $dest) -Force | Out-Null
                    Copy-Item -LiteralPath (Join-Path $InstallPath $relative) -Destination $dest -Force
                }
            }

            Write-Step 'Removing previous version...'
            Get-ChildItem -LiteralPath $InstallPath -Force | Remove-Item -Recurse -Force
        }
        else {
            New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
        }

        Write-Step ("Installing to {0}..." -f $InstallPath)
        Get-ChildItem -LiteralPath $stage -Force | Copy-Item -Destination $InstallPath -Recurse -Force

        if ($preserved.Count -gt 0) {
            Write-Step 'Restoring user state...'
            foreach ($relative in $preserved) {
                $dest = Join-Path $InstallPath $relative
                New-Item -ItemType Directory -Path (Split-Path -Parent $dest) -Force | Out-Null
                Copy-Item -LiteralPath (Join-Path $backup $relative) -Destination $dest -Force
            }
        }

        # Defence in depth: an alternate data stream inherited from a proxy or
        # policy-rewritten download would block module imports at run time.
        Get-ChildItem -LiteralPath $InstallPath -Recurse -File -ErrorAction SilentlyContinue |
            Unblock-File -ErrorAction SilentlyContinue

        Write-Step ''
        Write-Step ("AppPackager {0} installed." -f $realVersion)
        Write-Step ("Launch: {0}" -f (Join-Path $InstallPath 'start-apppackager.ps1'))
    }
    finally {
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Dot-sourcing loads the helpers for testing without performing an install.
if ($MyInvocation.InvocationName -ne '.') {
    Invoke-Install -InstallPath $InstallPath -Version $Version -ZipPath $ZipPath -Force:$Force
}
