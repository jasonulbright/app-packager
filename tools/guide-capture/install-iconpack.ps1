<#
.SYNOPSIS
    Installs the latest icon pack into the lab copy of the app, so that every guide image shows icons.

.DESCRIPTION
    Runs on the lab computer (Invoke-GuideCapture sends it through the PowerShell session). It does what the
    app does for Options, Tools, Icon pack: read the latest release of jasonulbright/app-packager-icons,
    download icon-pack.zip and checksums.txt, check the SHA-256 of the zip, and extract it into
    <AppRoot>\Packagers\Icons. The script is safe to run again: it replaces the icons of the pack and leaves
    other files in the folder alone. It stops with an error when the release has no pack, the checksum does
    not match, or the folder holds no icon after the extraction, so a guide run never takes pictures
    without icons.
    The script returns one object with Tag, IconCount and Folder.

.PARAMETER AppRoot
    The lab copy of the app (the folder that holds Packagers).
#>
[CmdletBinding()]
param([Parameter(Mandatory)][string]$AppRoot)
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$ProgressPreference = 'SilentlyContinue'
$repo = 'jasonulbright/app-packager-icons'
$headers = @{ 'User-Agent' = 'AppPackager-IconPack' }
$destination = Join-Path (Join-Path $AppRoot 'Packagers') 'Icons'
$scratch = Join-Path ([IO.Path]::GetTempPath()) ('apicons-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $scratch -Force | Out-Null
try {
    $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/releases/latest" -Headers $headers -UseBasicParsing
    $pack = @($release.assets | Where-Object { $_.name -eq 'icon-pack.zip' })[0]
    $sums = @($release.assets | Where-Object { $_.name -eq 'checksums.txt' })[0]
    if (-not $pack -or -not $sums) { throw "The latest release of $repo has no icon-pack.zip and checksums.txt." }
    $zip = Join-Path $scratch 'icon-pack.zip'
    $sumsFile = Join-Path $scratch 'checksums.txt'
    Invoke-WebRequest -Uri $pack.browser_download_url -OutFile $zip -Headers $headers -UseBasicParsing
    Invoke-WebRequest -Uri $sums.browser_download_url -OutFile $sumsFile -Headers $headers -UseBasicParsing
    $expected = $null
    foreach ($line in (Get-Content -LiteralPath $sumsFile)) {
        $m = [regex]::Match($line.Trim(), '^([0-9a-fA-F]{64})\s+\*?(.+)$')
        if ($m.Success -and $m.Groups[2].Value.Trim() -eq 'icon-pack.zip') { $expected = $m.Groups[1].Value }
    }
    if (-not $expected) { throw 'checksums.txt has no entry for icon-pack.zip.' }
    if ((Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash -ne $expected.ToUpperInvariant()) { throw 'The icon pack checksum did not match. Nothing was extracted.' }
    $extracted = Join-Path $scratch 'extracted'
    Expand-Archive -LiteralPath $zip -DestinationPath $extracted -Force
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
    Get-ChildItem -LiteralPath $extracted -Force | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $destination -Recurse -Force }
    $count = @(Get-ChildItem -LiteralPath $destination -File | Where-Object { $_.Extension -in '.png', '.ico' }).Count
    if ($count -lt 1) { throw 'The icon pack folder holds no icon after the extraction.' }
    [pscustomobject]@{ Tag = [string]$release.tag_name; IconCount = $count; Folder = $destination }
}
finally { Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue }
