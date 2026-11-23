<#
.SYNOPSIS
    Lists the guide images that are older than the last change to the screen they show (H-5).

.DESCRIPTION
    Reads the map GuideImages.sources.json (guide image id to the source files of its screen) and compares two git commits for each image:
    the last commit that touched docs/images/<id>.png and the last commit that touched one of the sources of the image. An image is Stale
    when the source commit is newer. Use the list to pick the images of a capture run instead of retaking all of them.
    A commit whose subject matches a regular expression of 'ignoreSubjects' in the map (for example a release commit) never makes an image stale.
    Status values: Fresh, Stale, Unknown (a shallow clone, or an image that has no commit) and Unmapped (no rule of the map matches the id).
    The script reports. It exits 1 only with -FailOnStale and at least one Stale or Unmapped image.

.EXAMPLE
    powershell -NoProfile -File tools/guide-capture/Get-StaleGuideImages.ps1
    pwsh -NoProfile -File tools/guide-capture/Get-StaleGuideImages.ps1 -Format Json -All
#>
[CmdletBinding()]
param(
    [string]$Root,
    [string]$MapPath,
    [string]$ImagesDir = 'docs/images',
    [ValidateSet('Text', 'Json', 'Markdown')][string]$Format = 'Text',
    [switch]$All,
    [switch]$FailOnStale
)

$ErrorActionPreference = 'Stop'

function Invoke-Git {
    param([string]$Repository, [string[]]$Arguments)
    # Windows PowerShell 5.1 turns a stderr line of a native command into an error under 'Stop'.
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { $output = & git -C $Repository @Arguments 2>&1 } finally { $ErrorActionPreference = $previous }
    if ($LASTEXITCODE -ne 0) { throw "git $($Arguments -join ' ') failed: $output" }
    return @($output | ForEach-Object { [string]$_ })
}

function ConvertTo-WildcardRegex {
    param([string]$Pattern)
    return '^' + ([regex]::Escape($Pattern) -replace '\\\*', '.*') + '$'
}

function Get-LastCommit {
    param([string]$Repository, [string[]]$Pathspec, [string[]]$IgnoreSubjects = @())
    $filter = @()
    foreach ($pattern in $IgnoreSubjects) { $filter += "--grep=$pattern" }
    if ($filter.Count -gt 0) { $filter += '--invert-grep' }
    $line = Invoke-Git $Repository (@('log', '-1', '--format=%H%x09%ct%x09%s') + $filter + @('--') + $Pathspec) | Select-Object -First 1
    if (-not $line) { return $null }
    $parts = ([string]$line).Split("`t", 3)
    return [pscustomobject]@{ Sha = $parts[0]; Time = [DateTimeOffset]::FromUnixTimeSeconds([long]$parts[1]); Subject = $parts[2] }
}

if (-not $Root) { $Root = (& git -C $PSScriptRoot rev-parse --show-toplevel).Trim() }
if (-not $MapPath) { $MapPath = Join-Path $Root 'tools/guide-capture/GuideImages.sources.json' }

$map = Get-Content -LiteralPath $MapPath -Raw | ConvertFrom-Json
$shallow = ((Invoke-Git $Root @('rev-parse', '--is-shallow-repository')) -join '').Trim() -eq 'true'
$ids = Get-ChildItem -LiteralPath (Join-Path $Root $ImagesDir) -Filter '*.png' | ForEach-Object { $_.BaseName } | Sort-Object

$results = foreach ($id in $ids) {
    $sources = [System.Collections.Generic.List[string]]::new()
    $matched = $false
    foreach ($rule in $map.rules) {
        if (@($rule.ids | Where-Object { $id -match (ConvertTo-WildcardRegex $_) }).Count -gt 0) {
            $matched = $true
            foreach ($source in $rule.sources) { if (-not $sources.Contains($source)) { $sources.Add($source) } }
        }
    }
    foreach ($source in $map.shared) { if (-not $sources.Contains($source)) { $sources.Add($source) } }

    $row = [ordered]@{ Id = $id; Status = 'Fresh'; ImageCommit = $null; ImageDate = $null; SourceCommit = $null; SourceDate = $null; ChangedFiles = @(); Reason = '' }
    if (-not $matched) {
        $row.Status = 'Unmapped'; $row.Reason = 'No rule of the map matches this id.'
    }
    elseif ($shallow) {
        $row.Status = 'Unknown'; $row.Reason = 'Shallow clone: the history is incomplete. Fetch the full history (git fetch --unshallow).'
    }
    else {
        $imageCommit = Get-LastCommit $Root @("$ImagesDir/$id.png")
        $sourceCommit = Get-LastCommit $Root $sources.ToArray() @($map.ignoreSubjects)
        if (-not $imageCommit) {
            $row.Status = 'Unknown'; $row.Reason = 'The image file has no commit.'
        }
        else {
            $row.ImageCommit = $imageCommit.Sha.Substring(0, 8); $row.ImageDate = $imageCommit.Time.ToString('yyyy-MM-dd HH:mm')
            if ($sourceCommit) {
                $row.SourceCommit = $sourceCommit.Sha.Substring(0, 8); $row.SourceDate = $sourceCommit.Time.ToString('yyyy-MM-dd HH:mm')
                if ($sourceCommit.Time -gt $imageCommit.Time) {
                    $row.Status = 'Stale'
                    $row.Reason = $sourceCommit.Subject
                    $row.ChangedFiles = @(Invoke-Git $Root (@('diff', '--name-only', "$($imageCommit.Sha)..HEAD", '--') + $sources.ToArray()) | ForEach-Object { [string]$_ })
                }
            }
        }
    }
    [pscustomobject]$row
}

$shown = if ($All) { @($results) } else { @($results | Where-Object { $_.Status -ne 'Fresh' }) }
$count = { param($status) @($results | Where-Object { $_.Status -eq $status }).Count }
$summary = "Guide images: $($results.Count) total, $(& $count 'Stale') stale, $(& $count 'Unknown') unknown, $(& $count 'Unmapped') unmapped, $(& $count 'Fresh') fresh."

switch ($Format) {
    'Json' { [pscustomobject]@{ Summary = $summary; Images = $shown } | ConvertTo-Json -Depth 5 }
    'Markdown' {
        "### Guide image staleness"
        ""
        $summary
        if ($shown.Count -gt 0) {
            ""
            "| Image | Status | Image commit | Source commit | Changed or reason |"
            "|---|---|---|---|---|"
            foreach ($r in $shown) {
                $detail = if ($r.ChangedFiles.Count -gt 0) { ($r.ChangedFiles | Select-Object -First 3 | ForEach-Object { '`' + (Split-Path $_ -Leaf) + '`' }) -join ', ' } else { $r.Reason }
                "| $($r.Id) | $($r.Status) | $($r.ImageCommit) $($r.ImageDate) | $($r.SourceCommit) $($r.SourceDate) | $detail |"
            }
        }
    }
    default {
        $summary
        foreach ($r in $shown) {
            $line = '{0,-9} {1}' -f $r.Status, $r.Id
            if ($r.Status -eq 'Stale') { $line += "  (image $($r.ImageDate), source $($r.SourceDate): $($r.Reason))" }
            elseif ($r.Reason) { $line += "  ($($r.Reason))" }
            $line
        }
    }
}

if ($FailOnStale -and (@($results | Where-Object { $_.Status -in 'Stale', 'Unmapped' }).Count -gt 0)) { exit 1 }
