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
    endpoints or download installers, so they are opt-in. -ThrottleLimit runs
    that many packagers at a time.

.EXAMPLE
    .\Tests\Invoke-PackagerSmoke.ps1

.EXAMPLE
    .\Tests\Invoke-PackagerSmoke.ps1 -IncludeLatest -LatestTimeoutSec 90 -ThrottleLimit 6

.EXAMPLE
    .\Tests\Invoke-PackagerSmoke.ps1 -Packager package-git.ps1 -IncludeLatest -IncludeStage -DownloadRoot C:\temp\ap
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

    [int]$ThrottleLimit = 1,

    [switch]$Json,

    [switch]$Matrix,

    [string]$MatrixPath
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

# ---------------------------------------------------------------------------
# Level-1 catalog inventory (-Matrix)
# ---------------------------------------------------------------------------

function Get-PackagerAstExpression {
    # Unwraps the pipeline/command/paren layers the parser puts around a value
    # so a hashtable value node can be read directly.
    param($Node)

    $current = $Node
    for ($i = 0; $i -lt 8 -and $null -ne $current; $i++) {
        if ($current -is [System.Management.Automation.Language.PipelineAst]) {
            if ($current.PipelineElements.Count -ne 1) { return $current }
            $current = $current.PipelineElements[0]
            continue
        }
        if ($current -is [System.Management.Automation.Language.CommandExpressionAst]) { $current = $current.Expression; continue }
        if ($current -is [System.Management.Automation.Language.ParenExpressionAst]) { $current = $current.Pipeline; continue }
        if ($current -is [System.Management.Automation.Language.StatementBlockAst]) {
            if ($current.Statements.Count -ne 1) { return $current }
            $current = $current.Statements[0]
            continue
        }
        break
    }
    return $current
}

function ConvertFrom-PackagerAstValue {
    <#
        Literal value of an AST node, or $null when the packager computes it at
        run time. A computed value is what a dry generation would resolve; the
        inventory records the shape, not the run-time content.
    #>
    param($Node, [int]$Depth = 0)

    if ($null -eq $Node -or $Depth -gt 8) { return $null }
    $node = Get-PackagerAstExpression -Node $Node

    if ($node -is [System.Management.Automation.Language.HashtableAst]) {
        $map = [ordered]@{}
        foreach ($pair in $node.KeyValuePairs) {
            $key = ConvertFrom-PackagerAstValue -Node $pair.Item1 -Depth ($Depth + 1)
            if ($null -eq $key) { continue }
            $map[[string]$key] = ConvertFrom-PackagerAstValue -Node $pair.Item2 -Depth ($Depth + 1)
        }
        return $map
    }
    if ($node -is [System.Management.Automation.Language.ArrayLiteralAst]) {
        return @($node.Elements | ForEach-Object { ConvertFrom-PackagerAstValue -Node $_ -Depth ($Depth + 1) })
    }
    if ($node -is [System.Management.Automation.Language.ArrayExpressionAst]) {
        $items = New-Object System.Collections.ArrayList
        foreach ($statement in $node.SubExpression.Statements) {
            $value = ConvertFrom-PackagerAstValue -Node $statement -Depth ($Depth + 1)
            if ($value -is [object[]]) { foreach ($item in $value) { [void]$items.Add($item) } }
            else { [void]$items.Add($value) }
        }
        return $items.ToArray()
    }
    if ($node -is [System.Management.Automation.Language.ConstantExpressionAst]) { return $node.Value }
    if ($node -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) { return $null }
    return $null
}

function Resolve-PackagerVariableHashtable {
    # First hashtable assigned to a variable name anywhere in the script. A
    # packager that branches assigns the shape it shares with every branch
    # first, which is the shape the inventory records.
    param([Parameter(Mandatory)]$Ast, [Parameter(Mandatory)][string]$Name)

    foreach ($assignment in $Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)) {
        $target = Get-PackagerAstExpression -Node $assignment.Left
        if (-not ($target -is [System.Management.Automation.Language.VariableExpressionAst])) { continue }
        if ($target.VariablePath.UserPath -ne $Name) { continue }
        $value = ConvertFrom-PackagerAstValue -Node $assignment.Right
        if ($value -is [System.Collections.IDictionary]) { return $value }
    }
    return $null
}

function Get-PackagerDetectionBlocks {
    <#
        Every Detection block authored in a packager: the top-level one plus
        one per deployment-type variant. Read from the source because an
        offline sweep cannot run a real Stage.
    #>
    param([Parameter(Mandatory)]$Ast)

    $blocks = New-Object System.Collections.ArrayList
    foreach ($hashtable in $Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.HashtableAst] }, $true)) {
        foreach ($pair in $hashtable.KeyValuePairs) {
            $key = ConvertFrom-PackagerAstValue -Node $pair.Item1
            if ([string]$key -ne 'Detection') { continue }
            $value = ConvertFrom-PackagerAstValue -Node $pair.Item2
            if (-not ($value -is [System.Collections.IDictionary])) {
                # A packager that composes its detection above the manifest
                # assigns a variable here; read that variable's hashtable.
                $node = Get-PackagerAstExpression -Node $pair.Item2
                if ($node -is [System.Management.Automation.Language.VariableExpressionAst]) {
                    $value = Resolve-PackagerVariableHashtable -Ast $Ast -Name $node.VariablePath.UserPath
                }
            }
            if ($value -is [System.Collections.IDictionary]) { [void]$blocks.Add($value) }
        }
    }
    return $blocks.ToArray()
}

function Get-DetectionShape {
    # Detector type, operators, and compound shape of one Detection block.
    param([AllowNull()]$Detection)

    $shape = [pscustomobject]@{
        Type       = ''
        ClauseTypes = @()
        Operators  = @()
        Compound   = $false
        Connector  = ''
        GroupSizes = ''
    }
    if ($null -eq $Detection) { return $shape }

    $type = [string]$Detection['Type']
    if ([string]::IsNullOrWhiteSpace($type)) { $type = 'RegistryKeyValue' }
    $shape.Type = $type

    $operators = New-Object System.Collections.Generic.List[string]
    $clauseTypes = New-Object System.Collections.Generic.List[string]
    if ($Detection.Contains('Operator') -and $Detection['Operator']) { $operators.Add([string]$Detection['Operator']) }

    if ($type -eq 'Compound') {
        $shape.Compound = $true
        if ($Detection.Contains('Connector')) { $shape.Connector = [string]$Detection['Connector'] }
        if ($Detection.Contains('GroupSizes') -and $Detection['GroupSizes']) {
            $shape.GroupSizes = (@($Detection['GroupSizes']) -join '|')
        }
        foreach ($clause in @($Detection['Clauses'])) {
            if (-not ($clause -is [System.Collections.IDictionary])) { continue }
            $clauseType = [string]$clause['Type']
            if ([string]::IsNullOrWhiteSpace($clauseType)) { $clauseType = 'RegistryKeyValue' }
            $clauseTypes.Add($clauseType)
            if ($clause.Contains('Operator') -and $clause['Operator']) { $operators.Add([string]$clause['Operator']) }
        }
    }
    else {
        $clauseTypes.Add($type)
    }

    $shape.ClauseTypes = @($clauseTypes | Sort-Object -Unique)
    $shape.Operators = @($operators | Sort-Object -Unique)
    return $shape
}

function ConvertTo-MatrixManifest {
    # The smallest manifest Get-IntuneCompatibilityFindings needs: the authored
    # Detection block plus the architecture field the findings look at.
    param([AllowNull()]$Detection)

    $json = (@{ Detection = $Detection; Architecture = 'x64' } | ConvertTo-Json -Depth 10)
    return ($json | ConvertFrom-Json)
}

function Test-LauncherPolicyToken {
    <#
        A launcher string that relaxes the execution policy. PowerShell binds a
        parameter by unique prefix, so -ep and every prefix of
        -ExecutionPolicy reach the same parameter; -enc reaches
        -EncodedCommand.
    #>
    param([AllowEmptyString()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $hits = New-Object System.Collections.Generic.List[string]
    foreach ($raw in ($Text -split '[\s=:]+')) {
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }
        if ($raw[0] -ne '-' -and $raw[0] -ne '/') { continue }
        $name = $raw.TrimStart('-', '/').Trim('"', "'").ToLowerInvariant()
        if ([string]::IsNullOrEmpty($name)) { continue }
        if ('executionpolicy'.StartsWith($name) -or $name -eq 'ep') { $hits.Add($raw) }
        elseif ('encodedcommand'.StartsWith($name)) { $hits.Add($raw) }
    }
    return @($hits | Sort-Object -Unique)
}

function Get-SignedLauncherProbe {
    <#
        Generates the deployment wrappers once with SignDeployment on, in a
        child process so the sweep host never imports the product modules.
        The sweep proves the generated strings; it never signs anything.
    #>
    param([Parameter(Mandatory)][string]$PowerShellExe, [Parameter(Mandatory)][string]$RepoRoot)

    $probe = @'
$ErrorActionPreference = 'Stop'
$env:APP_PACKAGER_SIGNING = '{"SignDeployment":true,"SignDetection":false,"SignRequirements":false,"RequireDeployment":false,"RequireDetection":false,"RequireRequirements":false,"CertificateThumbprint":"","StoreLocation":"CurrentUser","TimestampServer":"","TimestampRequired":false,"HashAlgorithm":"SHA256"}'
Import-Module (Join-Path '__ROOT__' 'Packagers\AppPackagerCommon.psd1') -Force
$out = Join-Path ([IO.Path]::GetTempPath()) ('ap-launch-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $out -Force | Out-Null
try {
    Write-ContentWrappers -OutputPath $out -InstallPs1Content 'exit 0' -UninstallPs1Content 'exit 0' 6>$null | Out-Null
    $rebootOut = Join-Path $out 'reboot'
    New-Item -ItemType Directory -Path $rebootOut -Force | Out-Null
    Write-ContentWrappers -OutputPath $rebootOut -InstallPs1Content 'exit 0' -UninstallPs1Content 'exit 0' -InstallBatExitCode '3010' 6>$null | Out-Null
    $result = [ordered]@{
        InstallBat    = [IO.File]::ReadAllText((Join-Path $out 'install.bat'))
        UninstallBat  = [IO.File]::ReadAllText((Join-Path $out 'uninstall.bat'))
        RebootBat     = [IO.File]::ReadAllText((Join-Path $rebootOut 'install.bat'))
        LauncherX64   = (New-DeploymentLauncherCommand -Script 'install.ps1' -Signed $true).CommandLine
        LauncherX86   = (New-DeploymentLauncherCommand -Script 'install.ps1' -Signed $true -ScriptHost x86).CommandLine
        BatBodyX64    = (New-DeploymentLauncherCommand -Script 'install.ps1' -Signed $true).BatBody
        UnsignedBat   = (New-DeploymentLauncherCommand -Script 'install.ps1' -Signed $false).BatInvoke
    }
    ([pscustomobject]$result | ConvertTo-Json -Depth 4)
}
finally { Remove-Item -LiteralPath $out -Recurse -Force -ErrorAction SilentlyContinue }
'@
    $probe = $probe.Replace('__ROOT__', $RepoRoot)
    $probePath = Join-Path ([System.IO.Path]::GetTempPath()) ('ap-launch-probe-' + [guid]::NewGuid().ToString('N') + '.ps1')
    [System.IO.File]::WriteAllText($probePath, $probe, (New-Object System.Text.UTF8Encoding($false)))
    try {
        $result = Invoke-ChildProcess -FileName $PowerShellExe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $probePath) -TimeoutSec 180
        if ($result.TimedOut -or $result.ExitCode -ne 0) {
            return [pscustomobject]@{ Ok = $false; Reason = ("Launcher probe failed: {0}" -f ($result.StdErr, $result.StdOut -join ' ').Trim()); Strings = @() }
        }
        $json = ($result.StdOut -split "`r?`n" | Where-Object { $_ } ) -join "`n"
        $start = $json.IndexOf('{')
        if ($start -lt 0) { return [pscustomobject]@{ Ok = $false; Reason = 'Launcher probe produced no JSON.'; Strings = @() } }
        $data = $json.Substring($start) | ConvertFrom-Json
        $strings = New-Object System.Collections.ArrayList
        foreach ($property in $data.PSObject.Properties) {
            if ($property.Name -eq 'UnsignedBat') { continue }
            foreach ($line in ([string]$property.Value -split "`r?`n")) {
                [void]$strings.Add([pscustomobject]@{ Name = $property.Name; Line = $line })
            }
        }
        return [pscustomobject]@{ Ok = $true; Reason = ''; Strings = $strings.ToArray(); Data = $data }
    }
    finally { Remove-Item -LiteralPath $probePath -Force -ErrorAction SilentlyContinue }
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

    $job = [pscustomobject]@{ Key = 'single'; FileName = $FileName; ArgumentList = $ArgumentList; TimeoutSec = $TimeoutSec }
    return (Invoke-ChildProcessBatch -Jobs @($job) -ThrottleLimit 1)['single']
}

function Invoke-ChildProcessBatch {
    <#
    .SYNOPSIS
        Runs child processes up to ThrottleLimit at a time and returns a
        hashtable of Key -> { ExitCode, TimedOut, StdOut, StdErr }.
    .DESCRIPTION
        Both streams are drained by ReadToEndAsync tasks. A scriptblock bound
        to the DataReceived events would run on a thread-pool thread with no
        runspace and terminate the host without a message; Start-Process
        -PassThru reports no ExitCode in Windows PowerShell 5.1.
    #>
    param(
        [Parameter(Mandatory)][object[]]$Jobs,
        [int]$ThrottleLimit = 1
    )

    if ($ThrottleLimit -lt 1) { $ThrottleLimit = 1 }
    $results = @{}
    $pending = New-Object System.Collections.Generic.Queue[object]
    foreach ($job in $Jobs) { $pending.Enqueue($job) }
    $running = New-Object System.Collections.Generic.List[object]

    while ($pending.Count -gt 0 -or $running.Count -gt 0) {
        while ($pending.Count -gt 0 -and $running.Count -lt $ThrottleLimit) {
            $job = $pending.Dequeue()
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $job.FileName
            $psi.Arguments = (($job.ArgumentList | ForEach-Object { ConvertTo-CommandLineArgument $_ }) -join ' ')
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $proc = New-Object System.Diagnostics.Process
            $proc.StartInfo = $psi
            [void]$proc.Start()
            $running.Add([pscustomobject]@{
                Job     = $job
                Proc    = $proc
                Out     = $proc.StandardOutput.ReadToEndAsync()
                Err     = $proc.StandardError.ReadToEndAsync()
                Started = [DateTime]::UtcNow
            })
        }

        Start-Sleep -Milliseconds 200

        for ($i = $running.Count - 1; $i -ge 0; $i--) {
            $entry = $running[$i]
            $timedOut = $false
            if (-not $entry.Proc.HasExited) {
                if (([DateTime]::UtcNow - $entry.Started).TotalSeconds -lt $entry.Job.TimeoutSec) { continue }
                $timedOut = $true
                # Kill the tree: a packager's own children (curl.exe, the Office
                # Deployment Tool) would otherwise outlive the timed-out host.
                try { & taskkill.exe /PID $entry.Proc.Id /T /F 2>&1 | Out-Null } catch { }
                try { if (-not $entry.Proc.HasExited) { $entry.Proc.Kill() } } catch { }
            }
            $entry.Proc.WaitForExit()
            $results[$entry.Job.Key] = [pscustomobject]@{
                ExitCode = if ($timedOut) { $null } else { $entry.Proc.ExitCode }
                TimedOut = $timedOut
                StdOut   = [string]$entry.Out.Result
                StdErr   = [string]$entry.Err.Result
            }
            $entry.Proc.Dispose()
            $running.RemoveAt($i)
        }
    }
    return $results
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

function Test-PackagerSignedLauncher {
    <#
        Signed deployment mode must leave no execution-policy relaxation in
        anything the endpoint runs. Two halves: the shared wrapper strings the
        packager's Write-ContentWrappers call produces (generated once with
        SignDeployment on), and any launcher string the packager authors
        itself.
    #>
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [Parameter(Mandatory)]$Ast,
        [Parameter(Mandatory)]$Probe
    )

    if (-not $Probe.Ok) {
        return New-SmokeResult -Script $File.Name -Check 'Signed launcher policy' -Status Fail -Detail $Probe.Reason
    }

    $findings = New-Object System.Collections.Generic.List[string]
    foreach ($entry in $Probe.Strings) {
        if ($entry.Line -notmatch '(?i)powershell') { continue }
        foreach ($token in (Test-LauncherPolicyToken -Text $entry.Line)) {
            $findings.Add(("shared {0} carries {1}" -f $entry.Name, $token))
        }
    }

    foreach ($literal in $Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)) {
        $text = [string]$literal.Value
        if ($text -notmatch '(?i)powershell') { continue }
        foreach ($token in (Test-LauncherPolicyToken -Text $text)) {
            $findings.Add(("line {0} authors a launcher carrying {1}" -f $literal.Extent.StartLineNumber, $token))
        }
    }

    if ($findings.Count -gt 0) {
        return New-SmokeResult -Script $File.Name -Check 'Signed launcher policy' -Status Fail -Detail (($findings | Sort-Object -Unique) -join '; ')
    }
    return New-SmokeResult -Script $File.Name -Check 'Signed launcher policy' -Status Pass -Detail 'No execution-policy relaxation in signed deployment mode.'
}

function Get-PackagerMatrixRow {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [Parameter(Mandatory)]$Ast,
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)]$Results,
        [bool]$IntuneAvailable
    )

    $meta = Get-PackagerHeaderMetadata -Path $File.FullName
    $blocks = @(Get-PackagerDetectionBlocks -Ast $Ast)
    $base = if ($blocks.Count -gt 0) { $blocks[0] } else { $null }
    $shape = Get-DetectionShape -Detection $base

    $allOperators = New-Object System.Collections.Generic.List[string]
    $allTypes = New-Object System.Collections.Generic.List[string]
    foreach ($block in $blocks) {
        $blockShape = Get-DetectionShape -Detection $block
        foreach ($operator in $blockShape.Operators) { $allOperators.Add($operator) }
        foreach ($type in $blockShape.ClauseTypes) { $allTypes.Add($type) }
    }

    $launcher = if ($Source -match 'Test-PsadtLayout') { 'PSADT' }
        elseif ($Source -match 'Write-ContentWrappers') { 'Write-ContentWrappers' }
        else { 'other' }

    $intuneVerdict = 'NotEvaluated'
    $intuneDetail = ''
    if ($IntuneAvailable -and $base) {
        try {
            $findings = @(Get-IntuneCompatibilityFindings -Manifest (ConvertTo-MatrixManifest -Detection $base))
            $blocking = @($findings | Where-Object { $_.Severity -eq 'Blocking' })
            $review = @($findings | Where-Object { $_.Severity -eq 'Review' })
            $intuneVerdict = if ($blocking.Count -gt 0) { 'Unsupported' } elseif ($review.Count -gt 0) { 'Review' } else { 'Ready' }
            $intuneDetail = (@($blocking + $review) | ForEach-Object { $_.Code }) -join '|'
        }
        catch {
            $intuneVerdict = 'Unsupported'
            $intuneDetail = 'FindingsThrew'
        }
    }
    elseif (-not $base) {
        $intuneVerdict = 'NoAuthoredDetection'
    }

    $row = [ordered]@{
        PackagerId        = $File.BaseName
        Vendor            = [string]$meta.Vendor
        App               = [string]$meta.App
        DetectorType      = $shape.Type
        DetectorClauses   = (@($allTypes | Sort-Object -Unique) -join '|')
        Operators         = (@($allOperators | Sort-Object -Unique) -join '|')
        Compound          = $shape.Compound
        Connector         = $shape.Connector
        GroupSizes        = $shape.GroupSizes
        DetectionBlocks   = $blocks.Count
        Variants          = [bool]($Source -match 'SupportsVariants')
        InstallModes      = [bool]($Source -match 'SupportsInstallModes')
        LauncherSource    = $launcher
        ExecutionPolicyInContent = ''
        IntuneVerdict     = $intuneVerdict
        IntuneFindings    = $intuneDetail
    }

    $policyResult = @($Results | Where-Object { $_.Check -eq 'Signed launcher policy' }) | Select-Object -First 1
    $row['ExecutionPolicyInContent'] = if ($policyResult -and $policyResult.Status -eq 'Fail') { $policyResult.Detail } else { 'none' }

    foreach ($result in $Results) {
        $row[('Check_' + ($result.Check -replace '[^A-Za-z0-9]', ''))] = $result.Status
    }
    return [pscustomobject]$row
}

function Get-LatestSmokeArguments {
    param([Parameter(Mandatory)][System.IO.FileInfo]$File)
    @(
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
}

function ConvertTo-LatestSmokeResult {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [Parameter(Mandatory)][object]$Result
    )

    $result = $Result
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

function Get-StageSmokeArguments {
    param([Parameter(Mandatory)][System.IO.FileInfo]$File)
    @(
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
}

function ConvertTo-StageSmokeResult {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [Parameter(Mandatory)][object]$Result
    )

    $result = $Result
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
    # powershell.exe -File hands a comma list over as one string.
    $wanted = @($Packager | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ })
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

$launcherProbe = Get-SignedLauncherProbe -PowerShellExe $powershellExe -RepoRoot (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path

$intuneAvailable = $false
if ($Matrix) {
    try {
        Import-Module (Join-Path $resolvedRoot 'AppPackagerCommon.psd1') -Force -ErrorAction Stop
        $intuneAvailable = [bool](Get-Command -Name Get-IntuneCompatibilityFindings -ErrorAction SilentlyContinue)
    }
    catch {
        Write-Warning ("Intune verdict unavailable ({0}); matrix rows report NotEvaluated." -f $_.Exception.Message)
    }
}

$matrixRows = New-Object System.Collections.ArrayList
foreach ($file in $files) {
    $fileResults = New-Object System.Collections.Generic.List[object]
    foreach ($result in (Test-PackagerSyntaxAndContract -File $file)) { $fileResults.Add($result) }

    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$parseErrors)
    $fileResults.Add((Test-PackagerSignedLauncher -File $file -Ast $ast -Probe $launcherProbe))

    foreach ($result in $fileResults) { $allResults.Add($result) }

    if ($Matrix) {
        $source = Get-Content -LiteralPath $file.FullName -Raw -ErrorAction Stop
        [void]$matrixRows.Add((Get-PackagerMatrixRow -File $file -Ast $ast -Source $source -Results $fileResults -IntuneAvailable $intuneAvailable))
    }
}

# Live checks run as child processes, ThrottleLimit at a time; the results
# keep the packager order.
if ($IncludeLatest) {
    $jobs = @($files | Where-Object { $SkipLatest -notcontains $_.Name } | ForEach-Object {
        [pscustomobject]@{ Key = $_.Name; FileName = $powershellExe; ArgumentList = (Get-LatestSmokeArguments -File $_); TimeoutSec = $LatestTimeoutSec }
    })
    $latestResults = if ($jobs.Count -gt 0) { Invoke-ChildProcessBatch -Jobs $jobs -ThrottleLimit $ThrottleLimit } else { @{} }
    foreach ($file in $files) {
        if ($SkipLatest -contains $file.Name) {
            $allResults.Add((New-SmokeResult -Script $file.Name -Check 'Latest live' -Status Skip -Detail 'Skipped by -SkipLatest.'))
        }
        else {
            $allResults.Add((ConvertTo-LatestSmokeResult -File $file -Result $latestResults[$file.Name]))
        }
    }
}

if ($IncludeStage) {
    $jobs = @($files | Where-Object { $SkipStage -notcontains $_.Name } | ForEach-Object {
        [pscustomobject]@{ Key = $_.Name; FileName = $powershellExe; ArgumentList = (Get-StageSmokeArguments -File $_); TimeoutSec = $StageTimeoutSec }
    })
    $stageResults = if ($jobs.Count -gt 0) { Invoke-ChildProcessBatch -Jobs $jobs -ThrottleLimit $ThrottleLimit } else { @{} }
    foreach ($file in $files) {
        if ($SkipStage -contains $file.Name) {
            $allResults.Add((New-SmokeResult -Script $file.Name -Check 'Stage live' -Status Skip -Detail 'Skipped by -SkipStage.'))
        }
        else {
            $allResults.Add((ConvertTo-StageSmokeResult -File $file -Result $stageResults[$file.Name]))
        }
    }
}

# Non-ASCII guard: PowerShell 5.1 reads a BOM-less file as ANSI, so any
# non-ASCII character in shipped .ps1/.psm1/.psd1 risks mojibake or parse
# damage depending on which tool last saved the file. Runs before the
# summary so -Json output stays a single JSON document.
$nonAscii = @()
foreach ($f in Get-ChildItem (Join-Path $PSScriptRoot '..') -Recurse -Include *.ps1,*.psm1,*.psd1 -File | Where-Object FullName -notmatch '\\Tests\\|\\Logs\\|\\Icons\\') {
    if ([regex]::IsMatch([System.IO.File]::ReadAllText($f.FullName), '[^\x00-\x7F]')) { $nonAscii += $f.Name }
}
foreach ($name in $nonAscii) {
    $allResults.Add([pscustomobject]@{ Script = $name; Check = 'Pure ASCII'; Status = 'Fail'; Detail = 'Shipped PowerShell file contains non-ASCII characters.' })
}

$resultArray = @($allResults.ToArray())

# The catalog matrix is one level-1 inventory row per packager. Tests/out is
# gitignored, so the default output never reaches a commit.
$matrixWritten = ''
if ($Matrix) {
    if ([string]::IsNullOrWhiteSpace($MatrixPath)) {
        $MatrixPath = Join-Path $PSScriptRoot 'out\catalog-matrix.csv'
    }
    $matrixFolder = Split-Path -Parent $MatrixPath
    if ($matrixFolder -and -not (Test-Path -LiteralPath $matrixFolder)) {
        New-Item -ItemType Directory -Path $matrixFolder -Force | Out-Null
    }
    @($matrixRows.ToArray()) | Export-Csv -LiteralPath $MatrixPath -NoTypeInformation -Encoding ASCII
    $matrixWritten = $MatrixPath
}

$summary = [pscustomobject]@{
    Packagers = $files.Count
    Checks    = $resultArray.Count
    Passed    = @($resultArray | Where-Object { $_.Status -eq 'Pass' }).Count
    Failed    = @($resultArray | Where-Object { $_.Status -eq 'Fail' }).Count
    Skipped   = @($resultArray | Where-Object { $_.Status -eq 'Skip' }).Count
    MatrixPath = $matrixWritten
    Results   = $resultArray
}

if ($Json) {
    $summary | ConvertTo-Json -Depth 5
}
else {
    Write-Host ("Packager smoke: {0} script(s), {1} check(s), {2} passed, {3} failed, {4} skipped" -f $summary.Packagers, $summary.Checks, $summary.Passed, $summary.Failed, $summary.Skipped)
    Write-Host $(if ($nonAscii.Count -gt 0) { "Non-ASCII guard: FAILED - " + ($nonAscii -join ", ") } else { "Non-ASCII guard: all shipped PowerShell files are pure ASCII" })
    if ($matrixWritten) {
        $rows = @($matrixRows.ToArray())
        Write-Host ("Catalog matrix: {0} row(s) -> {1}" -f $rows.Count, $matrixWritten)
        Write-Host ("  detectors: {0} native single, {1} compound, {2} script, {3} without an authored block" -f
            @($rows | Where-Object { -not $_.Compound -and $_.DetectorType -ne 'Script' -and $_.DetectionBlocks -gt 0 }).Count,
            @($rows | Where-Object { $_.Compound }).Count,
            @($rows | Where-Object { $_.DetectorType -eq 'Script' }).Count,
            @($rows | Where-Object { $_.DetectionBlocks -eq 0 }).Count)
        Write-Host ("  Intune: {0} Ready, {1} Review, {2} Unsupported, {3} not evaluated" -f
            @($rows | Where-Object { $_.IntuneVerdict -eq 'Ready' }).Count,
            @($rows | Where-Object { $_.IntuneVerdict -eq 'Review' }).Count,
            @($rows | Where-Object { $_.IntuneVerdict -eq 'Unsupported' }).Count,
            @($rows | Where-Object { $_.IntuneVerdict -notin @('Ready', 'Review', 'Unsupported') }).Count)
    }
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


