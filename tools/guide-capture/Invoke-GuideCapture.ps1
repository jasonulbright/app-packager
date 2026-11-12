<#
.SYNOPSIS
    Regenerates every image of the user guide with one command.

.DESCRIPTION
    Copies this app to a lab computer, runs the capture steps in an interactive
    session of a lab account (scheduled task, interactive logon), copies the PNG
    files back, and removes the task and the copy. Only the images that already
    exist in docs\images are replaced. A failed step stops the command before any
    image is replaced.

    One capture runs at a time on a lab computer. The script takes the lock file
    C:\guide\capture.lock (New-Item fails when the file exists) and waits while another
    capture holds it. The wait never spins: each retry follows a sleep. The script touches
    the lock every 5 seconds while its capture runs, so a live capture never looks stale.
    A lock that nobody touched for 60 minutes is stale. The script renames it to
    capture.lock.stale.<token>, reads the age of the renamed file again, and deletes it only
    when it is still stale. A lock that was touched in the meantime goes back to its name.
    An empty lock file (the holder has not written its token yet) is a lock like any other.
    The script works in its own folder C:\guide\app-packager and its own task
    GuideShots-app-packager. It removes the lock and its own files at the end. It never
    removes the folder C:\guide and never signs out a session.

    The lab computer needs a Configuration Manager console, a reachable SMS
    Provider and PowerShell remoting for the account in the credential script.
    Read the steps before you run them: they open dialogs of buttons that change
    the site, and they answer No or Cancel. They never confirm a change.

.PARAMETER CredentialScript
    A script that sets $LabCredential (a PSCredential). Keep it outside the repository.

.PARAMETER Target
    The lab computer name or address.

.PARAMETER LogonUser
    The domain account that has the interactive session on the lab computer.

.PARAMETER SiteCode
    The site code that the capture types in Options.

.PARAMETER SmsProvider
    The SMS Provider name that the capture types in Options.

.PARAMETER Minutes
    The time limit of each run.

.PARAMETER LockWaitMinutes
    How long to wait for the capture lock of another capture before the script stops.

.PARAMETER AddNew
    Also copy images that are not in docs\images yet. Use it for the first capture of a new page.

.PARAMETER OnlyCveChrome
    Run only steps-cve-chrome.ps1: the Chrome CVE images (cve-chrome-*.png and light-cve-chrome-*.png).
    The steps write a packager history with Chrome 157.0.8103.40 on the site and 158.0.8127.37 as the
    latest version, then run Check CVEs against the live NVD. The default run does not include them.

.EXAMPLE
    .\tools\guide-capture\Invoke-GuideCapture.ps1 -CredentialScript C:\lab\credentials.local.ps1 -Target CLIENT01 -LogonUser contoso\labadmin -SmsProvider cm01.contoso.com
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', 'CredentialScript', Justification = 'The parameter is the path of a script, not a password.')]
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$CredentialScript,
    [Parameter(Mandatory)][string]$Target,
    [Parameter(Mandatory)][string]$LogonUser,
    [string]$SiteCode = 'MCM',
    [Parameter(Mandatory)][string]$SmsProvider,
    [int]$Minutes = 15,
    [int]$LockWaitMinutes = 90,
    [switch]$AddNew,
    [switch]$OnlyCveChrome
)
$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$repo = (Resolve-Path (Join-Path $here '..\..')).Path
$app = 'app-packager'
$images = Join-Path $repo 'docs\images'
$LabCredential = $null
. $CredentialScript
if (-not ($LabCredential -is [System.Management.Automation.PSCredential])) { throw 'The credential script must set $LabCredential to a PSCredential.' }
# Own folder and own task name, so that two captures in the lab never meet. The C:\guide folder itself stays.
$root = "C:\guide\$app"
$taskName = "GuideShots-$app"
$lockPath = 'C:\guide\capture.lock'
$lockToken = [guid]::NewGuid().ToString('N')
$haveLock = $false
$settings = @{ SiteCode = $SiteCode; SMSProvider = $SmsProvider } | ConvertTo-Json
$failed = New-Object System.Collections.Generic.List[string]
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('guide-capture-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $tmp | Out-Null
$session = $null
$ok = $false
try {
    $session = New-PSSession -ComputerName $Target -Credential $LabCredential
    # The lock: New-Item with -ErrorAction Stop fails with an IOException when the file exists.
    $haveLock = Invoke-Command $session {
        param($path, $token, $staleMinutes, $waitMinutes)
        New-Item -ItemType Directory -Path (Split-Path -Path $path -Parent) -Force | Out-Null
        $deadline = (Get-Date).AddMinutes($waitMinutes)
        while ($true) {
            try {
                # The token goes in with the file, in one call. A failure of a second write could not be told from a held lock.
                $null = New-Item -ItemType File -Path $path -Value $token -ErrorAction Stop
                return $true
            }
            catch [System.IO.IOException] {
                $held = Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
                if ($held -and ((Get-Date) - $held.LastWriteTime).TotalMinutes -gt $staleMinutes) {
                    # Rename first, so that two waiting captures never delete the same lock. Then read the age of the
                    # renamed file again: the holder may have touched the lock between the first read and the rename.
                    $stale = '{0}.stale.{1}' -f $path, $token
                    try {
                        Move-Item -LiteralPath $path -Destination $stale -ErrorAction Stop
                        $moved = Get-Item -LiteralPath $stale -ErrorAction Stop
                        if (((Get-Date) - $moved.LastWriteTime).TotalMinutes -gt $staleMinutes) { Remove-Item -LiteralPath $stale -Force -ErrorAction SilentlyContinue }
                        else {
                            # The holder touched the lock. It goes back to its name. When another capture took the name in the
                            # meantime, the renamed file is deleted. Either way the lock is busy: the loop sleeps and tries again.
                            try { Move-Item -LiteralPath $stale -Destination $path -ErrorAction Stop }
                            catch { Remove-Item -LiteralPath $stale -Force -ErrorAction SilentlyContinue }
                        }
                    }
                    catch { $null = $_ }
                }
                if ((Get-Date) -gt $deadline) { return $false }
                # Every retry sleeps, so that the loop never spins.
                Start-Sleep -Seconds 15
            }
        }
    } -ArgumentList $lockPath, $lockToken, 60, $LockWaitMinutes
    if (-not $haveLock) { throw "Another capture holds $lockPath. The wait of $LockWaitMinutes minutes is over." }
    $runs = if ($OnlyCveChrome) { @('steps-cve-chrome.ps1') } else { @('steps-guide.ps1', 'steps-console-missing.ps1') }
    foreach ($steps in $runs) {
        # A fresh copy for each run. Stop only the leftovers of an earlier capture of this app.
        Invoke-Command $session {
            param($d)
            Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*$d\*" -and $_.ProcessId -ne $PID } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
            Start-Sleep -Seconds 2
            if (Test-Path $d) { Remove-Item $d -Recurse -Force }
            New-Item -ItemType Directory $d, (Join-Path $d 'ui') -Force | Out-Null
        } -ArgumentList $root
        # The app files only. The names are root names of the repository: the folder .claude (worktrees) and the folders of tests, tools and documents stay out.
        foreach ($item in Get-ChildItem -LiteralPath $repo -Force | Where-Object { $_.Name -notin '.git', '.claude', '.github', '.vs', '.vscode', '.idea', 'node_modules', 'TestOutput', 'Logs', 'History', 'Reports', 'screenshots', 'TestResults', 'docs', 'Tests', 'tools', 'Samples', 'CATALOG-PARITY.csv' -and $_.Name -notlike '*.prefs.json' -and $_.Name -notlike '*.windowstate.json' -and $_.Name -notlike '*.lastscan.json' -and $_.Name -notlike '*.suppressions.json' -and $_.Name -notlike '*.local.*' -and $_.Name -notlike '.env*' -and $_.Name -notlike 'secrets*' -and $_.Name -notlike '*.pfx' -and $_.Name -notlike '*.key' -and $_.Name -notlike '*.png' }) {
            Copy-Item -LiteralPath $item.FullName -Destination "$root\$($item.Name)" -ToSession $session -Recurse -Force
        }
        Copy-Item (Join-Path $here 'guilib.ps1') (Join-Path $root 'ui\guilib.ps1') -ToSession $session -Force
        Copy-Item (Join-Path $here 'main.ps1') (Join-Path $root 'ui\main.ps1') -ToSession $session -Force
        Copy-Item (Join-Path $here $steps) (Join-Path $root 'ui\steps.ps1') -ToSession $session -Force
        Invoke-Command $session {
            param($r, $user, $json, $task)
            Set-Content (Join-Path $r 'ui\capture.json') $json -Encoding UTF8
            Remove-Item (Join-Path $r 'ui\out') -Recurse -Force -ErrorAction SilentlyContinue
            $act = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$r\ui\main.ps1`" -GuideRoot `"$r`""
            $pr = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Highest
            Register-ScheduledTask -TaskName $task -Action $act -Principal $pr -Force | Out-Null
            Start-ScheduledTask $task
        } -ArgumentList $root, $LogonUser, $settings, $taskName
        $end = (Get-Date).AddMinutes($Minutes)
        $done = $false
        do {
            Start-Sleep 5
            # The poll also touches the lock, so that a long capture is never taken for a stale one.
            $state = Invoke-Command $session {
                param($r, $path, $task)
                (Get-Item -LiteralPath $path -ErrorAction SilentlyContinue).LastWriteTime = Get-Date
                [pscustomobject]@{ Done = (Test-Path (Join-Path $r 'ui\out\finished.flag')); Task = (Get-ScheduledTask $task -ErrorAction SilentlyContinue).State; Result = (Get-ScheduledTaskInfo $task -ErrorAction SilentlyContinue).LastTaskResult }
            } -ArgumentList $root, $lockPath, $taskName
            $done = $state.Done
            # The task is not running any more, no finished flag exists and the last result is an error: stop now.
            if (-not $done -and $state.Task -eq 'Ready' -and $state.Result -notin 0, 267009, 267011) { throw ("The capture task stopped with result {0}: {1}" -f $state.Result, $steps) }
        } while (-not $done -and (Get-Date) -lt $end)
        if (-not $done) { throw "The capture did not finish: $steps" }
        $out = Join-Path $tmp ([IO.Path]::GetFileNameWithoutExtension($steps))
        New-Item -ItemType Directory $out | Out-Null
        Copy-Item (Join-Path $root 'ui\out\*') $out -FromSession $session -Recurse -Force
        # The flag file holds finished or failed: <reason>. A failure is reported even when the log could not be written.
        $flagText = (Get-Content (Join-Path $out 'finished.flag') -Raw -ErrorAction SilentlyContinue)
        if ($flagText -and $flagText.Trim() -like 'failed:*') { $failed.Add($flagText.Trim()) }
        $log = Join-Path $out 'run.log'
        if (-not (Test-Path $log)) { throw "The capture wrote no run.log: $steps" }
        Get-Content $log | Where-Object { $_ -match '^\d\d:\d\d:\d\d\s+ERROR' } | ForEach-Object { Write-Warning $_; $failed.Add($_) }
    }
    # A failed step stops the capture before any guide image is replaced.
    if ($failed.Count) { throw ("The capture failed. No image was replaced.`n" + ($failed -join "`n")) }
    # A failure screenshot is evidence for the operator, not a guide image.
    $captured = @(Get-ChildItem $tmp -Recurse -Filter *.png | Where-Object { $_.Name -notlike 'fail-*' })
    $capturedNames = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $duplicate = New-Object System.Collections.Generic.List[string]
    foreach ($png in $captured) { if (-not $capturedNames.Add($png.Name)) { $duplicate.Add($png.Name) } }
    if ($duplicate.Count) { throw ("Two runs produced the same image. No image was replaced.`n" + ($duplicate -join "`n")) }
    # The Chrome CVE images come from the switch -OnlyCveChrome only. A full run neither makes them nor misses them.
    $missing = @(Get-ChildItem $images -Filter *.png | Where-Object { -not $capturedNames.Contains($_.Name) -and ($_.Name -notmatch '^(light-)?cve-chrome-') } | ForEach-Object { $_.Name })
    if ($missing.Count -and -not $OnlyCveChrome) { throw ("The capture did not produce these guide images. No image was replaced.`n" + ($missing -join "`n")) }
    foreach ($png in $captured) {
        if ($AddNew -or (Test-Path (Join-Path $images $png.Name))) { Copy-Item $png.FullName (Join-Path $images $png.Name) -Force }
    }
    $ok = $true
}
finally {
    # Each cleanup runs on its own, so that a failed cleanup cannot replace the original error. Order: stop the task,
    # stop the runner and the app processes of this folder, unregister the task, delete the folder, release the lock.
    if ($session) {
        try {
            Invoke-Command $session {
                param($t, $r)
                Stop-ScheduledTask -TaskName $t -ErrorAction SilentlyContinue
                Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*$r\*" -and $_.ProcessId -ne $PID } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
                Start-Sleep -Seconds 2
                Unregister-ScheduledTask -TaskName $t -Confirm:$false -ErrorAction SilentlyContinue
            } -ArgumentList $taskName, $root
        } catch { Write-Warning "Cleanup of the scheduled task failed: $($_.Exception.Message)" }
        try { Invoke-Command $session { param($r) Remove-Item $r -Recurse -Force -ErrorAction SilentlyContinue } -ArgumentList $root } catch { Write-Warning "Cleanup of $root failed: $($_.Exception.Message)" }
        # Remove the lock only when this capture took it.
        try {
            if ($haveLock) {
                Invoke-Command $session { param($path, $token) if ((Test-Path -LiteralPath $path) -and (([string](Get-Content -LiteralPath $path -Raw -ErrorAction SilentlyContinue)).Trim() -eq $token)) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue } } -ArgumentList $lockPath, $lockToken
            }
        } catch { Write-Warning "Removing the capture lock failed: $($_.Exception.Message)" }
        try { Remove-PSSession $session } catch { Write-Warning "Closing the session failed: $($_.Exception.Message)" }
    }
    if ($ok) { try { Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue } catch { Write-Warning "Cleanup of $tmp failed: $($_.Exception.Message)" } } else { Write-Warning "The captured files stay in $tmp for inspection." }
}
