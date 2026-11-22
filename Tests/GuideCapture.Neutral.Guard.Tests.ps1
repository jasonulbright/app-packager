# Guards for the capture tool: the lab folder carries the repository name (not the name of the folder that runs the capture),
# and every shot moves the real mouse away first so that no tooltip stays in an image.
BeforeAll {
    $script:Dir = Join-Path (Split-Path -Parent $PSScriptRoot) 'tools\guide-capture'
    $script:Invoke = Get-Content -LiteralPath (Join-Path $script:Dir 'Invoke-GuideCapture.ps1') -Raw
    $script:Lib = Get-Content -LiteralPath (Join-Path $script:Dir 'guilib.ps1') -Raw
}
Describe 'Capture tool: neutral lab folder and no stray tooltip' {
    It 'names the app with a literal, never with the name of the folder that runs the capture' {
        $script:Invoke | Should -Not -Match 'Split-Path\s+\$repo\s+-Leaf'
        $script:Invoke | Should -Match "(?m)^\`$app = '[a-z0-9-]+'\s*$"
    }
    It 'moves the real mouse away before every shot that is not a hover shot' {
        $script:Lib | Should -Match 'function Move-MouseAway'
        $script:Lib | Should -Match 'mouse_event\(1,\s*3,\s*3'
        $shot = [regex]::Match($script:Lib, '(?s)function Shot\(.*?\r?\n\}').Value
        $shot | Should -Match 'Move-MouseAway'
    }
    It 'waits until no tooltip shows and logs a tooltip that stays' {
        $script:Lib | Should -Match 'ControlType\]::ToolTip'
        $script:Lib | Should -Match 'TOOLTIP still shown'
    }
}
