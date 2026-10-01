BeforeAll {
    $root = Join-Path $PSScriptRoot '..'
    $script:Blocks = foreach ($path in @(Join-Path $root 'start-apppackager.ps1') + @(Get-ChildItem (Join-Path $root 'Packagers') -Filter '*.psm1' -File | ForEach-Object FullName)) {
        $tokens = $null; $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
        $ast.FindAll({
            param($n)
            ($n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or
             $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) -and
            $n.Value -match 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
        }, $true) | ForEach-Object {
            [pscustomobject]@{ Where = '{0}:{1}' -f (Split-Path $path -Leaf), $_.Extent.StartLineNumber; Text = $_.Value }
        }
    }
}

Describe 'XAML blocks' {
    It 'finds the inline XAML blocks' {
        @($script:Blocks).Count | Should -BeGreaterThan 10
    }

    # A panel is parsed only when it is first shown; an undeclared prefix
    # there throws out of ShowDialog and closes the app.
    It 'parses every inline XAML block as XML' {
        $failed = foreach ($b in $script:Blocks) {
            try { [void][xml]$b.Text } catch { '{0}: {1}' -f $b.Where, $_.Exception.InnerException.Message }
        }
        @($failed) | Should -BeNullOrEmpty
    }

    It 'parses every XAML file as XML' {
        $failed = foreach ($f in Get-ChildItem (Join-Path $PSScriptRoot '..') -Filter '*.xaml' -File) {
            try { [void][xml](Get-Content -LiteralPath $f.FullName -Raw) } catch { $f.Name }
        }
        @($failed) | Should -BeNullOrEmpty
    }
}
