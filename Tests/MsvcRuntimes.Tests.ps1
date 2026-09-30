Describe 'Visual C++ v14 Redistributable detection in <Script>' -ForEach @(
    @{ Script = 'package-msvcruntimesx64.ps1'; Installer = 'vc_redist.x64.exe'; Folder = '%SystemRoot%\System32' }
    @{ Script = 'package-msvcruntimesx86.ps1'; Installer = 'vc_redist.x86.exe'; Folder = '%SystemRoot%\SysWOW64' }
) {
    BeforeAll {
        $t = $null; $e = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot "..\Packagers\$Script"), [ref]$t, [ref]$e)
        if ($e) { throw ($e.Message -join '; ') }
        # A key that names a variable reads as the literal string the packager
        # assigns to it.
        $literals = @{}
        foreach ($a in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)) {
            if ($a.Right -is [System.Management.Automation.Language.CommandExpressionAst] -and $a.Right.Expression -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                $literals[$a.Left.VariablePath.UserPath] = $a.Right.Expression.Value
            }
        }
        function Get-TableValue {
            param($Table, [string]$Key, [hashtable]$Literals)
            $pair = @($Table.KeyValuePairs | Where-Object { $_.Item1.Extent.Text -eq $Key })
            if ($pair.Count -eq 0) { return $null }
            $v = $pair[0].Item2
            if ($v -is [System.Management.Automation.Language.PipelineAst]) { $v = $v.PipelineElements[0] }
            if ($v -is [System.Management.Automation.Language.CommandExpressionAst]) { $v = $v.Expression }
            if ($v -is [System.Management.Automation.Language.StringConstantExpressionAst]) { return $v.Value }
            if ($v -is [System.Management.Automation.Language.VariableExpressionAst]) {
                $name = $v.VariablePath.UserPath
                if ($name -in @('true', 'false')) { return ($name -eq 'true') }
                if ($Literals.ContainsKey($name)) { return $Literals[$name] }
                return '$' + $name
            }
            return $v.Extent.Text
        }
        $manifest = $ast.Find({ param($n)
            $n -is [System.Management.Automation.Language.HashtableAst] -and
            @($n.KeyValuePairs | Where-Object { $_.Item1.Extent.Text -eq 'Detection' }).Count -gt 0
        }, $true)
        if (-not $manifest) { throw "No stage manifest with a Detection block in $Script" }
        $detectionValue = ($manifest.KeyValuePairs | Where-Object { $_.Item1.Extent.Text -eq 'Detection' }).Item2
        $detection = $detectionValue.Find({ param($n) $n -is [System.Management.Automation.Language.HashtableAst] }, $true)
    }

    It 'stages one installer for one architecture' {
        Get-TableValue $manifest 'InstallerFile' $literals | Should -Be $Installer
        Get-TableValue $manifest 'InstallerFiles' $literals | Should -BeNullOrEmpty
        Get-TableValue $manifest 'InstallArgs' $literals | Should -Be '/install /quiet /norestart'
    }

    It 'compares the vcruntime140.dll file version in <Folder> through the 64-bit view' {
        Get-TableValue $detection 'Type' $literals | Should -Be 'File'
        Get-TableValue $detection 'FilePath' $literals | Should -Be $Folder
        Get-TableValue $detection 'FileName' $literals | Should -Be 'vcruntime140.dll'
        Get-TableValue $detection 'PropertyType' $literals | Should -Be 'Version'
        Get-TableValue $detection 'Operator' $literals | Should -Be 'GreaterEquals'
        Get-TableValue $detection 'ExpectedValue' $literals | Should -Be '$quadVersion'
        Get-TableValue $detection 'Is64Bit' $literals | Should -BeTrue
    }
}
