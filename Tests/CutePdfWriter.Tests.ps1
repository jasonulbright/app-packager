BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1') -Force
    $t = $null; $e = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-cutepdfwriter.ps1'), [ref]$t, [ref]$e)
    $script:ParseErrors = $e

    $literal = $script:Ast.Find({ param($n)
        $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
        $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
        $n.Left.VariablePath.UserPath -eq 'DetectionRegistryKey'
    }, $true)
    . ([scriptblock]::Create($literal.Extent.Text))

    $table = $script:Ast.Find({ param($n)
        $n -is [System.Management.Automation.Language.HashtableAst] -and
        @($n.KeyValuePairs | Where-Object { $_.Item1.Extent.Text -eq 'RegistryKeyRelative' }).Count -gt 0
    }, $true)
    $script:Detection = [pscustomobject](& ([scriptblock]::Create($table.Extent.Text)))
}

Describe 'CutePDF Writer detection' {
    It 'parses without errors' {
        $script:ParseErrors | Should -BeNullOrEmpty
    }

    It 'detects the existence of the 64-bit uninstall key and compares no value' {
        $det = $script:Detection
        $det.Type | Should -Be 'RegistryKey'
        $det.RegistryKeyRelative | Should -Be 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\CutePDF Writer Installation'
        $det.Is64Bit | Should -BeTrue
        $det.PSObject.Properties.Name | Should -Not -Contain 'ExpectedValue'
        $det.PSObject.Properties.Name | Should -Not -Contain 'ValueName'
    }

    It 'maps to one Intune registry existence rule' {
        $rules = @(ConvertTo-IntuneWin32Rules -Manifest ([pscustomobject]@{ Detection = $script:Detection }))
        $rules.Count | Should -Be 1
        $rules[0].keyPath | Should -Be 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\CutePDF Writer Installation'
        $rules[0].operationType | Should -Be 'exists'
        $rules[0].check32BitOn64System | Should -BeFalse
    }
}

Describe 'CutePDF Writer installer source' {
    BeforeAll {
        $script:ExpectedSignerCommonName = 'Acro Software Inc.'
        $fn = $script:Ast.Find({ param($n)
            $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Assert-CutePdfWriterSignature'
        }, $true)
        . ([scriptblock]::Create($fn.Extent.Text))
        function Write-Log { param($Message, $Level) }
    }

    It 'downloads over HTTPS' {
        $url = $script:Ast.Find({ param($n)
            $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
            $n.Left.VariablePath.UserPath -eq 'ExeDownloadUrl'
        }, $true)
        $url.Right.Extent.Text | Should -Match '^"https://'
    }

    It 'accepts a valid signature from the vendor' {
        Mock Get-AuthenticodeSignature { [pscustomobject]@{ Status = 'Valid'; SignerCertificate = [pscustomobject]@{ Subject = 'CN=Acro Software Inc., O=Acro Software Inc., C=US' } } }
        { Assert-CutePdfWriterSignature -ExePath 'x.exe' } | Should -Not -Throw
    }

    It 'refuses an unsigned file' {
        Mock Get-AuthenticodeSignature { [pscustomobject]@{ Status = 'NotSigned'; SignerCertificate = $null } }
        { Assert-CutePdfWriterSignature -ExePath 'x.exe' } | Should -Throw '*NotSigned*'
    }

    It 'refuses a valid signature from another publisher' {
        Mock Get-AuthenticodeSignature { [pscustomobject]@{ Status = 'Valid'; SignerCertificate = [pscustomobject]@{ Subject = 'CN=Other Software Inc., O=Other, C=US' } } }
        { Assert-CutePdfWriterSignature -ExePath 'x.exe' } | Should -Throw '*not by*'
    }

    It 'refuses a subject that only contains the vendor name' {
        Mock Get-AuthenticodeSignature { [pscustomobject]@{ Status = 'Valid'; SignerCertificate = [pscustomobject]@{ Subject = 'CN=Acro Software Inc. Fake, O=Other, C=US' } } }
        { Assert-CutePdfWriterSignature -ExePath 'x.exe' } | Should -Throw '*not by*'
    }
}
