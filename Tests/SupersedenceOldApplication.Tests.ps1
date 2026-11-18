#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The older-application override: "Use this as the old application".

.DESCRIPTION
    The specification of App Packager Desktop (docs/specs/supersedence.md, test 6) describes the list and its rules; the code of
    that application does not build it. The cases here are the rules of the specification: only an application without the
    marker of the line is offered, the title and the version are shown and confirmed first, the marker is written only after the
    confirmation, an application that holds keywords of its own is never rewritten, and the lookup then finds the linked
    application like any other application of the line. The site is mocked.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    $script:Marker = 'APD-Line:catalog:vendor-app|default|x64'
    # The ConfigMgr cmdlets exist only with the console, so the module scope gets stubs that the cases mock.
    InModuleScope AppPackagerCommon {
        function script:Get-CMApplication { param($Name, [switch]$DisableWildcardHandling, $ErrorAction) }
        function script:Set-CMApplication { param($InputObject, $Keyword, $ErrorAction) }
        function script:Get-CMDeploymentType { param($ApplicationName, $DeploymentTypeName, [switch]$DisableWildcardHandling, $ErrorAction) }
    }

    function New-TestSiteApp {
        param([string]$Title, [string]$Version, [bool]$Expired = $false, $Keywords = @())
        [pscustomobject]@{ Title = $Title; SoftwareVersion = $Version; IsExpired = $Expired; Keywords = $Keywords }
    }
    function Get-Candidates {
        param([object[]]$Apps, [string]$NewTitle = 'Vendor App 2.0', [string]$NewVersion = '2.0.0')
        InModuleScope AppPackagerCommon -Parameters @{ Apps = $Apps; NewTitle = $NewTitle; NewVersion = $NewVersion } {
            Get-SupersedenceOldApplicationCandidate -Applications $Apps -Marker 'APD-Line:catalog:vendor-app|default|x64' -NewTitle $NewTitle -NewVersion $NewVersion
        }
    }
}

AfterAll {
    InModuleScope AppPackagerCommon {
        foreach ($n in 'Get-CMApplication', 'Set-CMApplication', 'Get-CMDeploymentType') { Remove-Item -LiteralPath "function:script:$n" -ErrorAction SilentlyContinue }
    }
}

Describe 'The list of applications that can be linked' {
    It 'offers an application without the marker and without keywords' {
        $rows = @(Get-Candidates @(New-TestSiteApp 'Vendor App 1.0' '1.0.0'))
        $rows.Count | Should -Be 1
        $rows[0].State | Should -Be 'Linkable'
        $rows[0].CanLink | Should -BeTrue
    }

    It 'lists the linkable applications first, the newest version first' {
        $rows = @(Get-Candidates @(
                (New-TestSiteApp 'Vendor App retired' '1.9.0' $true), (New-TestSiteApp 'Vendor App 1.0' '1.0.0'), (New-TestSiteApp 'Vendor App 1.5' '1.5.0')))
        @($rows | ForEach-Object { $_.Title }) | Should -Be @('Vendor App 1.5', 'Vendor App 1.0', 'Vendor App retired')
    }

    It 'never lists an application with the title of the new one' {
        @(Get-Candidates @(New-TestSiteApp 'vendor app 2.0' '1.0.0')).Count | Should -Be 0
    }

    It 'marks a retired application and does not offer it' {
        $row = @(Get-Candidates @(New-TestSiteApp 'Vendor App 1.0' '1.0.0' $true))[0]
        $row.State | Should -Be 'Retired'
        $row.CanLink | Should -BeFalse
    }

    It 'does not offer a version that cannot be read and says why' {
        $row = @(Get-Candidates @(New-TestSiteApp 'Vendor App old' 'banana'))[0]
        $row.State | Should -Be 'VersionUnreadable'
        $row.Reason | Should -Match 'banana'
        $row.CanLink | Should -BeFalse
    }

    It 'does not offer <Version> as the old application of 2.0.0, because the lookup would not choose it' -ForEach @(@{ Version = '2.0.0' }, @{ Version = '3.0.0' }) {
        $row = @(Get-Candidates @(New-TestSiteApp 'Vendor App 3' $Version))[0]
        $row.State | Should -Be 'NotOlder'
        $row.CanLink | Should -BeFalse
    }

    It 'does not offer an application that has keywords of its own' {
        $row = @(Get-Candidates @(New-TestSiteApp 'Vendor App 1.0' '1.0.0' $false @('finance', 'vendor')))[0]
        $row.State | Should -Be 'OwnKeywords'
        $row.CanLink | Should -BeFalse
        $row.CanClear | Should -BeFalse
    }

    It 'does not offer an application that holds the marker of another line' {
        $row = @(Get-Candidates @(New-TestSiteApp 'Vendor App 1.0' '1.0.0' $false @('APD-Line:catalog:other-app|default|x64')))[0]
        $row.State | Should -Be 'OtherLine'
        $row.CanLink | Should -BeFalse
    }

    It 'does not offer the marker of another profile or architecture of the same title' {
        @(Get-Candidates @(New-TestSiteApp 'A' '1.0.0' $false @('APD-Line:catalog:vendor-app|pilot|x64')))[0].State | Should -Be 'OtherLine'
        @(Get-Candidates @(New-TestSiteApp 'B' '1.0.0' $false @('APD-Line:catalog:vendor-app|default|x86')))[0].State | Should -Be 'OtherLine'
    }

    It 'shows an application that already holds the marker as linked, and it can be cleared' {
        $row = @(Get-Candidates @(New-TestSiteApp 'Vendor App 1.0' '1.0.0' $false @($script:Marker)))[0]
        $row.State | Should -Be 'Linked'
        $row.CanLink | Should -BeFalse
        $row.CanClear | Should -BeTrue
    }

    It 'cannot clear a marker that sits next to keywords of an administrator' {
        $row = @(Get-Candidates @(New-TestSiteApp 'Vendor App 1.0' '1.0.0' $false @($script:Marker, 'finance')))[0]
        $row.State | Should -Be 'Linked'
        $row.CanClear | Should -BeFalse
    }

    It 'does not offer an application whose keywords could not be read' {
        $row = @(Get-Candidates @(New-TestSiteApp 'Vendor App 1.0' '1.0.0' $false $null))[0]
        $row.State | Should -Be 'KeywordsUnreadable'
        $row.CanLink | Should -BeFalse
    }

    It 'gives a title without a version no old application (the new version cannot be read)' {
        (@(Get-Candidates -Apps @(New-TestSiteApp 'Vendor App 1.0' '1.0.0') -NewVersion 'banana')[0]).CanLink | Should -BeFalse
    }

    It 'gives a title with no applications an empty list' {
        @(Get-Candidates @()).Count | Should -Be 0
    }
}

Describe 'The lookup honors a linked application' {
    It 'does not choose an application without the marker, and chooses it once the marker is written' {
        $unlinked = @([pscustomobject]@{ CiId = '1'; Title = 'Vendor App 1.0'; SoftwareVersion = '1.0.0'; IsExpired = $false; DeploymentTypes = @() })
        # The lookup reads the site by the marker, so an application without it is not in the list that the lookup gets.
        (Select-SupersedenceOldApplication -Applications @() -NewTitle 'Vendor App 2.0' -NewVersion '2.0.0').Old | Should -BeNullOrEmpty
        (Select-SupersedenceOldApplication -Applications $unlinked -NewTitle 'Vendor App 2.0' -NewVersion '2.0.0').Old.CiId | Should -Be '1'
    }

    It 'finds an application by the marker that the link wrote into the stored XML' {
        Mock Get-CMApplication -ModuleName AppPackagerCommon {
            @([pscustomobject]@{ CI_ID = 1; LocalizedDisplayName = 'Vendor App 1.0'; SoftwareVersion = '1.0.0'; IsExpired = $false
                    SDMPackageXML = '<AppMgmtDigest><Application><DisplayInfo><Info><Tags><Tag>APD-Line:catalog:vendor-app|default|x64</Tag></Tags></Info></DisplayInfo></Application></AppMgmtDigest>' }
                [pscustomobject]@{ CI_ID = 2; LocalizedDisplayName = 'Vendor App Other'; SoftwareVersion = '1.5.0'; IsExpired = $false; SDMPackageXML = '<AppMgmtDigest/>' })
        }
        Mock Get-CMDeploymentType -ModuleName AppPackagerCommon { [pscustomobject]@{ CI_ID = 9; LocalizedDisplayName = $ApplicationName + ' - x64' } }
        $apps = InModuleScope AppPackagerCommon { Get-SupersedenceLineApplication -Marker 'APD-Line:catalog:vendor-app|default|x64' -NewApplicationName 'Vendor App 2.0' }
        @($apps | ForEach-Object { $_.Title }) | Should -Be @('Vendor App 1.0')
    }
}

Describe 'The read of the site for the list' {
    BeforeEach {
        Mock Get-CMApplication -ModuleName AppPackagerCommon {
            $global:OldAppPattern = $Name
            @([pscustomobject]@{ LocalizedDisplayName = 'Vendor App 1.0'; SoftwareVersion = '1.0.0'; IsExpired = $false; SDMPackageXML = '<x/>' },
                [pscustomobject]@{ LocalizedDisplayName = 'Vendor App 1.5'; SoftwareVersion = '1.5.0'; IsExpired = $false; SDMPackageXML = '<y/>' })
        }
        Mock Get-SupersedenceApplicationKeyword -ModuleName AppPackagerCommon { if ($Application.SoftwareVersion -eq '1.5.0') { throw 'cannot read' } else { @() } }
    }
    AfterEach { Remove-Variable -Name OldAppPattern -Scope Global -ErrorAction SilentlyContinue }

    It 'asks for a name that contains the text, and treats the wildcard characters of the text as literal' {
        $rows = InModuleScope AppPackagerCommon { Get-SupersedenceOldApplicationCandidateFromSite -Marker 'APD-Line:catalog:vendor-app|default|x64' -NewTitle 'Vendor App 2.0' -NewVersion '2.0.0' -NameContains 'App [1]*' }
        $global:OldAppPattern | Should -Be '*App `[1`]`**'
        @($rows).Count | Should -Be 2
    }

    It 'lists an application whose keywords cannot be read as not linkable' {
        $rows = @(InModuleScope AppPackagerCommon { Get-SupersedenceOldApplicationCandidateFromSite -Marker 'APD-Line:catalog:vendor-app|default|x64' -NewTitle 'Vendor App 2.0' -NewVersion '2.0.0' -NameContains 'Vendor' })
        ($rows | Where-Object { $_.Title -eq 'Vendor App 1.5' }).State | Should -Be 'KeywordsUnreadable'
        ($rows | Where-Object { $_.Title -eq 'Vendor App 1.0' }).State | Should -Be 'Linkable'
    }

    It 'refuses an empty search text and reads nothing' {
        { InModuleScope AppPackagerCommon { Get-SupersedenceOldApplicationCandidateFromSite -Marker 'm' -NewTitle 'T' -NewVersion '2.0.0' -NameContains '  ' } } | Should -Throw '*part of the application name*'
        Should -Invoke Get-CMApplication -ModuleName AppPackagerCommon -Times 0 -Exactly
    }
}

Describe 'The confirmation before the marker is written' {
    It 'shows the title and the version first, names the new application and says what is not changed' {
        $text = InModuleScope AppPackagerCommon { Get-SupersedenceLinkConfirmationText -Title 'Vendor App 1.0' -Version '1.0.0' -NewTitle 'Vendor App 2.0' }
        $text | Should -Match "^Use 'Vendor App 1\.0' version 1\.0\.0 as the old application of 'Vendor App 2\.0'\?"
        $text | Should -Match 'identity marker'
        $text | Should -Match 'deployment types, detection and deployments are not changed'
    }

    It 'says that the next publish no longer finds an application whose marker is removed' {
        (InModuleScope AppPackagerCommon { Get-SupersedenceClearConfirmationText -Title 'Vendor App 1.0' -Version '1.0.0' }) | Should -Match 'no longer finds it'
    }

    It 'writes nothing without the confirmation' {
        Mock Get-CMApplication -ModuleName AppPackagerCommon { [pscustomobject]@{ SDMPackageXML = '<x/>' } }
        Mock Get-SupersedenceApplicationKeyword -ModuleName AppPackagerCommon { @() }
        Mock Set-CMApplicationIdentityMarker -ModuleName AppPackagerCommon { $null }
        $result = InModuleScope AppPackagerCommon { Set-SupersedenceOldApplication -ApplicationName 'Vendor App 1.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' }
        $result.Ok | Should -BeFalse
        Should -Invoke Set-CMApplicationIdentityMarker -ModuleName AppPackagerCommon -Times 0 -Exactly
        Should -Invoke Get-CMApplication -ModuleName AppPackagerCommon -Times 0 -Exactly
    }

    It 'removes nothing without the confirmation' {
        Mock Get-CMApplication -ModuleName AppPackagerCommon { [pscustomobject]@{ SDMPackageXML = '<x/>' } }
        Mock Set-CMApplication -ModuleName AppPackagerCommon {}
        $result = InModuleScope AppPackagerCommon { Clear-SupersedenceOldApplication -ApplicationName 'Vendor App 1.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' }
        $result.Ok | Should -BeFalse
        Should -Invoke Set-CMApplication -ModuleName AppPackagerCommon -Times 0 -Exactly
    }
}

Describe 'Linking an application' {
    BeforeEach {
        Mock Write-Log -ModuleName AppPackagerCommon {}
        Mock Get-CMApplication -ModuleName AppPackagerCommon { [pscustomobject]@{ LocalizedDisplayName = $Name; SDMPackageXML = '<x/>' } }
        Mock Set-CMApplicationIdentityMarker -ModuleName AppPackagerCommon { $null }
    }

    It 'writes the marker of the line by the name of the application after the confirmation' {
        Mock Get-SupersedenceApplicationKeyword -ModuleName AppPackagerCommon { @() }
        $result = InModuleScope AppPackagerCommon { Set-SupersedenceOldApplication -ApplicationName 'Vendor App 1.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Confirmed }
        $result.Ok | Should -BeTrue
        Should -Invoke Set-CMApplicationIdentityMarker -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter { $ApplicationName -eq 'Vendor App 1.0' -and $Marker -eq 'APD-Line:catalog:vendor-app|default|x64' }
    }

    It 'refuses an application that has keywords of its own and writes nothing' {
        Mock Get-SupersedenceApplicationKeyword -ModuleName AppPackagerCommon { @('finance') }
        $result = InModuleScope AppPackagerCommon { Set-SupersedenceOldApplication -ApplicationName 'Vendor App 1.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Confirmed }
        $result.Ok | Should -BeFalse
        $result.Message | Should -Match 'keywords of its own'
        Should -Invoke Set-CMApplicationIdentityMarker -ModuleName AppPackagerCommon -Times 0 -Exactly
    }

    It 'refuses an application that holds the marker of another line and writes nothing' {
        Mock Get-SupersedenceApplicationKeyword -ModuleName AppPackagerCommon { @('APD-Line:catalog:other-app|default|x64') }
        $result = InModuleScope AppPackagerCommon { Set-SupersedenceOldApplication -ApplicationName 'Vendor App 1.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Confirmed }
        $result.Ok | Should -BeFalse
        $result.Message | Should -Match 'another line'
        Should -Invoke Set-CMApplicationIdentityMarker -ModuleName AppPackagerCommon -Times 0 -Exactly
    }

    It 'does not write again when the application already holds the marker' {
        Mock Get-SupersedenceApplicationKeyword -ModuleName AppPackagerCommon { @('APD-Line:catalog:vendor-app|default|x64') }
        $result = InModuleScope AppPackagerCommon { Set-SupersedenceOldApplication -ApplicationName 'Vendor App 1.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Confirmed }
        $result.Ok | Should -BeTrue
        Should -Invoke Set-CMApplicationIdentityMarker -ModuleName AppPackagerCommon -Times 0 -Exactly
    }

    It 'reports the reason when the keywords cannot be read, and writes nothing' {
        Mock Get-SupersedenceApplicationKeyword -ModuleName AppPackagerCommon { throw 'the digest is damaged' }
        $result = InModuleScope AppPackagerCommon { Set-SupersedenceOldApplication -ApplicationName 'Vendor App 1.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Confirmed }
        $result.Ok | Should -BeFalse
        $result.Message | Should -Match 'could not be read'
        Should -Invoke Set-CMApplicationIdentityMarker -ModuleName AppPackagerCommon -Times 0 -Exactly
    }

    It 'reports a failed write with its reason' {
        Mock Get-SupersedenceApplicationKeyword -ModuleName AppPackagerCommon { @() }
        Mock Set-CMApplicationIdentityMarker -ModuleName AppPackagerCommon { 'access denied' }
        $result = InModuleScope AppPackagerCommon { Set-SupersedenceOldApplication -ApplicationName 'Vendor App 1.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Confirmed }
        $result.Ok | Should -BeFalse
        $result.Message | Should -Match 'access denied'
    }

    It 'reports an application that is not on the site' {
        Mock Get-CMApplication -ModuleName AppPackagerCommon { $null }
        $result = InModuleScope AppPackagerCommon { Set-SupersedenceOldApplication -ApplicationName 'Gone' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Confirmed }
        $result.Ok | Should -BeFalse
        $result.Message | Should -Match 'not on the site'
    }
}

Describe 'Removing a link' {
    BeforeEach {
        Mock Write-Log -ModuleName AppPackagerCommon {}
        Mock Get-CMApplication -ModuleName AppPackagerCommon { [pscustomobject]@{ LocalizedDisplayName = $Name; SDMPackageXML = '<x/>' } }
        Mock Set-CMApplication -ModuleName AppPackagerCommon {}
    }

    It 'empties the keywords of an application whose only keyword is the marker of the line' {
        Mock Get-SupersedenceApplicationKeyword -ModuleName AppPackagerCommon { @('APD-Line:catalog:vendor-app|default|x64') }
        $result = InModuleScope AppPackagerCommon { Clear-SupersedenceOldApplication -ApplicationName 'Vendor App 1.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Confirmed }
        $result.Ok | Should -BeTrue
        Should -Invoke Set-CMApplication -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter { $Keyword -eq '' }
    }

    It 'does not touch an application that holds keywords of an administrator next to the marker' {
        Mock Get-SupersedenceApplicationKeyword -ModuleName AppPackagerCommon { @('APD-Line:catalog:vendor-app|default|x64', 'finance') }
        $result = InModuleScope AppPackagerCommon { Clear-SupersedenceOldApplication -ApplicationName 'Vendor App 1.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Confirmed }
        $result.Ok | Should -BeFalse
        Should -Invoke Set-CMApplication -ModuleName AppPackagerCommon -Times 0 -Exactly
    }

    It 'does not touch an application that holds the marker of another line' {
        Mock Get-SupersedenceApplicationKeyword -ModuleName AppPackagerCommon { @('APD-Line:catalog:other-app|default|x64') }
        $result = InModuleScope AppPackagerCommon { Clear-SupersedenceOldApplication -ApplicationName 'Vendor App 1.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Confirmed }
        $result.Ok | Should -BeFalse
        Should -Invoke Set-CMApplication -ModuleName AppPackagerCommon -Times 0 -Exactly
    }

    It 'reports a failed write' {
        Mock Get-SupersedenceApplicationKeyword -ModuleName AppPackagerCommon { @('APD-Line:catalog:vendor-app|default|x64') }
        Mock Set-CMApplication -ModuleName AppPackagerCommon { throw 'access denied' }
        $result = InModuleScope AppPackagerCommon { Clear-SupersedenceOldApplication -ApplicationName 'Vendor App 1.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Confirmed }
        $result.Ok | Should -BeFalse
        $result.Message | Should -Match 'access denied'
    }
}

Describe 'The grid menu and the dialog' {
    BeforeAll {
        $script:ShellText = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'start-apppackager.ps1') -Raw
    }

    It 'adds the menu item to the context menu of the grid and opens the dialog for the selected row' {
        $script:ShellText | Should -Match '\$menuOldApplication\.Header = "Use as old application\.\.\."'
        $script:ShellText | Should -Match '\$contextMenu\.Items\.Add\(\$menuOldApplication\)'
        $script:ShellText | Should -Match '(?s)\$menuOldApplication\.Add_Click\(\{.{0,120}Open-OldApplicationDialog -Owner \$window -Row \$row'
    }

    It 'takes the identity marker and the title from the newest staged build, as the publish does' {
        $script:ShellText | Should -Match '(?s)function Open-OldApplicationDialog.*Find-NewestStageManifestForPackager.*Get-SupersedenceIdentityMarker -Manifest \$manifest -FallbackApplicationId.*Get-PackagedApplicationName'
    }

    It 'asks before it writes and goes back to the earlier location after the site cmdlets' {
        $script:ShellText | Should -Match '(?s)function Show-OldApplicationDialog.*Get-SupersedenceLinkConfirmationText.*Show-ThemedMessage.*YesNo'
        $script:ShellText | Should -Match '(?s)function Open-OldApplicationDialog.*Push-Location.*Pop-Location'
    }
}

Describe 'Open-OldApplicationDialog site calls' {
    # The dialog calls the search, link and unlink blocks while the site drive is not the current location. Each block connects first.
    BeforeAll {
        $script:StartText = [System.IO.File]::ReadAllText((Join-Path $script:RepoRoot 'start-apppackager.ps1'))
    }

    It 'connects to the site in the search, the link and the unlink block' -ForEach @(
        @{ Name = 'reader' }, @{ Name = 'linker' }, @{ Name = 'unlinker' }
    ) {
        $match = [regex]::Match($script:StartText, '(?s)\$' + $Name + ' = \{(.*?)\}\.GetNewClosure\(\)')
        $match.Success | Should -BeTrue
        $match.Groups[1].Value | Should -Match 'Connect-CMSite'
    }
}
