#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    Guard tests of the names, the titles and the tooltips of the Options window.

.DESCRIPTION
    The tree labels, the section titles and the field titles are in sentence case, every field has a
    tooltip of one short sentence, no banned word is in any of these strings, and every message that
    sends the user to an Options page names a node of the tree. These cases follow the rules of
    Theme.Guard.Tests.ps1 for the windows that the repository builds from XAML.
#>

BeforeDiscovery {
    . (Join-Path $PSScriptRoot 'OptionsDefinitionHelpers.ps1')
    $script:LauncherD = Join-Path (Split-Path -Parent $PSScriptRoot) 'start-apppackager.ps1'
    $script:SectionD = @(Get-OptionsDefinitionFromLauncher -Path $script:LauncherD)
}

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    . (Join-Path $PSScriptRoot 'OptionsDefinitionHelpers.ps1')
    . (Join-Path $PSScriptRoot 'ThemeGuardHelpers.ps1')
    $script:LauncherPath = Join-Path $script:RepoRoot 'start-apppackager.ps1'
    $script:Section = @(Get-OptionsDefinitionFromLauncher -Path $script:LauncherPath)
    $script:AllField = @($script:Section | ForEach-Object { $_.Fields } | Where-Object { $_ })
    $script:ById = @{}
    foreach ($s in $script:Section) { $script:ById[[string]$s.Id] = $s }

    # Every string that the user reads: titles, labels, tooltips and descriptions, with the sections they belong to.
    $script:Text = New-Object System.Collections.Generic.List[object]
    foreach ($s in $script:Section) {
        foreach ($name in 'Title', 'Description') { if ($s[$name]) { $script:Text.Add([pscustomobject]@{ Where = [string]$s.Id; Kind = $name; Text = [string]$s[$name] }) } }
        foreach ($f in @($s.Fields | Where-Object { $_ })) {
            foreach ($name in 'Title', 'Label', 'Tip', 'Description', 'Text', 'Watermark') { if ($f[$name]) { $script:Text.Add([pscustomobject]@{ Where = [string]$f.Id; Kind = $name; Text = [string]$f[$name] }) } }
            foreach ($b in @($f.Buttons | Where-Object { $_ })) { foreach ($name in 'Label', 'Tip') { if ($b[$name]) { $script:Text.Add([pscustomobject]@{ Where = [string]$f.Id; Kind = 'Button' + $name; Text = [string]$b[$name] }) } } }
            foreach ($i in @($f.Items | Where-Object { $_ })) { if ($i -is [hashtable]) { foreach ($name in 'Label', 'Tip') { if ($i[$name]) { $script:Text.Add([pscustomobject]@{ Where = [string]$f.Id; Kind = 'Item' + $name; Text = [string]$i[$name] }) } } } }
        }
        foreach ($sf in @($s.SearchFields | Where-Object { $_ })) { $script:Text.Add([pscustomobject]@{ Where = [string]$sf.Id; Kind = 'Title'; Text = [string]$sf.Title }) }
    }
}

Describe 'The tree' {
    It 'has the approved top-level nodes in order' {
        (@($script:Section | Where-Object { -not $_.Parent } | ForEach-Object { $_.Title }) -join '|') | Should -Be 'Connections|Packaging|Applications|Updates and CVEs|Security|Tools|About'
    }

    It 'has the approved sections under <Parent>' -ForEach @(
        @{ Parent = 'connections'; Expected = 'Systems in use|ConfigMgr|WSUS|Intune|Test connection' }
        @{ Parent = 'packaging'; Expected = 'Folders|Deployment type defaults|Application title|After Package' }
        @{ Parent = 'applications'; Expected = 'Adobe Acrobat Reader|Beyond Compare 5|Citrix Workspace App|DBeaver Community|Microsoft 365 Apps|SQL Server Management Studio|TeamViewer Host|Local installer sources' }
        @{ Parent = 'cwa'; Expected = 'Store|Installation|Plugins and add-ons|Update and telemetry|Store policy|Components ADDLOCAL' }
        @{ Parent = 'updates'; Expected = 'One Click|Product filter|CVE lookup' }
        @{ Parent = 'security'; Expected = 'Script signing|GitHub API' }
    ) {
        (@($script:Section | Where-Object { $_.Parent -eq $Parent } | ForEach-Object { $_.Title }) -join '|') | Should -Be $Expected
    }

    It 'lists the applications in alphabetical order, with Local installer sources last' {
        $names = @($script:Section | Where-Object { $_.Parent -eq 'applications' } | ForEach-Object { $_.Title })
        $apps = @($names | Select-Object -First ($names.Count - 1))
        ($apps -join '|') | Should -Be (($apps | Sort-Object { $_.ToLowerInvariant() }) -join '|')
        $names[-1] | Should -Be 'Local installer sources'
    }

    It 'gives every node a unique id, a parent that comes before it, and a kind' {
        $seen = @{}
        foreach ($s in $script:Section) {
            $seen.ContainsKey([string]$s.Id) | Should -BeFalse -Because ('duplicate id ' + $s.Id)
            if ($s.Parent) { $seen.ContainsKey([string]$s.Parent) | Should -BeTrue -Because ($s.Id + ' needs its parent first') }
            $s.Kind | Should -BeIn 'Group', 'Fields', 'Panel'
            $seen[[string]$s.Id] = $true
        }
    }

    It 'gives every field an id that is unique in the window' {
        $ids = @($script:AllField | Where-Object { $_.Id } | ForEach-Object { [string]$_.Id })
        @($ids | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name }) | Should -BeNullOrEmpty
    }
}

Describe 'Names and titles' {
    It 'writes the title of <_.Title> in sentence case' -ForEach $script:SectionD {
        (Test-SentenceCase -Text ([string]$_.Title)) | Should -BeTrue
    }

    It 'writes the title of every field and group heading in sentence case, with no colon' {
        $bad = foreach ($f in $script:AllField) {
            if (-not $f.Title) { continue }
            if (([string]$f.Title).EndsWith(':') -or -not (Test-SentenceCase -Text ([string]$f.Title))) { '{0}: {1}' -f $f.Id, $f.Title }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'writes the search titles of the reused panels in sentence case' {
        $bad = foreach ($s in $script:Section) { foreach ($sf in @($s.SearchFields | Where-Object { $_ })) { if (-not (Test-SentenceCase -Text ([string]$sf.Title))) { '{0}: {1}' -f $sf.Id, $sf.Title } } }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'names ConfigMgr and never MECM or MCM in any string' {
        $bad = @($script:Text | Where-Object { $_.Text -cmatch '\b(MECM|MCM)\b' } | ForEach-Object { '{0}: {1}' -f $_.Where, $_.Text })
        $bad | Should -BeNullOrEmpty
    }

    It 'holds no banned word in any string' {
        $bad = @($script:Text | Where-Object { $_.Text -match $script:BannedPattern } | ForEach-Object { '{0}: {1}' -f $_.Where, $_.Text })
        $bad | Should -BeNullOrEmpty
    }

    It 'uses SMS Provider and no other name for it' {
        @($script:Text | Where-Object { $_.Text -cmatch 'Provider Machine|SMS provider' }) | Should -BeNullOrEmpty
    }
}

Describe 'Tooltips' {
    It 'gives every field of a page a tooltip' {
        $missing = @($script:AllField | Where-Object { $_.Id -and $_.Kind -notin 'Group', 'Note' -and [string]::IsNullOrWhiteSpace([string]$_.Tip) } | ForEach-Object { $_.Id })
        $missing | Should -BeNullOrEmpty
    }

    It 'gives every button a tooltip' {
        $missing = foreach ($f in $script:AllField) { foreach ($b in @($f.Buttons | Where-Object { $_ })) { if ([string]::IsNullOrWhiteSpace([string]$b.Tip)) { '{0}/{1}' -f $f.Id, $b.Id } } }
        @($missing) | Should -BeNullOrEmpty
    }

    It 'writes each tooltip as one short sentence' {
        $bad = @($script:Text | Where-Object { $_.Kind -in 'Tip', 'ButtonTip', 'ItemTip' -and -not (Test-OneSentenceTip -Text $_.Text) } | ForEach-Object { '{0}: {1}' -f $_.Where, $_.Text })
        $bad | Should -BeNullOrEmpty
    }

    It 'gives every check box a label' {
        $missing = @($script:AllField | Where-Object { $_.Kind -eq 'Check' -and [string]::IsNullOrWhiteSpace([string]$_.Label) } | ForEach-Object { $_.Id })
        $missing | Should -BeNullOrEmpty
    }

    It 'gives every combo box its items' {
        $missing = @($script:AllField | Where-Object { $_.Kind -eq 'Combo' -and -not $_.Items -and -not $_.ItemsFrom } | ForEach-Object { $_.Id })
        $missing | Should -BeNullOrEmpty
    }

    It 'checks a number against a range' {
        $missing = @($script:AllField | Where-Object { $_.Kind -eq 'Number' -and ($null -eq $_.Min -or $null -eq $_.Max) } | ForEach-Object { $_.Id })
        $missing | Should -BeNullOrEmpty
    }
}

Describe 'Messages that name an Options page' {
    BeforeAll {
        $script:Chain = @{}
        foreach ($s in $script:Section) {
            $names = New-Object System.Collections.Generic.List[string]
            $cursor = $s
            while ($cursor) { $names.Insert(0, [string]$cursor.Title); $cursor = $(if ($cursor.Parent) { $script:ById[[string]$cursor.Parent] } else { $null }) }
            $script:Chain[[string]$s.Id] = @($names)
        }
        $script:Root = @($script:Section | Where-Object { -not $_.Parent } | ForEach-Object { [string]$_.Title })
        $script:Child = @{}
        foreach ($s in $script:Section) { if ($s.Parent) { $script:Child[([string]$script:ById[[string]$s.Parent].Title) + '|' + [string]$s.Title] = $true } }

        function Get-OptionsPathMention {
            param([string]$Text)
            foreach ($m in [regex]::Matches($Text, '(?<![A-Za-z])Options(?:, | > )(?<rest>[^.`''"\r\n)*]+)')) {
                [string]$m.Groups['rest'].Value
            }
        }
        function Test-OptionsPath {
            # The first segment is a top-level node, and the second, when it names a node, is a child of the first.
            param([string]$Rest)
            $segments = @($Rest -split ', | > ' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            if ($segments.Count -eq 0) { return $true }
            if (-not ($script:Root | Where-Object { ([string]$segments[0]).StartsWith($_) })) { return $false }
            $segments[0] = @($script:Root | Where-Object { ([string]$segments[0]).StartsWith($_) })[0]
            if ($segments.Count -gt 1) {
                $allTitles = @($script:Section | ForEach-Object { [string]$_.Title })
                if ($allTitles -ccontains $segments[1] -and -not $script:Child.ContainsKey($segments[0] + '|' + $segments[1])) { return $false }
            }
            return $true
        }
    }

    It 'recognizes a path of the tree (the check itself)' {
        Test-OptionsPath -Rest 'Connections, ConfigMgr' | Should -BeTrue
        Test-OptionsPath -Rest 'Updates and CVEs, One Click' | Should -BeTrue
        Test-OptionsPath -Rest 'Applications > Local installer sources' | Should -BeTrue
        Test-OptionsPath -Rest 'ConfigMgr Preferences' | Should -BeFalse
        Test-OptionsPath -Rest 'WSUS Publishing' | Should -BeFalse
        Test-OptionsPath -Rest 'Connections, Folders' | Should -BeFalse
    }

    It 'names a node of the tree in <Name>' -ForEach @(
        @{ Name = 'start-apppackager.ps1'; Path = 'start-apppackager.ps1' }
        @{ Name = 'AppPackagerCommon.psm1'; Path = 'Packagers\AppPackagerCommon.psm1' }
        @{ Name = 'AppPackagerDiscovery.psm1'; Path = 'Packagers\AppPackagerDiscovery.psm1' }
        @{ Name = 'AppPackagerWsus.psm1'; Path = 'Packagers\AppPackagerWsus.psm1' }
        @{ Name = 'Invoke-AppPackagerBuild.ps1'; Path = 'Invoke-AppPackagerBuild.ps1' }
    ) {
        $text = Get-Content -LiteralPath (Join-Path $script:RepoRoot $Path) -Raw
        $bad = @(Get-OptionsPathMention -Text $text | Where-Object { -not (Test-OptionsPath -Rest $_) })
        $bad | Should -BeNullOrEmpty
    }

    It 'names a node of the tree in every packager' {
        $bad = foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'Packagers') -Filter 'package-*.ps1' -File)) {
            foreach ($rest in @(Get-OptionsPathMention -Text (Get-Content -LiteralPath $file.FullName -Raw))) { if (-not (Test-OptionsPath -Rest $rest)) { '{0}: {1}' -f $file.Name, $rest } }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'names a node of the tree in the README and the user guide' {
        $files = @(Get-Item -LiteralPath (Join-Path $script:RepoRoot 'README.md')) + @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'docs') -Filter '*.md' -File -Recurse)
        $bad = foreach ($file in $files) {
            foreach ($rest in @(Get-OptionsPathMention -Text (Get-Content -LiteralPath $file.FullName -Raw))) { if (-not (Test-OptionsPath -Rest $rest)) { '{0}: {1}' -f $file.Name, $rest } }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'has no old panel name outside the changelog' {
        $pattern = 'ConfigMgr Preferences|Packager Preferences|One Click Settings|WSUS Publishing page|Options > Product Filter|Options, WSUS Publishing|Options > WSUS Publishing|Options > Script Signing'
        $files = @(Get-Item -LiteralPath $script:LauncherPath) + @(Get-Item -LiteralPath (Join-Path $script:RepoRoot 'README.md')) +
        @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'docs') -Filter '*.md' -File -Recurse) + @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'Packagers') -File | Where-Object { $_.Extension -in '.ps1', '.psm1' -and $_.Name -notlike '*.Tests.ps1' })
        $bad = foreach ($file in $files) {
            $n = 0
            foreach ($line in (Get-Content -LiteralPath $file.FullName)) { $n++; if ($line -cmatch $pattern) { '{0}:{1}: {2}' -f $file.Name, $n, $Matches[0] } }
        }
        @($bad) | Should -BeNullOrEmpty
    }
}

Describe 'The Setup wizard and the Options window use one set of fields' {
    It 'has no field of its own in the wizard' {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:LauncherPath, [ref]$null, [ref]$null)
        $wizard = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Show-FirstRunWizard' }, $false)
        $text = $wizard.Extent.Text
        $text | Should -Not -Match '<TextBox|<PasswordBox|<ComboBox'
        $text | Should -Match 'New-OptionsFieldGroup'
        $text | Should -Match 'Get-OptionsFieldTable'
    }

    It 'asks for <Id> in the wizard through its field definition' -ForEach @(
        @{ Id = 'cm.siteCode' }, @{ Id = 'cm.provider' }, @{ Id = 'cm.fileShare' }, @{ Id = 'pkg.downloadRoot' }, @{ Id = 'int.tenant' }, @{ Id = 'int.client' }, @{ Id = 'int.secret' },
        @{ Id = 'wsus.server' }, @{ Id = 'wsus.port' }, @{ Id = 'wsus.ssl' }
    ) {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:LauncherPath, [ref]$null, [ref]$null)
        $wizard = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Show-FirstRunWizard' }, $false)
        $wizard.Extent.Text | Should -Match ([regex]::Escape("'$Id'"))
        @($script:AllField | Where-Object { $_.Id -eq $Id }).Count | Should -Be 1
    }

    It 'shows the port of the WSUS server in one field only' {
        @($script:AllField | Where-Object { $_.Pref -eq 'Wsus.PortNumber' }).Count | Should -Be 1
        @($script:AllField | Where-Object { $_.Title -ceq 'Port' }).Count | Should -Be 1
    }
}
