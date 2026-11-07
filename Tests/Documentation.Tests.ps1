param ()

BeforeAll {
    $strTestsFolder = Split-Path -Path $PSCommandPath -Parent
    $script:strRepositoryRoot = Split-Path -Path $strTestsFolder -Parent
    $script:strDocsRoot = Join-Path -Path $script:strRepositoryRoot -ChildPath 'docs'
    $script:arrPage = @(Get-ChildItem -LiteralPath $script:strDocsRoot -Recurse -File -Filter '*.md')
    $script:arrImage = @(Get-ChildItem -LiteralPath (Join-Path -Path $script:strDocsRoot -ChildPath 'images') -Recurse -File -Filter '*.png')

    function Get-DocLink {
        param ([System.IO.FileInfo]$Page)
        $strText = [regex]::Replace((Get-Content -LiteralPath $Page.FullName -Raw), '(?s)```.*?```', '')
        foreach ($objMatch in [regex]::Matches($strText, '(?<!\!)\[[^\]]*\]\(([^)\s]+)(?:\s+"[^"]*")?\)')) {
            [pscustomobject]@{ Page = $Page; Target = $objMatch.Groups[1].Value; Kind = 'Link' }
        }
        foreach ($objMatch in [regex]::Matches($strText, '\!\[[^\]]*\]\(([^)\s]+)(?:\s+"[^"]*")?\)')) {
            [pscustomobject]@{ Page = $Page; Target = $objMatch.Groups[1].Value; Kind = 'Image' }
        }
    }
    $script:arrRef = @(foreach ($objPage in $script:arrPage) { Get-DocLink -Page $objPage })
    $script:arrReadmeRef = @(Get-DocLink -Page (Get-Item -LiteralPath (Join-Path -Path $script:strRepositoryRoot -ChildPath 'README.md')) | Where-Object { $_.Target -match '^docs/' -or $_.Target -match '^\.?/?docs/' })
}

Describe 'User guide' {
    It 'has an index and the standard pages' {
        foreach ($strName in 'index.md', 'getting-started.md', 'troubleshooting.md', 'external-references.md') {
            Test-Path -LiteralPath (Join-Path -Path $script:strDocsRoot -ChildPath $strName) | Should -BeTrue -Because $strName
        }
    }

    It 'resolves every relative link and image to a file' {
        $arrMissing = foreach ($objRef in @(@($script:arrRef) + @($script:arrReadmeRef) | Where-Object { $_.Target -notmatch '^(https?:|mailto:|#)' })) {
            $strPath = ($objRef.Target -split '#')[0]
            if (-not $strPath) { continue }
            $strFull = Join-Path -Path (Split-Path -Path $objRef.Page.FullName -Parent) -ChildPath $strPath
            if (-not (Test-Path -LiteralPath $strFull -PathType Leaf)) { '{0}: {1}' -f $objRef.Page.Name, $objRef.Target }
        }
        @($arrMissing) | Should -BeNullOrEmpty
    }

    It 'links every page from the index' {
        $strIndex = [regex]::Replace((Get-Content -LiteralPath (Join-Path -Path $script:strDocsRoot -ChildPath 'index.md') -Raw), '(?s)```.*?```', '')
        # Link targets only: an image reference and a link in an HTML comment do not count. A fragment and a link title are allowed.
        $strIndex = [regex]::Replace($strIndex, '(?s)<!--.*?-->', '')
        $arrTarget = @([regex]::Matches($strIndex, '(?<!\!)\[[^\]]*\]\(([^)\s#]+)(?:#[^)\s]*)?(?:\s+"[^"]*")?\)') | ForEach-Object { $_.Groups[1].Value })
        $arrOrphan = foreach ($objPage in $script:arrPage) {
            $strRelative = $objPage.FullName.Substring($script:strDocsRoot.Length + 1).Replace('\', '/')
            if ($strRelative -ne 'index.md' -and $arrTarget -notcontains $strRelative) { $strRelative }
        }
        @($arrOrphan) | Should -BeNullOrEmpty
    }

    It 'uses every image in at least one page' {
        $arrUsed = @($script:arrRef | Where-Object { $_.Kind -eq 'Image' } | ForEach-Object { $strFull = Join-Path -Path (Split-Path -Path $_.Page.FullName -Parent) -ChildPath $_.Target; Resolve-Path -LiteralPath $strFull -ErrorAction SilentlyContinue | ForEach-Object { $_.ProviderPath } })
        $arrUnused = @($script:arrImage | Where-Object { $arrUsed -notcontains $_.FullName } | ForEach-Object { $_.Name })
        $arrUnused | Should -BeNullOrEmpty
    }

    It 'names every image in lower case with hyphens' {
        @($script:arrImage | Where-Object { $_.Name -cnotmatch '^[a-z0-9]+(-[a-z0-9]+)*\.png$' }) | Should -BeNullOrEmpty
    }

    It 'gives every image a description' {
        @($script:arrRef | Where-Object { $_.Kind -eq 'Image' }).Count | Should -BeGreaterThan 0
        foreach ($objPage in $script:arrPage) {
            (Get-Content -LiteralPath $objPage.FullName -Raw) | Should -Not -Match '\!\[\s*\]\(' -Because $objPage.Name
        }
    }

    It 'links the guide from the README' {
        $strReadme = [regex]::Replace((Get-Content -LiteralPath (Join-Path -Path $script:strRepositoryRoot -ChildPath 'README.md') -Raw), '(?s)```.*?```', '')
        $strReadme | Should -Match '\]\(docs/index\.md\)'
    }
}
