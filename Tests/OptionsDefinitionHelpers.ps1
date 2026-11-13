# Helpers of the guard tests of the Options window. Dot-source it from a test.
# The definitions are data: reading them runs no WPF code and touches no file.

function Get-OptionsDefinitionFromLauncher {
    param([Parameter(Mandatory)][string]$Path)
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-OptionsSectionDefinition' }, $false)
    if (-not $fn) { throw 'start-apppackager.ps1 no longer defines Get-OptionsSectionDefinition.' }
    $script = [scriptblock]::Create($fn.Extent.Text + [Environment]::NewLine + 'Get-OptionsSectionDefinition')
    return @(& $script)
}

function Get-OptionsFieldFromLauncher {
    param([Parameter(Mandatory)][string]$Path)
    foreach ($section in @(Get-OptionsDefinitionFromLauncher -Path $Path)) {
        foreach ($f in @($section.Fields | Where-Object { $_ })) { $f }
    }
}

# Sentence case: the first word starts a capital letter, and a later word starts with a capital letter only when it is a
# name, an acronym or a name of the product. The names that the approved tree uses are in the list.
$script:ProperWord = @(
    'ConfigMgr', 'Intune', 'WSUS', 'SQL', 'SSL', 'NVD', 'API', 'DP', 'DPs', 'UNC', 'ID', 'URL', 'CVE', 'CVEs', 'GitHub', 'SMS', 'Provider', 'ADDLOCAL', 'ADDONS', 'LTSR', 'CEIP', 'DPAPI',
    'Adobe', 'Acrobat', 'Reader', 'Beyond', 'Compare', 'Citrix', 'Workspace', 'App', 'Apps', 'DBeaver', 'Community', 'Microsoft', 'Server', 'Management', 'Studio', 'TeamViewer', 'Host',
    'Protection', 'Package', 'Click', 'One', 'Content', 'Prep', 'Tool', 'Office', 'Deployment', 'Windows', 'Entra', 'Graph', 'Discover', 'Product', 'Teams', 'Zoom', 'WebEx', 'VDI', 'EPA', 'SSO',
    'Console', 'Test', 'Connection', 'HDX', 'MUI', 'AI', 'PFX', 'RDS', 'MSI', 'DevOps', 'Current', 'Release', 'Enterprise', 'Channel', 'Monthly', 'Online', 'Managed', 'Outlook', 'OneDrive', 'Skype',
    'Business', 'Excel', 'Word', 'PowerPoint', 'Publisher', 'OneNote', 'Access', 'Groove', 'Lync', 'Visual', 'Studio', 'Installer', 'Session', 'Recording', 'Local', 'App', 'StoreFront', 'Gateway',
    'FTP', 'SFTP', 'WebDAV', 'Quiet', 'Passive', 'System', 'User', 'Nested', 'Flat', 'None', 'Preferences', 'Advanced', 'Allow', 'Neighbor', 'Fallback'
)

function Test-SentenceCase {
    param([Parameter(Mandatory)][string]$Text)
    $clean = ($Text -replace '\.\.\.$', '').Trim()
    $words = @($clean -split '\s+' | Where-Object { $_ })
    if ($words.Count -eq 0) { return $true }
    if ($words[0] -cnotmatch '^([A-Z0-9]|[a-z]+[A-Z])') { return $false }
    for ($i = 1; $i -lt $words.Count; $i++) {
        $w = $words[$i].Trim(',', ':', '(', ')', '/', '-')
        if (-not $w) { continue }
        if ($w -cmatch '^[a-z0-9<\[]') { continue }
        if ($script:ProperWord -ccontains $w) { continue }
        # An acronym, a name with a digit, or a name that has a capital letter inside it.
        if ($w -cmatch '^[A-Z0-9_]+s?$' -or $w -cmatch '\d' -or $w -cmatch '^[A-Z][a-z]+[A-Z]') { continue }
        return $false
    }
    return $true
}

# One short sentence: a capital letter first, one full stop at the end, 100 characters or less.
function Test-OneSentenceTip {
    param([Parameter(Mandatory)][string]$Text)
    if ($Text.Length -gt 100) { return $false }
    if ($Text -cnotmatch '^[A-Z]') { return $false }
    if ($Text -notmatch '\.$') { return $false }
    return ($Text.TrimEnd('.') -notmatch '\.\s')
}
