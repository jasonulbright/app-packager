#Requires -Version 5.1
<#
    Converts an existing ConfigMgr application into Intune Win32 apps.

    The reader works on the application document that ConfigMgr stores
    (SDMPackageXML), so the parse and the translation run on plain text and
    objects: no site, no WPF, no Graph. Only the functions at the end of this
    file touch the site (read-only), the file system, or Graph, and they call
    the existing App Packager publisher (Publish-IntuneWin32App).

    One ConfigMgr application can hold several deployment types. ConfigMgr
    installs the first deployment type, by priority, whose requirements pass.
    Intune has no deployment types and no priority, so each deployment type
    becomes its own Win32 app, and the requirement rules of every app keep
    the "first match wins" meaning (see Get-CMImportVariantExclusion).
#>

Set-StrictMode -Version 2.0

# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

function Write-ImpLog {
    param([string]$Message, [string]$Level = 'INFO')
    if (Get-Command -Name Write-Log -ErrorAction SilentlyContinue) {
        Write-Log $Message -Level $Level
    }
    else {
        Write-Verbose $Message
    }
}

function ConvertTo-ImpSingleQuoted {
    # A value as the inside of a PowerShell single-quoted string.
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { return '' }
    return $Value.Replace("'", "''")
}

function Get-ImpChildren {
    param([Parameter(Mandatory)][AllowNull()]$Node, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $Node) { return @() }
    return @($Node.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element -and $_.LocalName -eq $Name })
}

function Get-ImpChild {
    param([Parameter(Mandatory)][AllowNull()]$Node, [Parameter(Mandatory)][string]$Name)
    $found = @(Get-ImpChildren -Node $Node -Name $Name)
    if ($found.Count -eq 0) { return $null }
    return $found[0]
}

function Get-ImpText {
    param($Node)
    if ($null -eq $Node) { return '' }
    return [string]$Node.InnerText
}

function Get-ImpAttribute {
    param($Node, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $Node -or -not $Node.HasAttribute($Name)) { return '' }
    return [string]$Node.GetAttribute($Name)
}

function Get-ImpArgMap {
    # The Arg elements of an install or uninstall action as name to value.
    # Array arguments (Int32[]) become object arrays.
    param($ActionNode)
    $map = @{}
    if ($null -eq $ActionNode) { return $map }
    $argsNode = Get-ImpChild -Node $ActionNode -Name 'Args'
    foreach ($arg in (Get-ImpChildren -Node $argsNode -Name 'Arg')) {
        $name = Get-ImpAttribute -Node $arg -Name 'Name'
        $type = Get-ImpAttribute -Node $arg -Name 'Type'
        if ($type -like '*[[]]') {
            $map[$name] = @(Get-ImpChildren -Node $arg -Name 'Item' | ForEach-Object { Get-ImpText $_ })
        }
        else {
            $map[$name] = Get-ImpText $arg
        }
    }
    return $map
}

function Get-ImpMapValue {
    param([hashtable]$Map, [string]$Name, $Default = '')
    if ($Map.ContainsKey($Name)) { return $Map[$Name] }
    return $Default
}

function ConvertTo-ImpSafeName {
    param([AllowEmptyString()][string]$Name)
    $safe = ($Name -replace '[^\w.\- ]', '_').Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'item' }
    return $safe
}

# ---------------------------------------------------------------------------
# Detection: the stored expression tree
# ---------------------------------------------------------------------------

$script:ImpOperatorMap = @{
    'Equals'                  = 'IsEquals'
    'NotEquals'               = 'NotEquals'
    'GreaterThan'             = 'GreaterThan'
    'GreaterEquals'           = 'GreaterEquals'
    'GreaterThanOrEquals'     = 'GreaterEquals'
    'LessThan'                = 'LessThan'
    'LessEquals'              = 'LessEquals'
    'LessThanOrEquals'        = 'LessEquals'
    'BeginsWith'              = 'BeginsWith'
    'EndsWith'                = 'EndsWith'
    'Contains'                = 'Contains'
    'NotBeginsWith'           = 'NotBeginsWith'
    'NotEndsWith'             = 'NotEndsWith'
    'NotContains'             = 'NotContains'
}

function ConvertFrom-CMDetectionSettings {
    # The Settings element of an enhanced detection method as
    # logical name to setting.
    param($SettingsNode)
    $settings = @{}
    if ($null -eq $SettingsNode) { return $settings }
    foreach ($node in $SettingsNode.ChildNodes) {
        if ($node.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
        $logical = Get-ImpAttribute -Node $node -Name 'LogicalName'
        if ([string]::IsNullOrEmpty($logical)) { continue }
        $setting = [ordered]@{ Source = $node.LocalName; Is64Bit = $false; Hive = ''; Key = ''; Path = ''; Filter = ''; ValueName = ''; DataType = (Get-ImpAttribute -Node $node -Name 'DataType') }
        switch ($node.LocalName) {
            'File' {
                $setting.Is64Bit = ((Get-ImpAttribute -Node $node -Name 'Is64Bit') -eq 'true')
                $setting.Path = Get-ImpText (Get-ImpChild -Node $node -Name 'Path')
                $setting.Filter = Get-ImpText (Get-ImpChild -Node $node -Name 'Filter')
            }
            'Folder' {
                $setting.Is64Bit = ((Get-ImpAttribute -Node $node -Name 'Is64Bit') -eq 'true')
                $setting.Path = Get-ImpText (Get-ImpChild -Node $node -Name 'Path')
                $setting.Filter = Get-ImpText (Get-ImpChild -Node $node -Name 'Filter')
            }
            'RegistryKey' {
                $setting.Is64Bit = ((Get-ImpAttribute -Node $node -Name 'Is64Bit') -eq 'true')
                $setting.Hive = Get-ImpAttribute -Node $node -Name 'Hive'
                $setting.Key = Get-ImpText (Get-ImpChild -Node $node -Name 'Key')
            }
            'SimpleSetting' {
                $source = Get-ImpChild -Node $node -Name 'RegistryDiscoverySource'
                if ($source) {
                    $setting.Source = 'Registry'
                    $setting.Is64Bit = ((Get-ImpAttribute -Node $source -Name 'Is64Bit') -eq 'true')
                    $setting.Hive = Get-ImpAttribute -Node $source -Name 'Hive'
                    $setting.Key = Get-ImpText (Get-ImpChild -Node $source -Name 'Key')
                    $setting.ValueName = Get-ImpText (Get-ImpChild -Node $source -Name 'ValueName')
                }
                else {
                    $setting.Source = 'Other'
                }
            }
            default { $setting.Source = 'Other' }
        }
        $settings[$logical] = [pscustomobject]$setting
    }
    return $settings
}

function ConvertFrom-CMDetectionExpression {
    # One Expression element to a tree. A group is And or Or over children.
    # A leaf compares one stored setting with a constant.
    param([Parameter(Mandatory)]$Expression, [Parameter(Mandatory)][hashtable]$Settings)

    $operator = Get-ImpText (Get-ImpChild -Node $Expression -Name 'Operator')
    $operands = Get-ImpChild -Node $Expression -Name 'Operands'

    if ($operator -eq 'And' -or $operator -eq 'Or') {
        $children = @(Get-ImpChildren -Node $operands -Name 'Expression' | ForEach-Object { ConvertFrom-CMDetectionExpression -Expression $_ -Settings $Settings })
        return [pscustomobject]@{ Kind = 'Group'; Operator = $operator; Children = $children }
    }

    $reference = Get-ImpChild -Node $operands -Name 'SettingReference'
    $constant = Get-ImpChild -Node $operands -Name 'ConstantValue'
    $logical = Get-ImpAttribute -Node $reference -Name 'SettingLogicalName'
    $sourceType = Get-ImpAttribute -Node $reference -Name 'SettingSourceType'
    $method = Get-ImpAttribute -Node $reference -Name 'Method'
    $property = Get-ImpAttribute -Node $reference -Name 'PropertyPath'
    $dataType = Get-ImpAttribute -Node $reference -Name 'DataType'
    $expected = Get-ImpAttribute -Node $constant -Name 'Value'

    $leaf = [ordered]@{
        Kind = 'Leaf'; Source = $sourceType; Property = 'Value'; Operator = ''; Expected = $expected; DataType = $dataType
        Path = ''; Name = ''; Hive = ''; Key = ''; Is64Bit = $false; Supported = $true; Reason = ''
    }
    $mapped = ''
    if ($script:ImpOperatorMap.ContainsKey($operator)) { $mapped = $script:ImpOperatorMap[$operator] }
    $leaf.Operator = $mapped

    $setting = $null
    if ($logical -and $Settings.ContainsKey($logical)) { $setting = $Settings[$logical] }
    if (-not $setting -or $setting.Source -eq 'Other' -or -not $mapped) {
        $leaf.Supported = $false
        $leaf.Reason = "Detection setting source '$sourceType' with operator '$operator' is not converted."
        return [pscustomobject]$leaf
    }

    $leaf.Is64Bit = [bool]$setting.Is64Bit
    switch ($sourceType) {
        { $_ -eq 'File' -or $_ -eq 'Folder' } {
            $leaf.Path = $setting.Path
            $leaf.Name = $setting.Filter
            if ($method -eq 'Count') {
                # "Count not equals 0" is existence; "count equals 0" is absence.
                $leaf.Property = 'Exists'
                $exists = ($mapped -eq 'NotEquals' -and $expected -eq '0') -or ($mapped -eq 'GreaterThan' -and $expected -eq '0')
                $absent = ($mapped -eq 'IsEquals' -and $expected -eq '0')
                if (-not ($exists -or $absent)) {
                    $leaf.Supported = $false
                    $leaf.Reason = "A $sourceType count comparison other than 'exists' or 'does not exist' is not converted."
                }
                $leaf.Expected = [string]([bool]$exists)
                $leaf.Operator = 'IsEquals'
            }
            else {
                $leaf.Property = $property
            }
        }
        'RegistryKey' {
            $leaf.Hive = $setting.Hive
            $leaf.Key = $setting.Key
            $leaf.Property = 'Exists'
            $want = ($expected -eq 'true')
            if ($mapped -eq 'NotEquals') { $want = -not $want }
            $leaf.Expected = [string]$want
            $leaf.Operator = 'IsEquals'
        }
        'Registry' {
            $leaf.Hive = $setting.Hive
            $leaf.Key = $setting.Key
            $leaf.Name = $setting.ValueName
            $leaf.Property = 'Value'
        }
        default {
            $leaf.Supported = $false
            $leaf.Reason = "Detection setting source '$sourceType' is not converted."
        }
    }
    return [pscustomobject]$leaf
}

function ConvertFrom-CMDeploymentTypeDetection {
    # The detection of one deployment type: a script, an MSI product code,
    # or an enhanced expression tree.
    param([Parameter(Mandatory)]$CustomData, $Installer)

    $method = Get-ImpText (Get-ImpChild -Node $CustomData -Name 'DetectionMethod')
    switch ($method) {
        'Script' {
            $scriptNode = Get-ImpChild -Node $CustomData -Name 'DetectionScript'
            return [pscustomobject]@{
                Method = 'Script'; Tree = $null
                ScriptText = (Get-ImpText $scriptNode); ScriptLanguage = (Get-ImpAttribute -Node $scriptNode -Name 'Language')
                ProductCode = ''; ProductVersion = ''
            }
        }
        'ProductCode' {
            return [pscustomobject]@{
                Method = 'ProductCode'; Tree = $null; ScriptText = ''; ScriptLanguage = ''
                ProductCode = (Get-ImpText (Get-ImpChild -Node $CustomData -Name 'ProductCode'))
                ProductVersion = (Get-ImpText (Get-ImpChild -Node $CustomData -Name 'ProductVersion'))
            }
        }
        'Enhanced' {
            $enhanced = Get-ImpChild -Node $CustomData -Name 'EnhancedDetectionMethod'
            $settings = ConvertFrom-CMDetectionSettings -SettingsNode (Get-ImpChild -Node $enhanced -Name 'Settings')
            $rule = Get-ImpChild -Node $enhanced -Name 'Rule'
            $expression = Get-ImpChild -Node $rule -Name 'Expression'
            $tree = $null
            if ($expression) { $tree = ConvertFrom-CMDetectionExpression -Expression $expression -Settings $settings }
            return [pscustomobject]@{ Method = 'Enhanced'; Tree = $tree; ScriptText = ''; ScriptLanguage = ''; ProductCode = ''; ProductVersion = '' }
        }
        default {
            return [pscustomobject]@{ Method = $(if ($method) { $method } else { 'None' }); Tree = $null; ScriptText = ''; ScriptLanguage = ''; ProductCode = ''; ProductVersion = '' }
        }
    }
}

# ---------------------------------------------------------------------------
# Requirements, dependencies, supersedence
# ---------------------------------------------------------------------------

function ConvertFrom-CMRequirementRule {
    param([Parameter(Mandatory)]$Rule)

    $annotation = Get-ImpChild -Node $Rule -Name 'Annotation'
    $display = Get-ImpAttribute -Node (Get-ImpChild -Node $annotation -Name 'DisplayName') -Name 'Text'
    $result = [ordered]@{ Kind = 'Unsupported'; DisplayName = $display; Operator = ''; PlatformIds = @(); LogicalName = ''; Values = @(); DataType = ''; Reason = '' }

    $os = Get-ImpChild -Node $Rule -Name 'OperatingSystemExpression'
    if ($os) {
        $result.Kind = 'OsPlatform'
        $result.Operator = Get-ImpText (Get-ImpChild -Node $os -Name 'Operator')
        $result.PlatformIds = @(Get-ImpChildren -Node (Get-ImpChild -Node $os -Name 'Operands') -Name 'RuleExpression' | ForEach-Object { Get-ImpAttribute -Node $_ -Name 'RuleId' })
        return [pscustomobject]$result
    }

    $expression = Get-ImpChild -Node $Rule -Name 'Expression'
    if (-not $expression) {
        $result.Reason = 'The requirement has no expression.'
        return [pscustomobject]$result
    }
    $operator = Get-ImpText (Get-ImpChild -Node $expression -Name 'Operator')
    $operands = Get-ImpChild -Node $expression -Name 'Operands'
    $global = Get-ImpChild -Node $operands -Name 'GlobalSettingReference'
    if (-not $global -or $operator -eq 'And' -or $operator -eq 'Or') {
        $result.Reason = 'The requirement uses a custom setting or a compound expression.'
        return [pscustomobject]$result
    }
    $values = @()
    $constant = Get-ImpChild -Node $operands -Name 'ConstantValue'
    if ($constant) { $values = @(Get-ImpAttribute -Node $constant -Name 'Value') }
    $list = Get-ImpChild -Node $operands -Name 'ConstantValueList'
    if ($list) { $values = @(Get-ImpChildren -Node $list -Name 'ConstantValue' | ForEach-Object { Get-ImpAttribute -Node $_ -Name 'Value' }) }

    $result.Kind = 'Global'
    $result.Operator = $operator
    $result.LogicalName = Get-ImpAttribute -Node $global -Name 'LogicalName'
    $result.DataType = Get-ImpAttribute -Node $global -Name 'DataType'
    $result.Values = $values
    if ((Get-ImpAttribute -Node $global -Name 'AuthoringScopeId') -ne 'GLOBAL') {
        $result.Kind = 'Unsupported'
        $result.Reason = 'The requirement uses a custom global condition.'
    }
    return [pscustomobject]$result
}

function ConvertFrom-CMDeploymentTypeRule {
    # Dependencies and supersedence share one rule shape: an expression over
    # references to other deployment types.
    param([Parameter(Mandatory)][AllowNull()]$Container)

    $groups = @()
    foreach ($rule in (Get-ImpChildren -Node $Container -Name 'DeploymentTypeRule')) {
        $annotation = Get-ImpChild -Node $rule -Name 'Annotation'
        $groupName = Get-ImpAttribute -Node (Get-ImpChild -Node $annotation -Name 'DisplayName') -Name 'Text'
        $expression = Get-ImpChild -Node $rule -Name 'DeploymentTypeExpression'
        $operator = ''
        $intents = @()
        if ($expression) {
            $operator = Get-ImpText (Get-ImpChild -Node $expression -Name 'Operator')
            $intents = @(Get-ImpChildren -Node (Get-ImpChild -Node $expression -Name 'Operands') -Name 'DeploymentTypeIntentExpression')
        }
        else {
            $intents = @(Get-ImpChildren -Node $rule -Name 'DeploymentTypeIntentExpression')
        }
        $targets = @()
        foreach ($intent in $intents) {
            $application = Get-ImpChild -Node $intent -Name 'DeploymentTypeApplicationReference'
            $deploymentType = Get-ImpChild -Node $intent -Name 'DeploymentTypeReference'
            $targets += [pscustomobject]@{
                DesiredState       = Get-ImpAttribute -Node $intent -Name 'DesiredState'
                ApplicationLogical = Get-ImpAttribute -Node $application -Name 'LogicalName'
                DeploymentTypeLogical = Get-ImpAttribute -Node $deploymentType -Name 'LogicalName'
                ApplicationName    = ''
                DeploymentTypeName = ''
            }
        }
        $groups += [pscustomobject]@{ GroupName = $groupName; Operator = $operator; Targets = $targets }
    }
    return $groups
}

# ---------------------------------------------------------------------------
# The application document
# ---------------------------------------------------------------------------

function ConvertFrom-CMApplicationDigest {
    <#
    .SYNOPSIS
        Parses the stored document of a ConfigMgr application into a model.

    .DESCRIPTION
        The document (SDMPackageXML of the application) is self-contained: it
        holds the application and every deployment type. Priority is the order
        of the deployment types, which matches PriorityInLatestApp. Read the
        application, not a deployment type: Get-CMDeploymentType returns the
        first deployment type's document for every type of an application.
    #>
    param([Parameter(Mandatory)][string]$Xml)

    $doc = New-Object System.Xml.XmlDocument
    $doc.XmlResolver = $null
    $doc.LoadXml($Xml.TrimStart([char]0xFEFF))
    $root = $doc.DocumentElement
    if ($null -eq $root -or $root.LocalName -ne 'AppMgmtDigest') {
        throw 'The document is not a ConfigMgr application digest (AppMgmtDigest).'
    }

    $app = Get-ImpChild -Node $root -Name 'Application'
    $info = Get-ImpChild -Node (Get-ImpChild -Node $app -Name 'DisplayInfo') -Name 'Info'
    $title = Get-ImpText (Get-ImpChild -Node $app -Name 'Title')
    if (-not $title) { $title = Get-ImpText (Get-ImpChild -Node $info -Name 'Title') }
    $publisher = Get-ImpText (Get-ImpChild -Node $app -Name 'Publisher')
    if (-not $publisher) { $publisher = Get-ImpText (Get-ImpChild -Node $info -Name 'Publisher') }
    $version = Get-ImpText (Get-ImpChild -Node $app -Name 'SoftwareVersion')
    if (-not $version) { $version = Get-ImpText (Get-ImpChild -Node $info -Name 'Version') }

    $iconData = ''
    $iconNode = Get-ImpChild -Node $info -Name 'Icon'
    if ($iconNode) { $iconData = Get-ImpText (Get-ImpChild -Node $iconNode -Name 'Data') }

    $deploymentTypes = @()
    $priority = 0
    foreach ($dtNode in (Get-ImpChildren -Node $root -Name 'DeploymentType')) {
        $priority++
        $installer = Get-ImpChild -Node $dtNode -Name 'Installer'
        $technology = Get-ImpText (Get-ImpChild -Node $dtNode -Name 'Technology')
        $dt = [ordered]@{
            Name = (Get-ImpText (Get-ImpChild -Node $dtNode -Name 'Title'))
            LogicalName = (Get-ImpAttribute -Node $dtNode -Name 'LogicalName')
            Priority = $priority
            Technology = $technology
            UnsupportedReason = ''
            ExecutionContext = 'System'
            RequiresLogOn = ''
            RequiresUserInteraction = $false
            PostInstallBehavior = 'BasedOnExitCode'
            MaxRuntimeMins = 0
            EstimatedRuntimeMins = 0
            RunAs32Bit = $false
            InstallCommand = ''
            UninstallCommand = ''
            UninstallSetting = ''
            ContentLocation = ''
            ContentId = ''
            UninstallContentId = ''
            ContentFiles = @()
            ExitCodes = @()
            Detection = $null
            Requirements = @()
            Dependencies = @()
            Supersedes = @()
        }

        if ($technology -ne 'MSI' -and $technology -ne 'Script') {
            $dt.UnsupportedReason = "The deployment technology '$technology' is not converted. Intune Win32 apps carry Windows Installer and script installers."
        }
        if ($installer) {
            $custom = Get-ImpChild -Node $installer -Name 'CustomData'
            $installAction = Get-ImpChild -Node $installer -Name 'InstallAction'
            $actionArgs = Get-ImpArgMap -ActionNode $installAction
            $dt.ExecutionContext = [string](Get-ImpMapValue -Map $actionArgs -Name 'ExecutionContext' -Default 'System')
            $dt.RequiresLogOn = [string](Get-ImpMapValue -Map $actionArgs -Name 'RequiresLogOn' -Default '')
            $dt.RequiresUserInteraction = ([string](Get-ImpMapValue -Map $actionArgs -Name 'RequiresUserInteraction' -Default 'false') -eq 'true')
            $dt.PostInstallBehavior = [string](Get-ImpMapValue -Map $actionArgs -Name 'PostInstallBehavior' -Default 'BasedOnExitCode')
            $dt.MaxRuntimeMins = [int](Get-ImpMapValue -Map $actionArgs -Name 'MaxExecuteTime' -Default 0)
            $dt.EstimatedRuntimeMins = [int](Get-ImpMapValue -Map $actionArgs -Name 'ExecuteTime' -Default 0)
            $dt.RunAs32Bit = ([string](Get-ImpMapValue -Map $actionArgs -Name 'RunAs32Bit' -Default 'false') -eq 'true')

            if ($custom) {
                $dt.InstallCommand = Get-ImpText (Get-ImpChild -Node $custom -Name 'InstallCommandLine')
                $dt.UninstallCommand = Get-ImpText (Get-ImpChild -Node $custom -Name 'UninstallCommandLine')
                $dt.UninstallSetting = Get-ImpText (Get-ImpChild -Node $custom -Name 'UninstallSetting')
                $dt.ContentId = Get-ImpAttribute -Node (Get-ImpChild -Node $custom -Name 'InstallContent') -Name 'ContentId'
                $dt.UninstallContentId = Get-ImpAttribute -Node (Get-ImpChild -Node $custom -Name 'UninstallContent') -Name 'ContentId'
                $dt.ExitCodes = @(Get-ImpChildren -Node (Get-ImpChild -Node $custom -Name 'ExitCodes') -Name 'ExitCode' | ForEach-Object {
                    [pscustomobject]@{ Code = [int](Get-ImpAttribute -Node $_ -Name 'Code'); Class = (Get-ImpAttribute -Node $_ -Name 'Class') }
                })
                $dt.Detection = ConvertFrom-CMDeploymentTypeDetection -CustomData $custom -Installer $installer
            }

            foreach ($content in (Get-ImpChildren -Node (Get-ImpChild -Node $installer -Name 'Contents') -Name 'Content')) {
                if ($dt.ContentId -and (Get-ImpAttribute -Node $content -Name 'ContentId') -ne $dt.ContentId) { continue }
                $dt.ContentLocation = (Get-ImpText (Get-ImpChild -Node $content -Name 'Location')).TrimEnd('\')
                $dt.ContentFiles = @(Get-ImpChildren -Node $content -Name 'File' | ForEach-Object { Get-ImpAttribute -Node $_ -Name 'Name' })
                break
            }
        }

        $requirements = Get-ImpChild -Node $dtNode -Name 'Requirements'
        $dt.Requirements = @(Get-ImpChildren -Node $requirements -Name 'Rule' | ForEach-Object { ConvertFrom-CMRequirementRule -Rule $_ })
        $dt.Dependencies = @(ConvertFrom-CMDeploymentTypeRule -Container (Get-ImpChild -Node $dtNode -Name 'Dependencies'))
        $dt.Supersedes = @(ConvertFrom-CMDeploymentTypeRule -Container (Get-ImpChild -Node $dtNode -Name 'Supersedes'))
        $deploymentTypes += [pscustomobject]$dt
    }

    return [pscustomobject]@{
        Name            = $title
        Publisher       = $publisher
        Version         = $version
        Description     = (Get-ImpText (Get-ImpChild -Node $info -Name 'Description'))
        LogicalName     = (Get-ImpAttribute -Node $app -Name 'LogicalName')
        IconBase64      = $iconData
        DeploymentTypes = $deploymentTypes
    }
}

# ---------------------------------------------------------------------------
# Detection to Intune rules
# ---------------------------------------------------------------------------

function Get-ImpLeaves {
    param($Tree)
    if ($null -eq $Tree) { return @() }
    if ($Tree.Kind -eq 'Leaf') { return @($Tree) }
    return @($Tree.Children | ForEach-Object { Get-ImpLeaves -Tree $_ })
}

function Test-ImpTreeIsFlatAnd {
    # True when the tree is a single leaf or a chain of And groups.
    param($Tree)
    if ($Tree.Kind -eq 'Leaf') { return $true }
    if ($Tree.Operator -ne 'And') { return $false }
    foreach ($child in $Tree.Children) { if (-not (Test-ImpTreeIsFlatAnd -Tree $child)) { return $false } }
    return $true
}

function ConvertTo-ImpHiveRoot {
    param([string]$Hive)
    if ($Hive -match 'CURRENT_USER|CurrentUser') { return 'HKEY_CURRENT_USER\' }
    return 'HKEY_LOCAL_MACHINE\'
}

function ConvertTo-ImpNativeRule {
    # One leaf to a Graph rule hashtable, or $null when no native rule fits.
    param([Parameter(Mandatory)]$Leaf)

    $operatorMap = @{
        IsEquals = 'equal'; NotEquals = 'notEqual'; GreaterThan = 'greaterThan'; GreaterEquals = 'greaterThanOrEqual'
        LessThan = 'lessThan'; LessEquals = 'lessThanOrEqual'
    }
    if (-not $Leaf.Supported) { return $null }

    if ($Leaf.Source -eq 'File' -or $Leaf.Source -eq 'Folder') {
        if ($Leaf.Name -match '[*?]') { return $null }
        if ($Leaf.Property -eq 'Exists' -and $Leaf.Expected -eq 'True') {
            return @{
                '@odata.type' = '#microsoft.graph.win32LobAppFileSystemRule'; ruleType = 'detection'
                path = $Leaf.Path; fileOrFolderName = $Leaf.Name; check32BitOn64System = (-not [bool]$Leaf.Is64Bit)
                operationType = 'exists'; operator = 'notConfigured'; comparisonValue = $null
            }
        }
        if ($Leaf.Source -eq 'File' -and $Leaf.Property -eq 'Version' -and $operatorMap.ContainsKey($Leaf.Operator)) {
            return @{
                '@odata.type' = '#microsoft.graph.win32LobAppFileSystemRule'; ruleType = 'detection'
                path = $Leaf.Path; fileOrFolderName = $Leaf.Name; check32BitOn64System = (-not [bool]$Leaf.Is64Bit)
                operationType = 'version'; operator = $operatorMap[$Leaf.Operator]; comparisonValue = $Leaf.Expected
            }
        }
        return $null
    }
    if ($Leaf.Source -eq 'RegistryKey' -and $Leaf.Property -eq 'Exists' -and $Leaf.Expected -eq 'True') {
        return @{
            '@odata.type' = '#microsoft.graph.win32LobAppRegistryRule'; ruleType = 'detection'
            check32BitOn64System = (-not [bool]$Leaf.Is64Bit); keyPath = ((ConvertTo-ImpHiveRoot $Leaf.Hive) + $Leaf.Key)
            valueName = $null; operationType = 'exists'; operator = 'notConfigured'; comparisonValue = $null
        }
    }
    if ($Leaf.Source -eq 'Registry' -and $Leaf.Property -eq 'Value' -and $operatorMap.ContainsKey($Leaf.Operator)) {
        $type = switch ($Leaf.DataType) { 'Int64' { 'integer' } 'Version' { 'version' } 'String' { 'string' } default { '' } }
        if (-not $type) { return $null }
        return @{
            '@odata.type' = '#microsoft.graph.win32LobAppRegistryRule'; ruleType = 'detection'
            check32BitOn64System = (-not [bool]$Leaf.Is64Bit); keyPath = ((ConvertTo-ImpHiveRoot $Leaf.Hive) + $Leaf.Key)
            valueName = $Leaf.Name; operationType = $type; operator = $operatorMap[$Leaf.Operator]; comparisonValue = $Leaf.Expected
        }
    }
    return $null
}

$script:ImpDetectionHelpers = @'
function Expand-ImpPath([string]$Path, [bool]$Is64Bit) {
    $p = $Path
    if (-not $Is64Bit -and [Environment]::Is64BitOperatingSystem) {
        $p = $p -replace '(?i)%ProgramFiles%', '%ProgramFiles(x86)%'
        $p = $p -replace '(?i)%CommonProgramFiles%', '%CommonProgramFiles(x86)%'
    }
    $p = [Environment]::ExpandEnvironmentVariables($p)
    if (-not $Is64Bit -and [Environment]::Is64BitOperatingSystem) {
        $system = Join-Path $env:SystemRoot 'System32'
        if ($p.StartsWith($system, [StringComparison]::OrdinalIgnoreCase)) { $p = (Join-Path $env:SystemRoot 'SysWOW64') + $p.Substring($system.Length) }
    }
    return $p
}
function Get-ImpRegistryValue([string]$Hive, [bool]$Is64Bit, [string]$Key, [string]$Name) {
    $hiveName = if ($Hive -match 'CURRENT_USER|CurrentUser') { 'CurrentUser' } else { 'LocalMachine' }
    $view = if ($Is64Bit -or -not [Environment]::Is64BitOperatingSystem) { [Microsoft.Win32.RegistryView]::Registry64 } else { [Microsoft.Win32.RegistryView]::Registry32 }
    if (-not [Environment]::Is64BitOperatingSystem) { $view = [Microsoft.Win32.RegistryView]::Default }
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]$hiveName, $view)
    try {
        $subKey = $base.OpenSubKey($Key)
        if ($null -eq $subKey) { return [pscustomobject]@{ KeyExists = $false; Value = $null } }
        try { return [pscustomobject]@{ KeyExists = $true; Value = $(if ($Name) { $subKey.GetValue($Name) } else { $null }) } }
        finally { $subKey.Close() }
    }
    finally { $base.Close() }
}
function ConvertTo-ImpVersion($Value) {
    $text = ([string]$Value) -replace ',\s*', '.' -replace '[^\d.].*$', ''
    $parsed = $null
    if ([version]::TryParse($text, [ref]$parsed)) { return $parsed }
    return $null
}
function Compare-ImpValue($Actual, $Expected, [string]$Type, [string]$Operator) {
    if ($null -eq $Actual) { return $false }
    switch ($Type) {
        'Version' { $a = ConvertTo-ImpVersion $Actual; $e = ConvertTo-ImpVersion $Expected; if ($null -eq $a -or $null -eq $e) { return $false }; $c = $a.CompareTo($e) }
        'Int64'   { $c = ([int64]$Actual).CompareTo([int64]$Expected) }
        'DateTime' { $c = ([datetime]$Actual).CompareTo([datetime]$Expected) }
        default   {
            $a = [string]$Actual; $e = [string]$Expected
            switch ($Operator) {
                'BeginsWith' { return $a.StartsWith($e, [StringComparison]::OrdinalIgnoreCase) }
                'EndsWith' { return $a.EndsWith($e, [StringComparison]::OrdinalIgnoreCase) }
                'Contains' { return ($a.IndexOf($e, [StringComparison]::OrdinalIgnoreCase) -ge 0) }
                'NotBeginsWith' { return -not $a.StartsWith($e, [StringComparison]::OrdinalIgnoreCase) }
                'NotEndsWith' { return -not $a.EndsWith($e, [StringComparison]::OrdinalIgnoreCase) }
                'NotContains' { return ($a.IndexOf($e, [StringComparison]::OrdinalIgnoreCase) -lt 0) }
            }
            $c = [string]::Compare($a, $e, $true)
        }
    }
    switch ($Operator) {
        'IsEquals' { return ($c -eq 0) } 'NotEquals' { return ($c -ne 0) } 'GreaterThan' { return ($c -gt 0) }
        'GreaterEquals' { return ($c -ge 0) } 'LessThan' { return ($c -lt 0) } 'LessEquals' { return ($c -le 0) }
    }
    return $false
}
function Test-ImpLeaf([string]$Source, [string]$Property, [string]$Operator, [string]$Expected, [string]$DataType, [string]$Path, [string]$Name, [string]$Hive, [string]$Key, [bool]$Is64Bit) {
    switch ($Source) {
        { $_ -eq 'File' -or $_ -eq 'Folder' } {
            $directory = Expand-ImpPath $Path $Is64Bit
            $item = $null
            if (Test-Path -LiteralPath $directory -PathType Container) {
                $kind = if ($Source -eq 'File') { @{ File = $true } } else { @{ Directory = $true } }
                $item = Get-ChildItem -LiteralPath $directory -Filter $Name -Force -ErrorAction SilentlyContinue @kind | Select-Object -First 1
            }
            if ($Property -eq 'Exists') { return (($null -ne $item) -eq ($Expected -eq 'True')) }
            if ($null -eq $item) { return $false }
            switch ($Property) {
                'Version' { return (Compare-ImpValue $item.VersionInfo.FileVersion $Expected 'Version' $Operator) }
                'DateModified' { return (Compare-ImpValue $item.LastWriteTime $Expected 'DateTime' $Operator) }
                'DateCreated' { return (Compare-ImpValue $item.CreationTime $Expected 'DateTime' $Operator) }
                'Size' { return (Compare-ImpValue $item.Length $Expected 'Int64' $Operator) }
            }
            return $false
        }
        'RegistryKey' {
            $found = Get-ImpRegistryValue $Hive $Is64Bit $Key ''
            return ($found.KeyExists -eq ($Expected -eq 'True'))
        }
        'Registry' {
            $found = Get-ImpRegistryValue $Hive $Is64Bit $Key $Name
            return (Compare-ImpValue $found.Value $Expected $DataType $Operator)
        }
    }
    return $false
}
'@

function ConvertTo-ImpScriptExpression {
    # The detection tree as one PowerShell boolean expression.
    param([Parameter(Mandatory)]$Tree)

    if ($Tree.Kind -eq 'Group') {
        $joiner = $(if ($Tree.Operator -eq 'Or') { ' -or ' } else { ' -and ' })
        $parts = @($Tree.Children | ForEach-Object { ConvertTo-ImpScriptExpression -Tree $_ })
        return '(' + ($parts -join $joiner) + ')'
    }
    if (-not $Tree.Supported) { throw $Tree.Reason }
    $q = { param($v) "'" + (ConvertTo-ImpSingleQuoted ([string]$v)) + "'" }
    return ('(Test-ImpLeaf -Source {0} -Property {1} -Operator {2} -Expected {3} -DataType {4} -Path {5} -Name {6} -Hive {7} -Key {8} -Is64Bit ${9})' -f
        (& $q $Tree.Source), (& $q $Tree.Property), (& $q $Tree.Operator), (& $q $Tree.Expected), (& $q $Tree.DataType),
        (& $q $Tree.Path), (& $q $Tree.Name), (& $q $Tree.Hive), (& $q $Tree.Key), $(if ($Tree.Is64Bit) { 'true' } else { 'false' }))
}

function New-IntuneExpressionDetectionScript {
    <#
    .SYNOPSIS
        Builds a PowerShell detection script from a stored detection tree.

    .DESCRIPTION
        Intune ANDs every detection rule and has no grouping. A ConfigMgr
        expression with Or, nested groups, or a comparison that no Graph rule
        carries therefore runs as one script. The script follows the Intune
        rule: exit 0 with output means installed.
    #>
    param([Parameter(Mandatory)]$Tree)

    $expression = ConvertTo-ImpScriptExpression -Tree $Tree
    $lines = @(
        '$ErrorActionPreference = ''Stop''',
        $script:ImpDetectionHelpers,
        'try { $detected = ' + $expression + ' } catch { $detected = $false }',
        'if ($detected) { Write-Output ''Detected''; exit 0 }',
        'exit 1'
    )
    return ($lines -join "`r`n")
}

function New-ImpScriptRule {
    param(
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][ValidateSet('detection', 'requirement')][string]$RuleType,
        [string]$DisplayName = '',
        [bool]$RunAs32Bit = $false
    )
    $rule = @{
        '@odata.type'         = '#microsoft.graph.win32LobAppPowerShellScriptRule'
        ruleType              = $RuleType
        enforceSignatureCheck = $false
        runAs32Bit            = $RunAs32Bit
        scriptContent         = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($Text))
        operationType         = 'notConfigured'
        operator              = 'notConfigured'
    }
    if ($RuleType -eq 'requirement') {
        $rule.displayName = $DisplayName
        $rule.runAsAccount = 'system'
        $rule.operationType = 'boolean'
        $rule.operator = 'equal'
        $rule.comparisonValue = 'true'
    }
    return $rule
}

function ConvertTo-IntuneDetectionRules {
    <#
    .SYNOPSIS
        Maps the detection of one deployment type to Graph detection rules.

    .OUTPUTS
        [pscustomobject] Rules (hashtables), Mode (Native, Script,
        ExpressionScript, ProductCode), Findings.
    #>
    param(
        [Parameter(Mandatory)]$Detection,
        [switch]$NativeOnly
    )

    $findings = New-Object System.Collections.Generic.List[object]
    $add = { param($Severity, $Code, $Message) $findings.Add([pscustomobject]@{ Severity = $Severity; Code = $Code; Message = $Message }) }

    switch ($Detection.Method) {
        'ProductCode' {
            $rule = @{
                '@odata.type' = '#microsoft.graph.win32LobAppProductCodeRule'; ruleType = 'detection'
                productCode = $Detection.ProductCode; productVersionOperator = 'notConfigured'; productVersion = $null
            }
            return [pscustomobject]@{ Rules = @($rule); Mode = 'ProductCode'; Findings = @() }
        }
        'Script' {
            if ($Detection.ScriptLanguage -and $Detection.ScriptLanguage -ne 'PowerShell') {
                & $add 'Blocking' 'DetectionLanguageNotSupported' ("The detection script language '{0}' is not supported. Intune runs PowerShell detection scripts only." -f $Detection.ScriptLanguage)
                return [pscustomobject]@{ Rules = @(); Mode = 'Script'; Findings = @($findings.ToArray()) }
            }
            return [pscustomobject]@{ Rules = @((New-ImpScriptRule -Text $Detection.ScriptText -RuleType detection)); Mode = 'Script'; Findings = @() }
        }
        'Enhanced' {
            if ($null -eq $Detection.Tree) {
                & $add 'Blocking' 'DetectionMissing' 'The deployment type stores an empty detection expression.'
                return [pscustomobject]@{ Rules = @(); Mode = 'Native'; Findings = @($findings.ToArray()) }
            }
            $leaves = @(Get-ImpLeaves -Tree $Detection.Tree)
            $unsupported = @($leaves | Where-Object { -not $_.Supported })
            if ($unsupported.Count -gt 0) {
                & $add 'Blocking' 'DetectionNotConvertible' $unsupported[0].Reason
                return [pscustomobject]@{ Rules = @(); Mode = 'Native'; Findings = @($findings.ToArray()) }
            }
            if (Test-ImpTreeIsFlatAnd -Tree $Detection.Tree) {
                $native = @($leaves | ForEach-Object { ConvertTo-ImpNativeRule -Leaf $_ })
                if (@($native | Where-Object { $null -eq $_ }).Count -eq 0) {
                    return [pscustomobject]@{ Rules = $native; Mode = 'Native'; Findings = @() }
                }
            }
            if ($NativeOnly) {
                & $add 'Blocking' 'DetectionNeedsScript' 'The detection uses Or, a group, or a comparison that no native Intune rule carries, and script conversion is turned off.'
                return [pscustomobject]@{ Rules = @(); Mode = 'ExpressionScript'; Findings = @($findings.ToArray()) }
            }
            try {
                $text = New-IntuneExpressionDetectionScript -Tree $Detection.Tree
            }
            catch {
                & $add 'Blocking' 'DetectionNotConvertible' $_.Exception.Message
                return [pscustomobject]@{ Rules = @(); Mode = 'ExpressionScript'; Findings = @($findings.ToArray()) }
            }
            & $add 'Review' 'DetectionConvertedToScript' 'Intune joins all detection rules with AND and has no groups. The stored detection expression runs as one PowerShell script, with the same And, Or and grouping.'
            return [pscustomobject]@{ Rules = @((New-ImpScriptRule -Text $text -RuleType detection)); Mode = 'ExpressionScript'; Findings = @($findings.ToArray()) }
        }
        default {
            & $add 'Blocking' 'DetectionMissing' ("The deployment type has no convertible detection (method '{0}')." -f $Detection.Method)
            return [pscustomobject]@{ Rules = @(); Mode = 'None'; Findings = @($findings.ToArray()) }
        }
    }
}

# ---------------------------------------------------------------------------
# Requirements
# ---------------------------------------------------------------------------

function ConvertFrom-CMPlatformId {
    # 'Windows/All_x64_Windows_10_higher' to architecture and build floor.
    # Returns $null for a platform that Intune Win32 apps do not run on.
    param([Parameter(Mandatory)][string]$PlatformId)

    $arch = ''
    if ($PlatformId -match '_(x64|x86|ARM64)_') { $arch = $Matches[1].ToLowerInvariant() }
    elseif ($PlatformId -match '_AMD64_') { $arch = 'x64' }
    $build = 0
    if ($PlatformId -match 'Windows_11') { $build = 22000 }
    elseif ($PlatformId -match 'Windows_10') { $build = 10240 }
    if ($PlatformId -match 'ARM64_Windows_10' -and $arch -eq '') { $arch = 'arm64' }
    if ($build -eq 0) { return $null }
    if ($PlatformId -notmatch '_higher') { return $null }
    return [pscustomobject]@{ Architecture = $arch; MinimumBuild = $build }
}

function ConvertTo-ImpComparison {
    param([string]$Operator)
    switch ($Operator) {
        'Equals' { return '-eq' } 'NotEquals' { return '-ne' } 'GreaterThan' { return '-gt' }
        'GreaterEquals' { return '-ge' } 'LessThan' { return '-lt' } 'LessEquals' { return '-le' }
        'OneOf' { return '-in' } 'NoneOf' { return '-notin' }
    }
    return ''
}

function ConvertTo-IntuneRequirementModel {
    <#
    .SYNOPSIS
        Reads the requirement rules of a deployment type.

    .DESCRIPTION
        Splits each rule into three groups. Native values go to properties
        that Graph v1.0 carries (architecture, memory, processors). A
        predicate is a rule a generated script can evaluate exactly. Rules
        that are neither are listed as unrecognized; their meaning cannot be
        carried to Intune, and they stop a later variant from proving that it
        excludes this one.

    .OUTPUTS
        [pscustomobject] Architectures (null when unconstrained), Native,
        Predicates, ScriptOnly, Unrecognized, HasRules, Evaluable, Findings.
    #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Requirements)

    $findings = New-Object System.Collections.Generic.List[object]
    $native = @{}
    $predicates = New-Object System.Collections.Generic.List[object]
    $scriptOnly = New-Object System.Collections.Generic.List[object]
    $unrecognized = New-Object System.Collections.Generic.List[object]
    $architectures = $null

    foreach ($rule in $Requirements) {
        switch ($rule.Kind) {
            'OsPlatform' {
                $pairs = @($rule.PlatformIds | ForEach-Object { ConvertFrom-CMPlatformId -PlatformId $_ })
                if ($pairs.Count -eq 0 -or @($pairs | Where-Object { $null -eq $_ }).Count -gt 0) {
                    $unrecognized.Add($rule)
                    $findings.Add([pscustomobject]@{ Severity = 'Review'; Code = 'OsPlatformNotConverted'; Message = ("The operating system requirement '{0}' names a platform that Intune Win32 apps do not run on, or one this importer does not read." -f $rule.DisplayName) })
                    continue
                }
                $negate = ($rule.Operator -eq 'NoneOf')
                $predicates.Add([pscustomobject]@{ Type = 'OsPlatform'; Pairs = $pairs; Negate = $negate })
                if ($negate) {
                    $scriptOnly.Add($predicates[$predicates.Count - 1])
                    continue
                }
                $archSet = @($pairs | ForEach-Object { $_.Architecture } | Select-Object -Unique)
                if ($archSet -contains '') { $archSet = @() }
                $builds = @($pairs | ForEach-Object { $_.MinimumBuild } | Select-Object -Unique)
                if ($archSet.Count -gt 0) {
                    $architectures = $(if ($null -eq $architectures) { $archSet } else { @($architectures | Where-Object { $archSet -contains $_ }) })
                }
                # A uniform Windows 10 floor is the Intune baseline. Mixed floors
                # or a Windows 11 floor need the script to be exact.
                if (@($builds | Where-Object { $_ -gt 10240 }).Count -gt 0) { $scriptOnly.Add($predicates[$predicates.Count - 1]) }
                elseif ($builds.Count -gt 1) { $scriptOnly.Add($predicates[$predicates.Count - 1]) }
            }
            'Global' {
                $name = $rule.LogicalName
                $comparison = ConvertTo-ImpComparison -Operator $rule.Operator
                if (-not $comparison) {
                    $unrecognized.Add($rule)
                    continue
                }
                switch ($name) {
                    'TotalPhysicalMemory' {
                        $bytes = [int64]$rule.Values[0]
                        $predicates.Add([pscustomobject]@{ Type = 'Memory'; Comparison = $comparison; Value = $bytes })
                        if ($rule.Operator -eq 'GreaterEquals') { $native['MinimumMemoryInMB'] = [int][math]::Ceiling($bytes / 1MB) }
                        else { $scriptOnly.Add($predicates[$predicates.Count - 1]) }
                    }
                    'NumberOfProcessors' {
                        $count = [int]$rule.Values[0]
                        $predicates.Add([pscustomobject]@{ Type = 'Processors'; Comparison = $comparison; Value = $count })
                        if ($rule.Operator -eq 'GreaterEquals') { $native['MinimumNumberOfProcessors'] = $count }
                        else { $scriptOnly.Add($predicates[$predicates.Count - 1]) }
                    }
                    'OSLanguage' {
                        $predicates.Add([pscustomobject]@{ Type = 'Language'; Comparison = $comparison; Values = @($rule.Values | ForEach-Object { [int]$_ }) })
                        $scriptOnly.Add($predicates[$predicates.Count - 1])
                    }
                    default {
                        $unrecognized.Add($rule)
                        $findings.Add([pscustomobject]@{ Severity = 'Review'; Code = 'RequirementNotConverted'; Message = ("The requirement '{0}' (condition {1}) has no Intune equivalent in this importer." -f $rule.DisplayName, $name) })
                    }
                }
            }
            default {
                $unrecognized.Add($rule)
                $reason = $(if ($rule.Reason) { $rule.Reason } else { 'The requirement kind is not converted.' })
                $findings.Add([pscustomobject]@{ Severity = 'Review'; Code = 'RequirementNotConverted'; Message = ("The requirement '{0}' is not converted. {1}" -f $rule.DisplayName, $reason) })
            }
        }
    }

    return [pscustomobject]@{
        Architectures = $architectures
        Native        = $native
        Predicates    = @($predicates.ToArray())
        ScriptOnly    = @($scriptOnly.ToArray())
        Unrecognized  = @($unrecognized.ToArray())
        HasRules      = (@($Requirements).Count -gt 0)
        Evaluable     = ($unrecognized.Count -eq 0)
        Findings      = @($findings.ToArray())
    }
}

function ConvertTo-ImpPredicateExpression {
    # One predicate as a PowerShell boolean expression over $os, $cs and $arch.
    param([Parameter(Mandatory)]$Predicate)

    switch ($Predicate.Type) {
        'OsPlatform' {
            $alternatives = @($Predicate.Pairs | ForEach-Object {
                $parts = @()
                if ($_.Architecture) { $parts += ("`$arch -eq '{0}'" -f $_.Architecture) }
                $parts += ('[int]$os.BuildNumber -ge {0}' -f $_.MinimumBuild)
                '(' + ($parts -join ' -and ') + ')'
            })
            $expression = '(' + ($alternatives -join ' -or ') + ')'
            if ($Predicate.Negate) { return '(-not ' + $expression + ')' }
            return $expression
        }
        'Memory' { return ('([int64]$cs.TotalPhysicalMemory {0} {1})' -f $Predicate.Comparison, $Predicate.Value) }
        'Processors' { return ('([int]$cs.NumberOfProcessors {0} {1})' -f $Predicate.Comparison, $Predicate.Value) }
        'Language' { return ('([int]$os.OSLanguage {0} @({1}))' -f $Predicate.Comparison, ($Predicate.Values -join ',')) }
    }
    throw "Unknown predicate type '$($Predicate.Type)'."
}

function New-IntuneVariantRequirementScript {
    <#
    .SYNOPSIS
        Builds the requirement script that keeps "first match wins".

    .DESCRIPTION
        The script prints true when the device passes the predicates of this
        variant that no native Intune property carries, and passes none of the
        requirement sets of the higher-priority variants. Each earlier set is
        the AND of its predicates.
    #>
    param(
        [AllowEmptyCollection()][object[]]$OwnPredicates = @(),
        [AllowEmptyCollection()][object[]]$EarlierSets = @()
    )

    $lines = @(
        '$ErrorActionPreference = ''Stop''',
        '$os = Get-CimInstance -ClassName Win32_OperatingSystem',
        '$cs = Get-CimInstance -ClassName Win32_ComputerSystem',
        '$cpu = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1',
        '$arch = ''x86''',
        'if ([Environment]::Is64BitOperatingSystem) { $arch = $(if ([int]$cpu.Architecture -eq 12) { ''arm64'' } else { ''x64'' }) }'
    )
    $terms = @()
    if (@($OwnPredicates).Count -gt 0) {
        $terms += '(' + ((@($OwnPredicates) | ForEach-Object { ConvertTo-ImpPredicateExpression -Predicate $_ }) -join ' -and ') + ')'
    }
    foreach ($set in @($EarlierSets)) {
        $expressions = @(@($set) | ForEach-Object { ConvertTo-ImpPredicateExpression -Predicate $_ })
        $terms += '(-not (' + ($expressions -join ' -and ') + '))'
    }
    if ($terms.Count -eq 0) { $terms += '$true' }
    $lines += 'try { $selected = ' + ($terms -join ' -and ') + ' } catch { $selected = $false }'
    $lines += 'Write-Output ([string]([bool]$selected)).ToLowerInvariant()'
    return ($lines -join "`r`n")
}

function Get-CMImportVariantExclusion {
    <#
    .SYNOPSIS
        Decides how a variant excludes the variants that rank above it.

    .DESCRIPTION
        ConfigMgr installs the first deployment type whose requirements pass.
        A later variant is therefore selected only when every earlier variant
        is not. An earlier variant with no requirements always passes, so a
        later variant can never be selected (Unreachable). An earlier variant
        with a requirement that cannot be evaluated makes the exclusion
        unprovable. Two variants with disjoint architecture sets exclude each
        other without a script.

    .OUTPUTS
        [pscustomobject] Status (None, Native, Script, Unreachable, Unprovable),
        EarlierSets, Reason.
    #>
    param(
        [Parameter(Mandatory)]$Own,
        [AllowEmptyCollection()][object[]]$Earlier = @()
    )

    if (@($Earlier).Count -eq 0) { return [pscustomobject]@{ Status = 'None'; EarlierSets = @(); Reason = '' } }

    $sets = @()
    foreach ($model in @($Earlier)) {
        if (-not $model.HasRules) {
            return [pscustomobject]@{ Status = 'Unreachable'; EarlierSets = @(); Reason = 'A higher-priority deployment type has no requirements, so ConfigMgr always chooses it first.' }
        }
        if (-not $model.Evaluable) {
            return [pscustomobject]@{ Status = 'Unprovable'; EarlierSets = @(); Reason = 'A higher-priority deployment type has a requirement that cannot be evaluated outside ConfigMgr.' }
        }
        $disjoint = $false
        if ($null -ne $Own.Architectures -and $null -ne $model.Architectures -and @($Own.Architectures).Count -gt 0 -and @($model.Architectures).Count -gt 0) {
            $overlap = @($Own.Architectures | Where-Object { @($model.Architectures) -contains $_ })
            $disjoint = ($overlap.Count -eq 0)
        }
        if (-not $disjoint) { $sets += , @($model.Predicates) }
    }
    if ($sets.Count -eq 0) { return [pscustomobject]@{ Status = 'Native'; EarlierSets = @(); Reason = '' } }
    return [pscustomobject]@{ Status = 'Script'; EarlierSets = $sets; Reason = '' }
}

# ---------------------------------------------------------------------------
# Commands, return codes, setup file
# ---------------------------------------------------------------------------

function ConvertTo-IntuneReturnCodes {
    <#
    .SYNOPSIS
        Maps ConfigMgr exit code classes to Graph return code types.
    #>
    param([AllowEmptyCollection()][object[]]$ExitCodes = @())
    $typeMap = @{ Success = 'success'; SoftReboot = 'softReboot'; HardReboot = 'hardReboot'; FastRetry = 'retry'; Failure = 'failed' }
    $codes = @()
    foreach ($exit in @($ExitCodes)) {
        $type = 'failed'
        if ($typeMap.ContainsKey([string]$exit.Class)) { $type = $typeMap[[string]$exit.Class] }
        $codes += @{ returnCode = [int]$exit.Code; type = $type }
    }
    if ($codes.Count -eq 0) {
        $codes = @(@{ returnCode = 0; type = 'success' }, @{ returnCode = 1707; type = 'success' }, @{ returnCode = 3010; type = 'softReboot' }, @{ returnCode = 1641; type = 'hardReboot' }, @{ returnCode = 1618; type = 'retry' })
    }
    return $codes
}

function Get-ImpFirstToken {
    # The program of a command line, quoted or not.
    param([string]$CommandLine)
    $text = $CommandLine.Trim()
    if ($text.StartsWith('"')) {
        $end = $text.IndexOf('"', 1)
        if ($end -gt 0) { return $text.Substring(1, $end - 1) }
    }
    return ($text -split '\s+', 2)[0]
}

function Get-ImpSetupFile {
    <#
    .SYNOPSIS
        Chooses the setup file the Intune content prep tool needs.

    .DESCRIPTION
        The file the install command runs, or the MSI it installs. When the
        command names no file of the content, the first installer-like file at
        the root of the content is used.
    #>
    param([Parameter(Mandatory)][string]$InstallCommand, [AllowEmptyCollection()][string[]]$ContentFiles = @())

    $files = @($ContentFiles)
    $lookup = @{}
    foreach ($file in $files) { $lookup[$file.ToLowerInvariant()] = $file }

    $program = Get-ImpFirstToken -CommandLine $InstallCommand
    $leaf = [System.IO.Path]::GetFileName($program)
    if ($leaf -and $lookup.ContainsKey($leaf.ToLowerInvariant())) { return $lookup[$leaf.ToLowerInvariant()] }

    foreach ($match in [regex]::Matches($InstallCommand, '(?i)"([^"]+\.(?:msi|msp|exe|bat|cmd|ps1|vbs))"|(\S+\.(?:msi|msp|exe|bat|cmd|ps1|vbs))')) {
        $candidate = [System.IO.Path]::GetFileName(($match.Groups[1].Value + $match.Groups[2].Value))
        if ($candidate -and $lookup.ContainsKey($candidate.ToLowerInvariant())) { return $lookup[$candidate.ToLowerInvariant()] }
    }
    foreach ($extension in '.msi', '.exe', '.bat', '.cmd', '.ps1') {
        $hit = @($files | Where-Object { $_ -notmatch '[\\/]' -and $_.EndsWith($extension, [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1)
        if ($hit.Count -gt 0) { return $hit[0] }
    }
    return ''
}

function ConvertTo-IntuneRestartBehavior {
    param([string]$PostInstallBehavior)
    switch ($PostInstallBehavior) {
        'NoAction' { return 'suppress' }
        'ProgramReboot' { return 'allow' }
        'ForceReboot' { return 'force' }
        default { return 'basedOnReturnCode' }
    }
}

# ---------------------------------------------------------------------------
# The plan
# ---------------------------------------------------------------------------

function New-CMImportPlan {
    <#
    .SYNOPSIS
        Plans the Intune apps for a parsed ConfigMgr application.

    .DESCRIPTION
        Produces one entry per deployment type with its findings, the Graph
        rules, and the flattened manifest that Publish-IntuneWin32App reads.
        Nothing is written. Variant strategy PerDeploymentType keeps "first
        match wins" with requirement rules; AssignByGroup publishes each
        variant with its own requirements only, and the operator assigns
        groups.
    #>
    param(
        [Parameter(Mandatory)]$Model,
        [ValidateSet('PerDeploymentType', 'AssignByGroup')][string]$VariantStrategy = 'PerDeploymentType',
        [switch]$NativeDetectionOnly
    )

    $dts = @($Model.DeploymentTypes)
    $isMulti = ($dts.Count -gt 1)
    $requirementModels = @($dts | ForEach-Object { ConvertTo-IntuneRequirementModel -Requirements @($_.Requirements) })
    $appId = ($Model.LogicalName -replace '^.*/', '')
    $relationships = New-Object System.Collections.Generic.List[object]
    $appFindings = New-Object System.Collections.Generic.List[object]
    $plans = @()

    for ($i = 0; $i -lt $dts.Count; $i++) {
        $dt = $dts[$i]
        $reqModel = $requirementModels[$i]
        $findings = New-Object System.Collections.Generic.List[object]
        $add = { param($Severity, $Code, $Message) $findings.Add([pscustomobject]@{ Severity = $Severity; Code = $Code; Message = $Message }) }
        foreach ($f in $reqModel.Findings) { $findings.Add($f) }

        $intuneName = $(if ($isMulti) { '{0} - {1}' -f $Model.Name, $dt.Name } else { $Model.Name })
        $blocked = $false

        if ($dt.UnsupportedReason) {
            & $add 'Blocking' 'TechnologyNotConverted' $dt.UnsupportedReason
            $blocked = $true
        }

        # Detection
        $detectionResult = $null
        if (-not $blocked) {
            if ($null -eq $dt.Detection) {
                & $add 'Blocking' 'DetectionMissing' 'The deployment type stores no detection.'
                $blocked = $true
            }
            else {
                $detectionResult = ConvertTo-IntuneDetectionRules -Detection $dt.Detection -NativeOnly:$NativeDetectionOnly
                foreach ($f in $detectionResult.Findings) { $findings.Add($f) }
                if (@($detectionResult.Findings | Where-Object { $_.Severity -eq 'Blocking' }).Count -gt 0) { $blocked = $true }
            }
        }

        # Variant selection
        $exclusion = [pscustomobject]@{ Status = 'None'; EarlierSets = @(); Reason = '' }
        if ($isMulti -and $i -gt 0) {
            $exclusion = Get-CMImportVariantExclusion -Own $reqModel -Earlier @($requirementModels[0..($i - 1)])
            switch ($exclusion.Status) {
                'Unreachable' {
                    & $add 'Review' 'UnreachableDeploymentType' ($exclusion.Reason + ' ConfigMgr never installs this deployment type, so it is not converted.')
                    $blocked = $true
                }
                'Unprovable' {
                    if ($VariantStrategy -eq 'PerDeploymentType') {
                        & $add 'Blocking' 'ExclusivityUnprovable' ($exclusion.Reason + ' The app would install on devices where ConfigMgr picks an earlier variant. Use the AssignByGroup strategy and assign each app to its own group.')
                        $blocked = $true
                    }
                }
            }
        }
        if ($isMulti -and $VariantStrategy -eq 'AssignByGroup') {
            & $add 'Review' 'AssignByGroupRequired' 'This app does not exclude the other variants. Assign each variant app to its own device group.'
            $exclusion = [pscustomobject]@{ Status = 'None'; EarlierSets = @(); Reason = '' }
        }

        # Requirement rules, minimums and architecture
        $rules = @()
        $nativeMinimums = @{}
        $allowedArchitectures = ''
        if (-not $blocked) {
            $rules += @($detectionResult.Rules)
            foreach ($key in $reqModel.Native.Keys) { $nativeMinimums[$key] = $reqModel.Native[$key] }
            if ($null -ne $reqModel.Architectures -and @($reqModel.Architectures).Count -gt 0) {
                $allowedArchitectures = (@($reqModel.Architectures) -join ',')
            }
            else {
                # Intune needs an architecture. ConfigMgr named none, so both
                # Windows architectures that ran the deployment type are allowed.
                $allowedArchitectures = 'x86,x64'
                & $add 'Info' 'ArchitectureNotConstrained' 'The deployment type names no architecture. The Intune app allows x86 and x64.'
            }
            $architectureSet = @($allowedArchitectures -split ',' | Sort-Object -Unique)
            if ($architectureSet -contains 'arm64') {
                & $add 'Review' 'Arm64ValueUnverified' 'The Graph v1.0 metadata lists the architectures x86, x64 and arm for allowedArchitectures; the documentation lists arm64. Check the published app on an ARM64 device.'
            }
            if ($architectureSet.Count -gt 1 -and -not ($architectureSet.Count -eq 2 -and $architectureSet -contains 'x86' -and $architectureSet -contains 'x64')) {
                & $add 'Review' 'ArchitectureCombinationUnverified' ('Graph documents one architecture for allowedArchitectures, or x86 with x64 for applicableArchitectures. The combination {0} is sent as allowedArchitectures. Check the published app.' -f ($architectureSet -join ' and '))
            }
            $ownScript = @($reqModel.ScriptOnly)
            if ($ownScript.Count -gt 0 -or $exclusion.Status -eq 'Script') {
                $text = New-IntuneVariantRequirementScript -OwnPredicates $ownScript -EarlierSets @($exclusion.EarlierSets)
                $rules += (New-ImpScriptRule -Text $text -RuleType requirement -DisplayName 'Variant selection')
                if ($exclusion.Status -eq 'Script') {
                    & $add 'Info' 'VariantExclusionScript' 'A requirement script keeps this app off devices where a higher-priority deployment type applies.'
                }
            }
            if ($reqModel.Unrecognized.Count -gt 0) {
                & $add 'Review' 'RequirementsNotTranslated' ("{0} requirement rule(s) have no Intune equivalent. The Intune app does not gate on them." -f $reqModel.Unrecognized.Count)
            }
            if (@($reqModel.Predicates | Where-Object { $_.Type -eq 'OsPlatform' -and @($_.Pairs | Where-Object { $_.MinimumBuild -lt 19044 }).Count -gt 0 }).Count -gt 0) {
                & $add 'Info' 'WindowsFloorRaised' 'ConfigMgr allowed Windows 10 releases that Microsoft no longer supports. The Intune app keeps its minimum release (Windows10_21H2).'
            }
        }

        # Commands and behavior
        $installCommand = $dt.InstallCommand
        $uninstallCommand = $dt.UninstallCommand
        if (-not $blocked -and -not $installCommand) {
            & $add 'Blocking' 'InstallCommandMissing' 'The deployment type stores no install command.'
            $blocked = $true
        }
        if (-not $uninstallCommand -or $dt.UninstallSetting -eq 'NoneRequired') {
            $uninstallCommand = 'cmd.exe /c exit 0'
            if (-not $blocked) { & $add 'Review' 'NoUninstallCommand' 'The deployment type has no uninstall command. The Intune app uses a command that does nothing and reports success.' }
        }
        if ($dt.UninstallContentId -and $dt.ContentId -and $dt.UninstallContentId -ne $dt.ContentId -and -not $blocked) {
            & $add 'Review' 'UninstallContentDiffers' 'The uninstall command uses different content than the install command. Only the install content is copied.'
        }
        if ($dt.RunAs32Bit -and -not $blocked) { & $add 'Info' 'RunAs32BitNotMapped' 'The install command runs as a 32-bit process in ConfigMgr. Intune starts it as a 64-bit process on 64-bit devices.' }
        $systemInteractive = ($dt.RequiresUserInteraction -and $dt.ExecutionContext -ne 'User')
        if ($dt.RequiresUserInteraction -and $dt.ExecutionContext -eq 'User' -and -not $blocked) {
            & $add 'Review' 'UserInteractionPossible' 'The deployment type lets the user see and interact with the install. Intune runs it in the user context and shows what the installer shows.'
        }
        if ($dt.RequiresLogOn -eq 'True' -and $dt.ExecutionContext -ne 'User' -and -not $blocked) {
            & $add 'Review' 'LogonRequirementNotMapped' 'The deployment type installs only when a user is logged on, with system rights. Intune installs without that condition.'
        }
        if ($dt.MaxRuntimeMins -gt 0 -and -not $blocked) {
            & $add 'Info' 'TimingNotApplied' ('The maximum run time ({0} minutes) is recorded in the manifest. Graph v1.0 has no property for it.' -f $dt.MaxRuntimeMins)
        }

        $setupFile = ''
        if (-not $blocked) {
            $setupFile = Get-ImpSetupFile -InstallCommand $installCommand -ContentFiles @($dt.ContentFiles)
            if (-not $setupFile) {
                & $add 'Blocking' 'SetupFileNotFound' 'The content holds no file that the install command runs, and no installer-like file at its root. The Intune content prep tool needs a setup file.'
                $blocked = $true
            }
        }

        # Relationships are reported, never created: Graph v1.0 has no resource for them.
        foreach ($group in @($dt.Dependencies)) {
            foreach ($target in $group.Targets) {
                $relationships.Add([pscustomobject]@{ Type = 'Dependency'; FromDeploymentType = $dt.Name; FromApp = $intuneName; GroupName = $group.GroupName; GroupOperator = $group.Operator; DesiredState = $target.DesiredState; TargetApplication = $target.ApplicationLogical; TargetDeploymentType = $target.DeploymentTypeLogical; TargetApplicationName = $target.ApplicationName; TargetDeploymentTypeName = $target.DeploymentTypeName })
            }
        }
        foreach ($group in @($dt.Supersedes)) {
            foreach ($target in $group.Targets) {
                $relationships.Add([pscustomobject]@{ Type = 'Supersedence'; FromDeploymentType = $dt.Name; FromApp = $intuneName; GroupName = $group.GroupName; GroupOperator = $group.Operator; DesiredState = $target.DesiredState; TargetApplication = $target.ApplicationLogical; TargetDeploymentType = $target.DeploymentTypeLogical; TargetApplicationName = $target.ApplicationName; TargetDeploymentTypeName = $target.DeploymentTypeName })
            }
        }

        $manifest = $null
        if (-not $blocked) {
            $architectureValue = ''
            if ($allowedArchitectures -and $allowedArchitectures -notmatch ',') { $architectureValue = $allowedArchitectures }
            $manifest = [pscustomobject]@{
                AppName                  = $intuneName
                Publisher                = $Model.Publisher
                SoftwareVersion          = $Model.Version
                Description              = $Model.Description
                ApplicationId            = ('cm:' + $appId)
                ProfileId                = ($dt.LogicalName -replace '^.*/', '')
                InstallCommandLine       = $installCommand
                UninstallCommandLine     = $uninstallCommand
                SetupFile                = $setupFile
                InstallationBehaviorType = $(if ($dt.ExecutionContext -eq 'User') { 'InstallForUser' } else { 'InstallForSystem' })
                RequireUserInteraction   = $systemInteractive
                Architecture             = $architectureValue
                Timing                   = [pscustomobject]@{ EstimatedRuntimeMins = $dt.EstimatedRuntimeMins; MaximumRuntimeMins = $dt.MaxRuntimeMins; IntuneApplied = $false }
                IntuneRules              = $rules
                IntuneReturnCodes        = (ConvertTo-IntuneReturnCodes -ExitCodes @($dt.ExitCodes))
                IntuneRestartBehavior    = (ConvertTo-IntuneRestartBehavior -PostInstallBehavior $dt.PostInstallBehavior)
                IntuneMinimums           = $nativeMinimums
                IntuneAllowedArchitectures = $allowedArchitectures
                Source                   = [pscustomobject]@{ Application = $Model.LogicalName; DeploymentType = $dt.LogicalName; Priority = $dt.Priority; ContentLocation = $dt.ContentLocation }
            }
        }

        $plans += [pscustomobject]@{
            Name            = $dt.Name
            LogicalName     = $dt.LogicalName
            Priority        = $dt.Priority
            Technology      = $dt.Technology
            IntuneName      = $intuneName
            Convert         = (-not $blocked)
            ContentLocation = $dt.ContentLocation
            ContentFiles    = @($dt.ContentFiles)
            Exclusion       = $exclusion.Status
            DetectionMode   = $(if ($detectionResult) { $detectionResult.Mode } else { '' })
            Manifest        = $manifest
            Findings        = @($findings.ToArray())
        }
    }

    if ($relationships.Count -gt 0) {
        $appFindings.Add([pscustomobject]@{ Severity = 'Review'; Code = 'RelationshipsNotCreated'; Message = ('{0} dependency and supersedence relationship(s) are listed in relationships.json. Microsoft Graph v1.0 has no resource for them, and App Packager calls v1.0 only. Create them in the Intune admin center.' -f $relationships.Count) })
    }
    if ($isMulti) {
        $appFindings.Add([pscustomobject]@{ Severity = 'Info'; Code = 'VariantApps'; Message = ('The application has {0} deployment types. Each becomes its own Intune app (strategy {1}).' -f $dts.Count, $VariantStrategy) })
    }

    return [pscustomobject]@{
        ApplicationName = $Model.Name
        Publisher       = $Model.Publisher
        Version         = $Model.Version
        LogicalName     = $Model.LogicalName
        Strategy        = $VariantStrategy
        DeploymentTypes = $plans
        Relationships   = @($relationships.ToArray())
        Findings        = @($appFindings.ToArray())
    }
}

function Format-CMImportReport {
    <#
    .SYNOPSIS
        Formats a conversion plan as the text of conversion-report.txt.
    #>
    param([Parameter(Mandatory)]$Plan)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(('ConfigMgr application : {0} {1}' -f $Plan.ApplicationName, $Plan.Version))
    $lines.Add(('Variant strategy      : {0}' -f $Plan.Strategy))
    foreach ($f in $Plan.Findings) { $lines.Add(('[{0}] {1}: {2}' -f $f.Severity, $f.Code, $f.Message)) }
    foreach ($dt in $Plan.DeploymentTypes) {
        $state = $(if ($dt.Convert) { 'converted' } else { 'not converted' })
        $lines.Add('')
        $lines.Add(('Deployment type {0}: {1} ({2}; {3})' -f $dt.Priority, $dt.Name, $dt.Technology, $state))
        if ($dt.Convert) { $lines.Add(('  Intune app          : {0}' -f $dt.IntuneName)) }
        if ($dt.DetectionMode) { $lines.Add(('  Detection           : {0}' -f $dt.DetectionMode)) }
        foreach ($f in $dt.Findings) { $lines.Add(('  [{0}] {1}: {2}' -f $f.Severity, $f.Code, $f.Message)) }
    }
    return ($lines -join "`r`n")
}

# ---------------------------------------------------------------------------
# Site, files, Graph
# ---------------------------------------------------------------------------

function Get-CMApplicationDigest {
    <#
    .SYNOPSIS
        Reads the stored document of an application from the site (read-only).

    .DESCRIPTION
        Call from the CMSite drive after Connect-CMSite. Reads the
        application, not a deployment type.
    #>
    param([Parameter(Mandatory)][string]$ApplicationName)

    $app = @(Get-CMApplication -Name $ApplicationName -DisableWildcardHandling -WarningAction SilentlyContinue -ErrorAction Stop)
    if ($app.Count -eq 0) { throw "ConfigMgr application '$ApplicationName' was not found." }
    if ($app.Count -gt 1) { throw "More than one ConfigMgr application is named '$ApplicationName'." }
    $xml = [string]$app[0].SDMPackageXML
    if ([string]::IsNullOrWhiteSpace($xml)) { throw "ConfigMgr application '$ApplicationName' returned no stored document." }
    return $xml
}

function Resolve-CMImportReferenceNames {
    # Fills the names of dependency and supersedence targets. Best effort:
    # a name that cannot be read stays as the logical name.
    param([Parameter(Mandatory)]$Plan)

    foreach ($relationship in $Plan.Relationships) {
        try {
            $applications = @(Get-CMApplication -Fast -ErrorAction Stop | Where-Object { $_.ModelName -like ('*' + $relationship.TargetApplication) })
            if ($applications.Count -eq 1) {
                $relationship.TargetApplicationName = [string]$applications[0].LocalizedDisplayName
                $deploymentTypes = @(Get-CMDeploymentType -ApplicationName $relationship.TargetApplicationName -ErrorAction Stop | Where-Object { $_.ModelName -like ('*' + $relationship.TargetDeploymentType) })
                if ($deploymentTypes.Count -eq 1) { $relationship.TargetDeploymentTypeName = [string]$deploymentTypes[0].LocalizedDisplayName }
            }
        }
        catch {
            Write-ImpLog ("Could not resolve the name of {0}: {1}" -f $relationship.TargetApplication, $_.Exception.Message) 'WARN'
        }
    }
}

function Copy-CMImportContent {
    param([Parameter(Mandatory)][string]$Source, [Parameter(Mandatory)][string]$Destination)
    # The copy mirrors the source, so a repeat run never keeps a file that the
    # site content no longer holds. Only a folder this module names Content is mirrored.
    if ((Split-Path -Leaf $Destination) -ne 'Content') { throw "The content copy refuses the destination '$Destination': its folder is not named Content." }
    New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    $output = & robocopy.exe $Source $Destination /MIR /R:2 /W:2 /NFL /NDL /NJH /NJS /NP 2>&1
    if ($LASTEXITCODE -ge 8) {
        throw ("Copying the content of '{0}' failed (robocopy exit {1}): {2}" -f $Source, $LASTEXITCODE, (($output | Select-Object -Last 3) -join ' '))
    }
}

function Export-CMImportStage {
    <#
    .SYNOPSIS
        Writes the plan to disk: one folder per converted deployment type.

    .DESCRIPTION
        Each folder holds Content (a copy of the deployment type's content),
        stage-manifest.json, and later the .intunewin. The application folder
        holds conversion-report.txt, conversion-plan.json and
        relationships.json. The site is not changed.
    #>
    param(
        [Parameter(Mandatory)]$Plan,
        [Parameter(Mandatory)][string]$OutputRoot,
        [switch]$SkipContent,
        [string]$IconBase64 = ''
    )

    $appFolder = Join-Path $OutputRoot (ConvertTo-ImpSafeName $Plan.ApplicationName)
    New-Item -ItemType Directory -Force -Path $appFolder | Out-Null
    $staged = @()
    foreach ($dt in $Plan.DeploymentTypes) {
        if (-not $dt.Convert) { continue }
        $folder = Join-Path $appFolder (ConvertTo-ImpSafeName $dt.Name)
        $contentFolder = Join-Path $folder 'Content'
        if (-not $SkipContent) {
            if (-not $dt.ContentLocation) { throw "Deployment type '$($dt.Name)' has no content location." }
            Write-ImpLog ("Copying content for '{0}' from {1}" -f $dt.Name, $dt.ContentLocation)
            Copy-CMImportContent -Source $dt.ContentLocation -Destination $contentFolder
            if (-not (Test-Path -LiteralPath (Join-Path $contentFolder $dt.Manifest.SetupFile))) {
                throw ("The setup file '{0}' is not in the copied content of '{1}'." -f $dt.Manifest.SetupFile, $dt.Name)
            }
        }
        else {
            New-Item -ItemType Directory -Force -Path $contentFolder | Out-Null
        }
        $manifest = $dt.Manifest
        $manifest | Add-Member -NotePropertyName ContentRoot -NotePropertyValue $contentFolder -Force
        if ($IconBase64) {
            $iconFile = Join-Path $folder 'app-icon.png'
            [System.IO.File]::WriteAllBytes($iconFile, [Convert]::FromBase64String($IconBase64))
            $manifest | Add-Member -NotePropertyName Icon -NotePropertyValue 'app-icon.png' -Force
        }
        $manifestPath = Join-Path $folder 'stage-manifest.json'
        [System.IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding($false)))
        $staged += [pscustomobject]@{ Name = $dt.Name; IntuneName = $dt.IntuneName; Folder = $folder; ContentFolder = $contentFolder; ManifestPath = $manifestPath; SetupFile = $manifest.SetupFile }
    }

    [System.IO.File]::WriteAllText((Join-Path $appFolder 'conversion-report.txt'), (Format-CMImportReport -Plan $Plan), (New-Object System.Text.UTF8Encoding($false)))
    [System.IO.File]::WriteAllText((Join-Path $appFolder 'relationships.json'), (ConvertTo-Json -InputObject @($Plan.Relationships) -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
    $summary = $Plan | Select-Object ApplicationName, Publisher, Version, LogicalName, Strategy, Findings, @{ n = 'DeploymentTypes'; e = { @($_.DeploymentTypes | Select-Object Name, LogicalName, Priority, Technology, IntuneName, Convert, Exclusion, DetectionMode, Findings) } }
    [System.IO.File]::WriteAllText((Join-Path $appFolder 'conversion-plan.json'), ($summary | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding($false)))

    return [pscustomobject]@{ ApplicationFolder = $appFolder; Staged = $staged }
}

function Initialize-CMImportDependencies {
    # The package and publish steps live in AppPackagerCommon. A caller that
    # imported only this module gets the common module from the same folder.
    if ((Get-Command -Name Publish-IntuneWin32App -ErrorAction SilentlyContinue) -and (Get-Command -Name New-IntuneWinPackage -ErrorAction SilentlyContinue)) { return }
    $common = Join-Path $PSScriptRoot 'AppPackagerCommon.psd1'
    if (-not (Test-Path -LiteralPath $common)) { throw "AppPackagerCommon.psd1 was not found beside AppPackagerCmImport.psm1." }
    Import-Module $common -DisableNameChecking -ErrorAction Stop
}

function Convert-CMApplicationToIntune {
    <#
    .SYNOPSIS
        Converts one ConfigMgr application to Intune Win32 apps.

    .DESCRIPTION
        Reads the application (read-only), plans one Intune app per
        deployment type, copies the content, writes the reports, and
        optionally builds the .intunewin files and publishes them through
        Microsoft Graph v1.0. Without -Package the run stops after staging.
        Run it where the ConfigMgr console cmdlets and the site drive are
        available.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationName,
        [Parameter(Mandatory)][string]$OutputRoot,
        [string]$SiteCode,
        [ValidateSet('PerDeploymentType', 'AssignByGroup')][string]$VariantStrategy = 'PerDeploymentType',
        [switch]$NativeDetectionOnly,
        [switch]$SkipContent,
        [switch]$Package,
        [string]$ToolPath = '',
        [switch]$Publish,
        [string]$TenantId = '',
        [string]$ClientId = '',
        [string]$ClientSecret = '',
        [string]$Xml = ''
    )

    if ($Publish -and -not $Package) { $Package = $true }
    Initialize-CMImportDependencies
    $fromSite = (-not $Xml)
    if ($fromSite) {
        if ($SiteCode -and (Get-Command -Name Connect-CMSite -ErrorAction SilentlyContinue)) {
            if (-not (Connect-CMSite -SiteCode $SiteCode)) { throw "Could not connect to ConfigMgr site '$SiteCode'." }
        }
        $Xml = Get-CMApplicationDigest -ApplicationName $ApplicationName
    }
    $model = ConvertFrom-CMApplicationDigest -Xml $Xml
    $plan = New-CMImportPlan -Model $model -VariantStrategy $VariantStrategy -NativeDetectionOnly:$NativeDetectionOnly
    if ($fromSite) {
        try { Resolve-CMImportReferenceNames -Plan $plan } catch { Write-ImpLog ("Reference names stay as logical names: {0}" -f $_.Exception.Message) 'WARN' }
    }

    $stage = Export-CMImportStage -Plan $plan -OutputRoot $OutputRoot -SkipContent:$SkipContent -IconBase64 $model.IconBase64
    $published = @()
    if ($Package) {
        if ($SkipContent) { throw '-Package needs the content. Run without -SkipContent.' }
        if (-not $ToolPath) {
            $toolFolder = Join-Path $OutputRoot '_tools'
            $ToolPath = Install-IntuneWinAppUtil -DestinationFolder $toolFolder
        }
        foreach ($item in $stage.Staged) {
            $built = New-IntuneWinPackage -ToolPath $ToolPath -ContentFolder $item.ContentFolder -SetupFile $item.SetupFile -OutputFolder (Join-Path $item.Folder 'Package') -OutputName ((ConvertTo-ImpSafeName $item.Name) + '.intunewin')
            $record = [pscustomobject]@{ Name = $item.Name; IntuneWinPath = $built.IntuneWinPath; AppId = '' }
            if ($Publish) {
                if (-not ($TenantId -and $ClientId -and $ClientSecret)) { throw '-Publish needs TenantId, ClientId and ClientSecret.' }
                $manifest = Get-Content -LiteralPath $item.ManifestPath -Raw | ConvertFrom-Json
                $record.AppId = Publish-IntuneWin32App -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -IntuneWinPath $built.IntuneWinPath -Manifest $manifest -Description ([string]$model.Description)
            }
            $published += $record
        }
    }
    return [pscustomobject]@{ Plan = $plan; Stage = $stage; Packages = $published }
}

Export-ModuleMember -Function ConvertFrom-CMApplicationDigest, ConvertTo-IntuneDetectionRules, New-IntuneExpressionDetectionScript, ConvertTo-IntuneRequirementModel, New-IntuneVariantRequirementScript, Get-CMImportVariantExclusion, ConvertTo-IntuneReturnCodes, Get-ImpSetupFile, New-CMImportPlan, Format-CMImportReport, Get-CMApplicationDigest, Export-CMImportStage, Convert-CMApplicationToIntune
