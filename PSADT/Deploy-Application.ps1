<#
.SYNOPSIS
	This script performs the installation or uninstallation of an application(s).
.DESCRIPTION
	The script is provided as a template to perform an install or uninstall of an application(s).
	The script either performs an "Install" deployment type or an "Uninstall" deployment type.
	The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.
	The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.
.PARAMETER DeploymentType
	The type of deployment to perform. Default is: Install.
.PARAMETER DeployMode
	Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.
.PARAMETER AllowRebootPassThru
	Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.
.PARAMETER TerminalServerMode
	Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Destkop Session Hosts/Citrix servers.
.PARAMETER DisableLogging
	Disables logging to file for the script. Default is: $false.
.EXAMPLE
	Deploy-Application.ps1
.EXAMPLE
	Deploy-Application.ps1 -DeployMode 'Silent'
.EXAMPLE
	Deploy-Application.ps1 -AllowRebootPassThru -AllowDefer
.EXAMPLE
	Deploy-Application.ps1 -DeploymentType Uninstall
.NOTES
	Toolkit Exit Code Ranges:
	60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
	69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
	70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK
	http://psappdeploytoolkit.com
#>
[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false)]
	[ValidateSet('Install','Uninstall')]
	[string]$DeploymentType = 'Install',
	[Parameter(Mandatory=$false)]
	[ValidateSet('Interactive','Silent','NonInteractive')]
	[string]$DeployMode = 'Interactive',
	[Parameter(Mandatory=$false)]
	[switch]$AllowRebootPassThru = $false,
	[Parameter(Mandatory=$false)]
	[switch]$TerminalServerMode = $false,
	[Parameter(Mandatory=$false)]
	[switch]$DisableLogging = $false,
	[Parameter(Mandatory=$false)]
	[string]$DeploymentJson = 'stage-manifest.json'
)

Try {

	## Variables: Exit Code
	[int32]$mainExitCode = 0
	
	## Set the script execution policy for this process
	Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch {}
	
	##*===============================================
	##* VARIABLE DECLARATION
	##*===============================================
	## Variables: Application
	Try {
		$scriptpath = Split-Path -parent $MyInvocation.MyCommand.Definition
		$JsonConfig = Get-Content -Path ($scriptpath + "\" + $DeploymentJson) -ErrorAction Stop -Raw | ConvertFrom-Json
	}
	Catch {
		If ($mainExitCode -eq 0){ [int32]$mainExitCode = 60008 }
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		## Exit the script, returning the exit code to SCCM
		If (Test-Path -LiteralPath 'variable:HostInvocation') { $script:ExitCode = $mainExitCode; Exit } Else { Exit $mainExitCode }
	}
	
	$App = [PSCustomObject]@{
		Vendor          = $JsonConfig.Publisher
		Name            = $JsonConfig.AppName
		Version         = $JsonConfig.SoftwareVersion
		Arch            = $JsonConfig.Architecture
		Lang            = $JsonConfig.Language
		Revision        = $JsonConfig.Revision
		InstallerFile   = $JsonConfig.InstallerFile
		InstallerType   = $JsonConfig.InstallerType
		InstallArgs     = $JsonConfig.InstallArgs
		UninstallFile   = $JsonConfig.UninstallCommand
		UninstallArgs   = $JsonConfig.UninstallArgs
		Processes       = $JsonConfig.RunningProcess
		Detection       = $JsonConfig.Detection
	}

	if ([string]::IsNullOrWhiteSpace($App.Arch)) { $App.Arch = 'x64' }
	if ([string]::IsNullOrWhiteSpace($App.Lang)) { $App.Lang = 'MUI' }
	if ([string]::IsNullOrWhiteSpace($App.Revision)) { $App.Revision = '01' }

	[string]$appNameWithoutVersion = $App.Name
	if (-not [string]::IsNullOrWhiteSpace($App.Version)) {
		$appNameWithoutVersion = ($App.Name-replace [regex]::Escape($App.Version), "").Trim()
	}
	if ([string]::IsNullOrWhiteSpace($appNameWithoutVersion)) {
		$appNameWithoutVersion = $App.Name
	}
	[string]$appNameMsg = $appNameWithoutVersion
	[string]$dirFiles = Join-Path -Path $scriptpath -ChildPath 'Files'
	[bool]$BlockExecution = $false
	[string]$StartProcessAsUser = ''
	[string]$StartProcessAsUserParam = ''
	[string]$appUninstallArgs = ''

	if ([string]::IsNullOrWhiteSpace($App.InstallerType) -or [string]::IsNullOrWhiteSpace($App.InstallerFile)) {
		throw "Manifest [$DeploymentJson] is missing required values: InstallerType and/or InstallerFile."
	}
	
	## Package Name
    [String]$PackageName = $App.Vendor + "_" + $appNameWithoutVersion + "_" + $App.Version + "_" + $App.Lang + "_" + $App.Arch + "_" + $App.Revision

	## Custom Registry Entry
	[String]$CustomRegKey = 'HKLM:\SOFTWARE\SCCM\' + $PackageName
	
	[string]$appScriptVersion = '1.4.0'
	[string]$appScriptDate = '20/04/2026'
	[string]$appScriptAuthor = 'SH'
	##*===============================================
	
	##* Do not modify section below
	#region DoNotModify
	
	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.6.5'
	[string]$deployAppScriptDate = '08/17/2015'
	[hashtable]$deployAppScriptParameters = $psBoundParameters
	
	## Variables: Environment
	If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }
	[string]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent
	
	## Dot source the required App Deploy Toolkit Functions
	Try {
		[string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
		If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) { Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]." }
		If ($DisableLogging) { . $moduleAppDeployToolkitMain -DisableLogging } Else { . $moduleAppDeployToolkitMain }
	}
	Catch {
		If ($mainExitCode -eq 0){ [int32]$mainExitCode = 60008 }
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		## Exit the script, returning the exit code to SCCM
		If (Test-Path -LiteralPath 'variable:HostInvocation') { $script:ExitCode = $mainExitCode; Exit } Else { Exit $mainExitCode }
	}
	
	#endregion
	##* Do not modify section above
	##*===============================================
	##* END VARIABLE DECLARATION
	##*===============================================
	
	
	##*===============================================
	##* OWN FUNCTIONS DECLARATION
	##*===============================================
	## Function: Select-ProcessName
	Function Select-ProcessName {

		[CmdletBinding()]
		Param (
			[Array]$processRunning
		)

		return ($processRunning | Select-Object ProcessName, Description | ForEach-Object { 
			If ([string]::IsNullOrEmpty($_.Description)) 
			{ 
				"$($_.Processname)=$($_.Processname)"
			} elseif ([string]::IsNullOrEmpty($_.Description.Trim()))  {
				"$($_.Processname)=$($_.Processname)"
			}
			else { 
			"$($_.Processname)=$($_.Description)" 
			} 

			}) -join ","
	}
	#endregion
	##*===============================================
	##* END OWN FUNCTIONS DECLARATION
	##*===============================================
	
	if($App.Processes){
		[array]$processNames = @()
		if ($App.Processes -is [array]) {
			$processNames = $App.Processes
		}
		else {
			$processNames = @("$($App.Processes)" -split ',')
		}
		$processNames = $processNames | ForEach-Object { "$($_)".Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
		if ($processNames.Count -gt 0) {
			$processRunning = (Get-Process -Name $processNames -ErrorAction SilentlyContinue)
		}
	}
	
	If ($deploymentType -ine 'Uninstall') {
		##*===============================================
		##* PRE-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Installation'
		
		## <Perform Pre-Installation tasks here>
		
		## Prompt the user to close the following applications if they are running:
		if($processRunning){
			Show-InstallationWelcome -CloseApps (Select-ProcessName -processRunning $processRunning) -AllowDefer -DeferTimes 3 -MinimizeWindows $false -BlockExecution
			# Show Progress Message (with the default message)
			Show-InstallationProgress -StatusMessage "Installing $appNameWithoutVersion $($App.Version)" -TopMost $false
		}

	    # Remove any previous versions if explicit uninstall commands are provided in manifest
        Remove-MSIApplications -Name $appNameWithoutVersion -Wildcard -ContinueOnError $true

        Get-InstalledApplication -name "*$($appNameWithoutVersion)*" -WildCard | ForEach-Object { 

			$prod = $_.ProductCode
			
			if ($prod -eq "") {
			    ## some installer e.g. WEB-Installer
				$unstr = $_.UninstallString + " --force-uninstall"
				$unstring = $unstr.split('"')
				Write-Log -Message "start Uninstall $($unString[1]) $($unString[2])" -LogType 'CMTrace' 
				#Execute-Process -Path $unString[1] -Parameters $unString[2] -ContinueOnError $True
				$erg = start-process $unString[1] -arg $unString[2] -Wait
				Write-Log -Message "$appNameWithoutVersion has been removed with" -LogType 'CMTrace'
			}
			else {
			    ## MSI Installation!
				Write-Log -Message "start MSI Uninstall $($prod)" -LogType 'CMTrace'
				Execute-MSI -Action 'Uninstall' -Path $prod -Parameters "/qn" -ContinueOnError $True
				Write-Log -Message "$appNameWithoutVersion has been removed" -LogType 'CMTrace'
			}
		}
				
		## Clean up Branding Keys
		If (Test-Path 'HKLM:\SOFTWARE\SCCM') {
			Get-ChildItem -Path 'HKLM:\SOFTWARE\SCCM' -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -match ($App.Vendor + "_" + $appNameWithoutVersion + "_") } | Remove-Item -Force -ErrorAction SilentlyContinue
		}
		
		##*===============================================
		##* INSTALLATION
		##*===============================================
		[string]$installPhase = 'Installation'
		Write-Log -Message "$($App.InstallerFile)"
		switch ([System.IO.Path]::GetExtension($App.InstallerFile.ToUpper())) {
			".MSI" {
				Write-Log -Message "Installing MSI file"
				$installerPath = Join-Path -Path $dirFiles -ChildPath $App.InstallerFile
				if (-not (Test-Path -LiteralPath $installerPath)) {
					throw "MSI file not found: $installerPath"
				}
				if([string]::IsNullOrEmpty($App.InstallArgs)){
					Execute-MSI -Action Install -Path $installerPath
				} else {
					Execute-MSI -Action Install -Path $installerPath -Parameters $App.InstallArgs
				}
			}
			default {
				Write-Log -Message "Installing Exe file"
				$installerPath = Join-Path -Path $dirFiles -ChildPath $App.InstallerFile
				if(-not (Test-Path -LiteralPath $installerPath)){
					throw "Installer file not found: $installerPath"
				}
				if([string]::IsNullOrEmpty($App.InstallArgs)){
					Write-Log -Message "Found $($App.InstallerFile), now attempting to install."
					Execute-Process -Path $installerPath
				} else {
					Write-Log -Message "Found $($App.InstallerFile), now attempting to install $($App.Name) with arguments $($App.InstallArgs)."
					Execute-Process -Path $installerPath -Parameters "$($App.InstallArgs)"
				}
			}
			
		}
		
		
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		## <Perform Post-Installation tasks here>
		
		## Set custom registry keys
		Set-RegistryKey -Key $CustomRegKey -Name 'Installed'       	-Type 'String' -Value 'True'
		Set-RegistryKey -Key $CustomRegKey -Name 'Date'            	-Type 'String' -Value (Get-Date -Format g)
		Set-RegistryKey -Key $CustomRegKey -Name 'Vendor'      		-Type 'String' -Value $App.Vendor
		Set-RegistryKey -Key $CustomRegKey -Name 'ApplicationName'  -Type 'String' -Value $App.Name
		Set-RegistryKey -Key $CustomRegKey -Name 'Version'         	-Type 'String' -Value $App.Version
		Set-RegistryKey -Key $CustomRegKey -Name 'Language'        	-Type 'String' -Value $App.Lang
		Set-RegistryKey -Key $CustomRegKey -Name 'Architecture'    	-Type 'String' -Value $App.Arch
		
		## If block execution variable is true, call the function to unblock execution
	    If ($BlockExecution) { Unblock-AppExecution }
		
		## Display a message at the end of the install
		#Show-InstallationPrompt -Message 'You can customise text to appear at the end of an install or remove it completely for unattended installations.' -ButtonRightText 'OK' -Icon Information -NoWait
		if($processRunning){
			[boolean]$configShowBalloonNotifications = $true
			Show-BalloonTip -BalloonTipIcon 'None' -BalloonTipText 'Installation Finished' -BalloonTipTitle $appNameMsg -BalloonTipTime 1000
			[boolean]$configShowBalloonNotifications = $false
		}
	}
	ElseIf ($deploymentType -ieq 'Uninstall')
	{
		##*===============================================
		##* PRE-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Uninstallation'
		
		## <Perform Pre-Uninstallation tasks here>
			
		## Prompt the user to close the following applications if they are running:
		if($processRunning){
			Show-InstallationWelcome -CloseApps (Select-ProcessName -processRunning $processRunning) -AllowDefer -DeferTimes 3 -MinimizeWindows $false -BlockExecution
			Show-InstallationProgress -StatusMessage "Uninstalling $appNameMsg $($App.Version)" -TopMost $false
		}
		
		## Show Progress Message (with a message to indicate the application is being uninstalled)
		#Show-InstallationProgress -StatusMessage 'Uninstalling application [$installTitle]. Please Wait...'
		
		
		##*===============================================
		##* UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Uninstallation'

		Remove-MSIApplications -Name $appNameWithoutVersion -Wildcard -ContinueOnError $true

		Get-InstalledApplication -name "*$($appNameWithoutVersion)*" -WildCard | ForEach-Object { 

			$prod = $_.ProductCode
			
			if ($prod -eq "") {
			    ## some installer e.g. WEB-Installer
				$unstr = $_.UninstallString + " --force-uninstall"
				$unstring = $unstr.split('"')
				Write-Log -Message "start Uninstall $($unString[1]) $($unString[2])" -LogType 'CMTrace' 
				#Execute-Process -Path $unString[1] -Parameters $unString[2] -ContinueOnError $True
				$erg = start-process $unString[1] -arg $unString[2] -Wait
				Write-Log -Message "$appNameWithoutVersion has been removed with" -LogType 'CMTrace'
			}
			else {
			    ## MSI Installation!
				Write-Log -Message "start MSI Uninstall $($prod)" -LogType 'CMTrace'
				Execute-MSI -Action 'Uninstall' -Path $prod -Parameters "/qn" -ContinueOnError $True
				Write-Log -Message "$appNameWithoutVersion has been removed" -LogType 'CMTrace'
			}
		}

        
		
		##*===============================================
		##* POST-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Uninstallation'
		
		## <Perform Post-Uninstallation tasks here>
		
	}
	
	##*===============================================
	##* END SCRIPT BODY
	##*===============================================
	
	## Call the Exit-Script function to perform final cleanup operations
	if($processRunning){
		Exit-Script -ExitCode $mainExitCode
	} else {
		[boolean]$deployModeSilent = $true
		Exit-Script -ExitCode $mainExitCode
	}
}
Catch {
	[int32]$mainExitCode = 1
	[string]$mainErrorMessage = "$(Resolve-Error)"
	Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
	#Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
	Exit-Script -ExitCode $mainExitCode
}
