<#
.SYNOPSIS
	This script performs the installation or uninstallation of an application(s).
	# LICENSE #
	PowerShell App Deployment Toolkit - Provides a set of functions to perform common application deployment tasks on Windows.
	Copyright (C) 2017 - Sean Lillis, Dan Cunningham, Muhammad Mashwani, Aman Motazedian.
	This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
	You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.
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
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeployMode 'Silent'; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -AllowRebootPassThru; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeploymentType 'Uninstall'; Exit $LastExitCode }"
.EXAMPLE
    Deploy-Application.exe -DeploymentType "Install" -DeployMode "Silent"
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
	[Parameter(Mandatory = $false)]
	[ValidateSet('Install', 'Uninstall', 'Repair')]
	[string]$DeploymentType = 'Install',
	[Parameter(Mandatory = $false)]
	[ValidateSet('Interactive', 'Silent', 'NonInteractive')]
	[string]$DeployMode = 'Silent',
	[Parameter(Mandatory = $false)]
	[switch]$AllowRebootPassThru = $false,
	[Parameter(Mandatory = $false)]
	[switch]$TerminalServerMode = $false,
	[Parameter(Mandatory = $false)]
	[switch]$DisableLogging = $false
)

Try {
	## Set the script execution policy for this process
	Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch {}

	##*===============================================
	##* VARIABLE DECLARATION
	##*===============================================
	##* Variables: Application
	[string]$appVendor = 'Tuolumne County'
	[string]$appName = 'Clerk of the Board Records Page'
	[string]$appVersion = '1.0.0'
	[string]$appArch = 'x64'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.1'
	[string]$appScriptDate = '2022.11.08'
	[string]$appScriptAuthor = 'JBSMITH'

	##* Variables: Install Titles (Only set here to override defaults set by the toolkit)
	[string]$installName = "$appName ($appVersion)"
	[string]$installTitle = "$appName ($appVersion)"

	##*===============================================
	##* ANCHOR: VARIABLES - Template
	##* Changeable Array(s)/Variable(s)
	##*===============================================
	# Template array(s)/variable(s) used within the PSADT.

	# 32-bit application
	# 32-bit application install name.
	$32bitAppInstallName = "Clerk of the Board Records Page" 
	# Application install Path.
	$32bitAppInstallPath = "${Env:ProgramFiles}\Tuolumne County\Clerk of the Board Records Page"
	# Application install parameters.
	#$32bitAppInstallParam = ""

	##* Remove Application Names 
	# Mainly used in the Pre-Installation, Pre-Uninstallation, Uninstallation and Post-Uninstallation phases.
	# These scalable Array(s)/Variable(s) are used to remove previous application(s) by name.
	#$RemoveAppNamesMSI = @("")
	#$RemoveAppNamesEXE = @("")

	##* Application uninstall parameters.
	# 64-bit application
	# 64-bit application uninstall name.
	#$64bitAppUninstallName = ""
	# Application uninstall path.
	#$64bitAppUninstallPath = ""
	# Application uninstall parameters.
	#$64bitAppUninstallParam = ""

	# 32-bit application
	# 32-bit application uninstall name.
	$32bitAppUninstallName = "Clerk of the Board Records Page"
	# Application uninstall path.
	#$32bitAppUninstallPath = ""
	# Application uninstall parameters.
	#$32bitAppUninstallParam = ""
    
	##* Application Settings File Name
	# Names of files used for application settings.
	#[string[]]$appSettingsNames = @("")

	##* Application Settings Directory
	# Directory where application settings be reside.
	#[string[]]$appSettingsDirs = @("")
    
	## Set variables to match script variables
	# These Variable(s) keep the spaces the PSADT script removes. These can and are used in titles, messages, logs and the PIRK information for the application being installed.
	$apVendor = $appVendor
	$apName = $appName
	$apversion = $appVersion
	$apScriptVersion = $appScriptVersion

	##*===============================================
	##* ANCHOR: VARIABLES - Author
	##* Changeable Array(s)/Variable(s)
	##*===============================================
	# If the template array(s)/variable(s) aren't enough, add more array(s)/variable(s) here.

	[string[]]$appShortcutDirs = @("${Env:ProgramFiles}\Tuolumne County\Clerk of the Board Records Page", "${Env:ProgramData}\Microsoft\Windows\Start Menu\Programs", "${Env:PUBLIC}\Desktop")
	[string[]]$removeShortcutDirs = @("${Env:ProgramFiles}\Tuolumne County", "${Env:ProgramData}\Microsoft\Windows\Start Menu\Programs", "${Env:PUBLIC}\Desktop")

	##*===============================================
	##* Do not modify section below
	#region DoNotModify

	## Variables: Exit Code
	[int32]$mainExitCode = 0

	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.8.3'
	[string]$deployAppScriptDate = '30/09/2020'
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
		If ($mainExitCode -eq 0) { [int32]$mainExitCode = 60008 }
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		## Exit the script, returning the exit code to SCCM
		If (Test-Path -LiteralPath 'variable:HostInvocation') { $script:ExitCode = $mainExitCode; Exit } Else { Exit $mainExitCode }
	}

	#endregion
	##* Do not modify section above
	##*===============================================
	##* END VARIABLE DECLARATION
	##*===============================================

	If ($deploymentType -ine 'Uninstall' -and $deploymentType -ine 'Repair') {
		##*===============================================
		##* ANCHOR: PRE-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Installation'

		# For each item in the array...
		# Check if previous versions of application install directory exist.
		# If application install directory exists, remove previous versions of application install directory.
		# If removal failed, log results. Exit script.
		# Else, log results from check.
		Foreach ($removeShortcutDir in $removeShortcutDirs) {
			# Check if previous versions of application install directory exist.
			If (Test-Path -Path "$removeShortcutDir\$32bitAppUninstallName*") {
				Write-Log -Message "Previous versions of $32bitAppUninstallName exist in $removeShortcutDir. Removing..."  
				# Remove previous versions of application install directory.
				Try {
					Remove-Item -Path "$removeShortcutDir\$32bitAppUninstallName*" -Force -Recurse
					Remove-Item -Path "${Env:PUBLIC}\Desktop\$32bitAppUninstallName*" -Force
				}
				# If removal failed, log results. Exit script.
				Catch [System.Exception] {
					Write-Log -Message "Removing previous versions of $32bitAppUninstallName in $removeShortcutDir failed with error: $_."
					Write-Log -Message "Exiting script with error."
					Exit-Script -ExitCode 1627
				}
				Write-Log -Message "Removing previous versions of $32bitAppUninstallName in $removeShortcutDir complete."
			}
			# Else, log results from check. 
			Else {  
				Write-Log -Message "Previous versions of $32bitAppUninstallName in $removeShortcutDir do not exist."        
			}
		}

		# Update (refresh) the desktop.
		Update-Desktop

		# Check if previous versions of package information registry key (PIRK) exist. 
		# If package information registry key (PIRK) exists, remove previous versions of package information registry key (PIRK).
		# If removal failed, log results. Exit script. 
		# Else, log results from check.
		If (Test-Path -Path "HKLM:\SOFTWARE\Tuolumne County\Package Information\$apName*") { 
			# Remove previous versions of package information registry key (PIRK).
			Try {
				Write-Log -Message "Previous versions of package information registry key (PIRK) exist. Removing..."
				Remove-Item -Path "HKLM:\SOFTWARE\Tuolumne County\Package Information\$apName*" -Force
				Write-Log -Message "Removing previous versions of package information registry key (PIRK) complete."
			}
			# If removal failed, log results. Exit script. 
			Catch [System.Exception] {
				Write-Log -Message "Removing previous versions of package information registry key (PIRK) failed with error: $_"
				Write-Log -Message "Exiting script with error."
				Exit-Script -ExitCode 1627
			}
		}
		# Else, log results from check. 
		Else { 
			Write-Log -Message "Previous versions of package information registry key (PIRK) do not exist."  
		}

		##*===============================================
		##* ANCHOR: INSTALLATION
		##*===============================================
		[string]$installPhase = 'Installation'

		## Handle Zero-Config MSI Installations
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Install'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
			Execute-MSI @ExecuteDefaultMSISplat; If ($defaultMspFiles) { $defaultMspFiles | ForEach-Object { Execute-MSI -Action 'Patch' -Path $_ } }
		}

		# Install application(s). 
		# If install failed, log results. Exit script.
		Try {
			Write-log -Message "Installing $apName ($apversion)."
			Copy-Item -Path "$dirFiles\$32bitAppInstallName" -Destination "$32bitAppInstallPath" -Force -Recurse
			Write-Log -Message "Installing $apName ($apversion) complete."
		}
		# If install failed, log results. Exit script.
		Catch [System.Exception] {
			Write-Log -Message "Installing $apName ($apversion) failed with error: $_."
			Write-Log -Message "Exiting script with error."
			Exit-Script -ExitCode 1627
		}

		##* Every package should have a package information registry key (PIRK), which details what the $apversion and $apScriptVErsion are, along with any other information.
		# Create package information registry key (PIRK).
		# If creation failed, log results. Exit script.
		Try {
			Write-Log -Message "Creating package information registry key (PIRK)."
			Set-RegistryKey -Key "HKLM:\Software\Tuolumne County\Package Information" -Name "Readme" -Value "These Package Information Registry Keys (PIRKs) are used for SCCM application detection. Please do not modify unless you know what you are doing." -Type String
			Set-RegistryKey -Key "HKLM:\Software\Tuolumne County\Package Information\$apName" -Name "apVersion" -Value "$apversion" -Type String
			Set-RegistryKey -Key "HKLM:\Software\Tuolumne County\Package Information\$apName" -Name "apScriptVersion" -Value "$apScriptVErsion" -Type String
			Write-Log -Message "Creating package information registry key (PIRK) complete." 
		}
		# If creation failed, log results. Exit script.
		Catch [System.Exception] {
			Write-Log -Message "Creating package information registry key (PIRK) failed with error: $_."
			Write-Log -Message "Exiting script with error."
			Exit-Script -ExitCode 1627
		}

		##*===============================================
		##* ANCHOR: POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'

		# For each item in the array...
		# Creates a new Shortcut in every location specified in the array.
		# If shortcut creation failed, log the results.
		# If shortcut creation was successful, log the results.
		Foreach ($appShortcutDir in $appShortcutDirs) {
			Try {
				Write-Log -Message "Creating shortcuts in $appShortcutDir..."
				New-Shortcut -Path "$appShortcutDir\$32bitAppInstallName.lnk" -TargetPath "https://www.tuolumnecounty.ca.gov/1481/Clerk-of-the-Board-Records" -IconLocation "$envProgramFiles\Tuolumne County\Clerk of the Board Records Page\Assets\DigitalReel_ico2.ico" -WindowStyle "Maximized" -Description "Clerk of the Board Records Page"
			}
			# If shortcut creation failed, log results. Exit script.
			Catch [System.Exception] {
				Write-Log -Message "Creating shortcuts in $appShortcutDir failed with error: $_."
				Write-Log -Message "Exiting script with error."
				Exit-Script -ExitCode 1627
			}
			Write-Log -Message "Creating shortcuts in $appShortcutDir complete."
		}
		
	
		# Update (refresh) the desktop.
		Update-Desktop

	}
	ElseIf ($deploymentType -ieq 'Uninstall') {
		##*===============================================
		##* ANCHOR: PRE-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Uninstallation'

		##*===============================================
		##* ANCHOR: UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Uninstallation'

		## Handle Zero-Config MSI Uninstallations
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Uninstall'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
			Execute-MSI @ExecuteDefaultMSISplat
		}

		# For each item in the array...
		# Check if previous versions of application install directory exist.
		# If application install directory exists, remove previous versions of application install directory.
		# If removal failed, log results. Exit script.
		# Else, log results from check.
		Foreach ($removeShortcutDir in $removeShortcutDirs) {
			# Check if previous versions of application install directory exist.
			If (Test-Path -Path "$removeShortcutDir\$32bitAppUninstallName*") {
				Write-Log -Message "Previous versions of $32bitAppUninstallName exist in $removeShortcutDir. Removing..."  
				# Remove previous versions of application install directory.
				Try {
					Remove-Item -Path "$removeShortcutDir\$32bitAppUninstallName*" -Force -Recurse
					Remove-Item -Path "${Env:PUBLIC}\Desktop\$32bitAppUninstallName*" -Force
				}
				# If removal failed, log results. Exit script.
				Catch [System.Exception] {
					Write-Log -Message "Removing previous versions of $32bitAppUninstallName in $removeShortcutDir failed with error: $_."
					Write-Log -Message "Exiting script with error."
					Exit-Script -ExitCode 1627
				}
				Write-Log -Message "Removing previous versions of $32bitAppUninstallName in $removeShortcutDir complete."
			}
			# Else, log results from check. 
			Else {  
				Write-Log -Message "Previous versions of $32bitAppUninstallName in $removeShortcutDir do not exist."        
			}
		}

		# Update (refresh) the desktop.
		Update-Desktop

		# Check if previous versions of package information registry key (PIRK) exist. 
		# If package information registry key (PIRK) exists, remove previous versions of package information registry key (PIRK).
		# If removal failed, log results. Exit script. 
		# Else, log results from check.
		If (Test-Path -Path "HKLM:\SOFTWARE\Tuolumne County\Package Information\$apName*") { 
			# Remove previous versions of package information registry key (PIRK).
			Try {
				Write-Log -Message "Previous versions of package information registry key (PIRK) exist. Removing..."
				Remove-Item -Path "HKLM:\SOFTWARE\Tuolumne County\Package Information\$apName*" -Force
				Write-Log -Message "Removing previous versions of package information registry key (PIRK) complete."
			}
			# If removal failed, log results. Exit script. 
			Catch [System.Exception] {
				Write-Log -Message "Removing previous versions of package information registry key (PIRK) failed with error: $_"
				Write-Log -Message "Exiting script with error."
				Exit-Script -ExitCode 1627
			}
		}
		# Else, log results from check. 
		Else { 
			Write-Log -Message "Previous versions of package information registry key (PIRK) do not exist."  
		}

		##*===============================================
		##* ANCHOR: POST-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Uninstallation'



	}
	ElseIf ($deploymentType -ieq 'Repair') {
		##*===============================================
		##* ANCHOR: PRE-REPAIR
		##*===============================================
		[string]$installPhase = 'Pre-Repair'

		## <Perform Pre-Repair tasks here>

		##*===============================================
		##* ANCHOR: REPAIR
		##*===============================================
		[string]$installPhase = 'Repair'

		## Handle Zero-Config MSI Repairs
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Repair'; Path = $defaultMsiFile; }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
			Execute-MSI @ExecuteDefaultMSISplat
		}
		
		# <Perform Repair tasks here>

		##*===============================================
		##* ANCHOR: POST-REPAIR
		##*===============================================
		[string]$installPhase = 'Post-Repair'

		## <Perform Post-Repair tasks here>

	}
	##*===============================================
	##* END SCRIPT BODY
	##*===============================================

	## Call the Exit-Script function to perform final cleanup operations
	Exit-Script -ExitCode $mainExitCode
}
Catch {
	[int32]$mainExitCode = 60001
	[string]$mainErrorMessage = "$(Resolve-Error)"
	Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
	Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
	Exit-Script -ExitCode $mainExitCode
}
