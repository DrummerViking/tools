﻿﻿<#
	.SYNOPSIS
		Installs the available tools from github
		
	.DESCRIPTION
		This script installs the different available tools from github.
		
		It does so by ...
		- downloading the master branch as zip to $env:TEMP
		- Unpacking that zip file to a folder in $env:TEMP
		- Moving the tool content to the user's Desktop

	.PARAMETER Tool
		Specifies the specific tool to be downloaded. The parameter can be auto completed with 'TAB' key and will display the available tools.

    .PARAMETER Force
		The install script will overwrite an existing module.
#>
[CmdletBinding()]
Param (
	[ValidateSet('Collect-DAuthtroubleshootingLogs', 'Collect-OAuthtroubleshootingLogs','DeleteMeetings-Gui','Get-ExchangeServerInfo','Get-MRMRoamingXMLData','Get-MRMStatistics','Manage-FolderPermissionsGUi','Manage-Mobiles-GUI','Manage-UserPhotoGui')]
	[string]
	$Tool = "Collect-DAuthtroubleshootingLogs"
)

#region Configuration for cloning script
# Brach selected
$Branch = "master"

# Name of the module that is being cloned
$ModuleName = "agallegoCSS-tools"

# Base path to the github repository
$BaseUrl = "https://github.com/agallego-css/tools"
#endregion Configuration for cloning script

function Write-LocalMessage
{
    [CmdletBinding()]
    Param (
        [string]$Message
    )

    if (Test-Path function:Write-PSFMessage) { Write-PSFMessage -Level Important -Message $Message }
    else { Write-Host $Message }
}

try
{
	[System.Net.ServicePointManager]::SecurityProtocol = "Tls12"

	Write-LocalMessage -Message "Downloading repository from '$($BaseUrl)/archive/$($Branch).zip'"
	Invoke-WebRequest -Uri "$($BaseUrl)/archive/$($Branch).zip" -UseBasicParsing -OutFile "$($env:TEMP)\$($ModuleName).zip" -ErrorAction Stop
	
	Write-LocalMessage -Message "Creating temporary project folder: '$($env:TEMP)\$($ModuleName)'"
	$null = New-Item -Path $env:TEMP -Name $ModuleName -ItemType Directory -Force -ErrorAction Stop
	
	Write-LocalMessage -Message "Extracting archive to '$($env:TEMP)\$($ModuleName)'"
	Expand-Archive -Path "$($env:TEMP)\$($ModuleName).zip" -DestinationPath "$($env:TEMP)\$($ModuleName)" -ErrorAction Stop
	
	$basePath = Get-ChildItem "$($env:TEMP)\$($ModuleName)\*" | Select-Object -First 1
	
	# Determine output path
	$path = "$home\Desktop"
    
    Write-LocalMessage -Message "Copying files to $($path)"
	$file = Get-ChildItem -Include "$tool.ps1" -Path $basePath -Recurse
	Move-Item -Path $file.FullName -Destination $path -ErrorAction Stop
	
	Write-LocalMessage -Message "Cleaning up temporary files"
	Remove-Item -Path "$($env:TEMP)\$($ModuleName)" -Force -Recurse
	Remove-Item -Path "$($env:TEMP)\$($ModuleName).zip" -Force
	
	Write-LocalMessage -Message "Installation of the tool $($tool) completed successfully!"
}
catch
{
	Write-LocalMessage -Message "Installation of the tool $($tool) failed!"
	
	Write-LocalMessage -Message "Cleaning up temporary files"
	Remove-Item -Path "$($env:TEMP)\$($ModuleName)" -Force -Recurse
	Remove-Item -Path "$($env:TEMP)\$($ModuleName).zip" -Force
	
	throw
}