<#
.SYNOPSIS
	Installatie VisualC++ Redistributable
	
	FileName:    VisualC++Redistributable-installer.ps1
    Author:      Mark Messink
    Contact:     
    Created:     2022-06-14
    Updated:     2022-12-20

    Version history:
    1.0.0 - (2022-06-14) Initial Script
	1.0.1 - (2022-12-20) Script verniewd
	1.1.0 - 

.DESCRIPTION
	Download en installatie VisualC++ Redistributable

.PARAMETER
	<beschrijf de parameters die eventueel aan het script gekoppeld moeten worden>

.INPUTS


.OUTPUTS
	logfiles:
	PSlog_<naam>	Log gegenereerd door een powershell script
	INlog_<naam>	Log gegenereerd door Intune (Win32)
	AIlog_<naam>	Log gegenereerd door de installer van een applicatie bij de installatie van een applicatie
	ADlog_<naam>	Log gegenereerd door de installer van een applicatie bij de de-installatie van een applicatie
	Een datum en tijd wordt automatisch toegevoegd

.EXAMPLE
	./scriptnaam.ps1

.LINK Information
	https://docs.microsoft.com/en-US/cpp/windows/latest-supported-vc-redist?view=msvc-170

.LINK DownloadLocations
    https://aka.ms/vs/17/release/vc_redist.x86.exe
    https://aka.ms/vs/17/release/vc_redist.x64.exe

.NOTES
	WindowsBuild:
	Het script wordt uitgevoerd tussen de builds LowestWindowsBuild en HighestWindowsBuild
	LowestWindowsBuild = 0 en HighestWindowsBuild 50000 zijn alle Windows 10/11 versies
	LowestWindowsBuild = 19000 en HighestWindowsBuild 19999 zijn alle Windows 10 versies
	LowestWindowsBuild = 22000 en HighestWindowsBuild 22999 zijn alle Windows 11 versies
	Zie: https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information

#>

#################### Variabelen #####################################
$logpath = "C:\IntuneLogs"
$NameLogfile = "inlog_VisualC++Redistributable.txt"
$LowestWindowsBuild = 0
$HighestWindowsBuild = 50000



#################### Einde Variabelen ###############################


#################### Start base script ##############################
### Niet aanpassen!!!

# Prevent terminating script on error.
$ErrorActionPreference = 'Continue'

# Create logpath (if not exist)
If(!(test-path $logpath))
{
      New-Item -ItemType Directory -Force -Path $logpath
}

# Add date + time to Logfile
$TimeStamp = "{0:yyyyMMdd}" -f (get-date)
$logFile = "$logpath\" + "$TimeStamp" + "_" + "$NameLogfile"

# Start Transcript logging
Start-Transcript $logFile -Append -Force

# Start script timer
$scripttimer = [system.diagnostics.stopwatch]::StartNew()

# Controle Windows Build
$WindowsBuild = [System.Environment]::OSVersion.Version.Build
Write-Output "------------------------------------"
Write-Output "Windows Build: $WindowsBuild"
Write-Output "------------------------------------"
If ($WindowsBuild -ge $LowestWindowsBuild -And $WindowsBuild -le $HighestWindowsBuild)
{
#################### Start base script ################################

#################### Start uitvoeren script code ####################
Write-Output "#####################################################################################"
Write-Output "### Start uitvoeren script code                                                   ###"
Write-Output "#####################################################################################"

Write-Output "-------------------------------------------------------------------"
Write-Output "Download and installation vc_redist.x86.exe from Microsoft"
Write-Output "-------------------------------------------------------------------"
Write-Output "Check registry key:"
$REGx86 = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\DevDiv\vc\Servicing\14.0\RuntimeMinimum'
$REGx86
	
	if (-not(Test-Path -Path "$REGx86")) {
		Write-Output "-------------------------------------------------------------------"
		Write-Output "Download vc_redist.x86.exe from Microsoft"
		$downloadLocation = "https://aka.ms/vs/17/release/vc_redist.x86.exe"
		$downloadDestination = "$($env:TEMP)\vc_redist.x86.exe"
		$webClient = New-Object System.Net.WebClient
		$webClient.DownloadFile($downloadLocation, $downloadDestination)
		Write-Output "-------------------------------------------------------------------"
		Write-Output "Version vc_redist.x86.exe Download"
		(Get-Item $downloadDestination).VersionInfo | FL Productname, FileName, Productversion
		Write-Output "-------------------------------------------------------------------"
		Write-Output "Install vc_redist.x86.exe"
		$installProcess = Start-Process $downloadDestination -ArgumentList "/install /quiet /norestart /log c:/IntuneLogs/ialog_vc_redist.x86.txt" -NoNewWindow -PassThru
		$installProcess.WaitForExit()
		} else {
		Write-Output "-------------------------------------------------------------------"
		Write-Output "vc_redist.x86 already exists. Installation skipped"
	}
		
	if (Test-Path -Path "$REGx86") {
	Write-Output "-------------------------------------------------------------------"
	Write-Output "vc_redist.x86 information:"
	$X86key =  Get-ItemProperty -Path $REGx86 -Name Version
	$X86key.Version
	Write-Output "-------------------------------------------------------------------"
	} else {
		Write-Output "-------------------------------------------------------------------"
		Write-Output "Registry key not found. Installation failed"
		Write-Output "-------------------------------------------------------------------"
	}
	
Write-Output "-------------------------------------------------------------------"
Write-Output "Download and installation vc_redist.x64.exe from Microsoft"
Write-Output "-------------------------------------------------------------------"
Write-Output "Check registry key:"
	$REGx64 = 'HKLM:\SOFTWARE\Microsoft\DevDiv\vc\Servicing\14.0\RuntimeMinimum'
	$REGx64
	
	if (-not(Test-Path -Path "$REGx64")) {
		Write-Output "-------------------------------------------------------------------"
		Write-Output "Download vc_redist.x64.exe from Microsoft"
		$downloadLocation = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
		$downloadDestination = "$($env:TEMP)\vc_redist.x64.exe"
		$webClient = New-Object System.Net.WebClient
		$webClient.DownloadFile($downloadLocation, $downloadDestination)
		Write-Output "-------------------------------------------------------------------"
		Write-Output "Version vc_redist.x64.exe Download"
		(Get-Item $downloadDestination).VersionInfo | FL Productname, FileName, Productversion
		Write-Output "-------------------------------------------------------------------"
		Write-Output "Install vc_redist.x64.exe"
		$installProcess = Start-Process $downloadDestination -ArgumentList "/install /quiet /norestart /log c:/IntuneLogs/ialog_vc_redist.x64.txt" -NoNewWindow -PassThru
		$installProcess.WaitForExit()
		} else {
		Write-Output "-------------------------------------------------------------------"
		Write-Output "vc_redist.x64 already exists. Installation skipped"
	}
		
	if (Test-Path -Path "$REGx64") {
	Write-Output "-------------------------------------------------------------------"
	Write-Output "vc_redist.x64 information:"
	$X64key =  Get-ItemProperty -Path $REGx64 -Name Version
	$X64key.Version
	Write-Output "-------------------------------------------------------------------"
	} else {
		Write-Output "-------------------------------------------------------------------"
		Write-Output "Registry key not found. Installation failed"
		Write-Output "-------------------------------------------------------------------"
	}

Write-Output "#####################################################################################"
Write-Output "### Einde uitvoeren script code                                                   ###"
Write-Output "#####################################################################################"
#################### Einde uitvoeren script code ####################

#################### End base script #######################

# Controle Windows Build
}Else {
Write-Output "-------------------------------------------------------------------------------------"
Write-Output "### Windows Build versie voldoet niet, de script code is niet uitgevoerd. ###"
Write-Output "-------------------------------------------------------------------------------------"
}

#Stop and display script timer
$scripttimer.Stop()
Write-Output "------------------------------------"
Write-Output "Script elapsed time in seconds:"
$scripttimer.elapsed.totalseconds
Write-Output "------------------------------------"

#Stop Logging
Stop-Transcript
#################### End base script ################################

#################### Einde Script ###################################
