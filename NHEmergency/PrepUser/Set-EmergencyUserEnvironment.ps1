#################################################################################################
# Name			: 	Set-EmergencyUserEnvironment.ps1
# Description	: 	Main
# Author		: 	Axel Anderson
# License		:	
# Annotations	:	Parts taken from project : Manage-AD	by Axel Anderson
#					Parts taken from project : PSST by Axel Anderson
#					Parts taken from Project : Set-OfficeLinks by Axel Anderson
#					Parts taken from Project : Update-TrainingsEnvironment by Axel Anderson
# Date			: 	02.11.2016 created
#
# History		:	0.1.0.0 02.11.2016	Create bei Axel Anderson
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
#Requires –Version 2
[CmdletBinding()]
Param   (
		)
Set-StrictMode -Version Latest	
#
#################################################################################################
#
#
#region ScriptVariables

$HTFolder=@{};
[Environment+SpecialFolder]::GetNames([Environment+SpecialFolder]) | % {
		$Path = [Environment]::GetFolderPath($_); 
		$HTFolder[$_] = $Path;
	}
	
$script:userDesktopFolder	= $HTFolder["Desktop"]
$script:userDocumentsFolder	= $HTFolder["Personal"]	

$script:tmpPath = ($env:TEMP)
	
#endregion ScriptVariables
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function Detect-OfficeVersion {
[CmdletBinding()]
Param	()

	# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Function Get-Officeversion {
		[CmdletBinding()]
		Param	()
		
		$ProgramFiles		= ${env:ProgramFiles}
		$ProgramFilesX86	= ${env:ProgramFiles(x86)}
		
		$AppName = "winword.exe"
		$version = 0
		if (Test-Path HKLM:\software\Microsoft\Office) {
			dir HKLM:\software\Microsoft\Office |%{
				if ($_.name -match '(\d+)\.') {
					if ([int]$matches[1] -gt $version) {
						$TestVersion = [int]$matches[1]
						$OfficePath = "Microsoft Office\Office"+$TestVersion	
						$OSPPPath    = (Join-Path (Join-Path $ProgramFiles    $OfficePath) $AppName)
						$OSPPPathX86 = (Join-Path (Join-Path $ProgramFilesX86 $OfficePath) $AppName)
		
						if ((Test-Path $OSPPPath) -or (Test-Path $OSPPPathX86)) {
							$version = [int]$matches[1]
						} else {
							$OfficePath = "Microsoft Office\root\Office"+$TestVersion	
							$OSPPPath    = (Join-Path (Join-Path $ProgramFiles    $OfficePath) $AppName)
							$OSPPPathX86 = (Join-Path (Join-Path $ProgramFilesX86 $OfficePath) $AppName)
							
							if ((Test-Path $OSPPPath) -or (Test-Path $OSPPPathX86)) {
								$version = [int]$matches[1]
							}						
						}
					}
				} 
			}
		} 
		Write-Output $Version
	}
	# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	$Username = $Env:Username
	$OfficeShortName = $Username.Substring(0,3)
	
	Switch ($OfficeShortName) {
		"O07"	{
					$OfficeVersion = 12
					break
				}
		"O10"	{
					$OfficeVersion = 14
					break
				}
		"O13"	{
					$OfficeVersion = 15
					break
				}
		"O16"	{
					$OfficeVersion = 16
					break
				}
		default {
					$OfficeVersion = Get-Officeversion
				}
	}
	Write-Output $OfficeVersion
}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function Detect-Application {
[CmdletBinding()]
Param	()

	$Username = $Env:Username
	$AppShortName = $Username.Substring(3,2)
	
	Switch ($AppShortName) {
		"EX"	{
					$AppName = "Excel"
					break
				}
		"WO"	{
					$AppName = "Word"	
					break
				}
		"PO"	{
					$AppName = "Powerpoint"
					break
				}
		"AC"	{
					$AppName = "Access"				
					break
				}
		"OU"	{
					$AppName = "Outlook"					
					break
				}
		"W7"	{
					$AppName = "Windows 7"					
					break					
				}
		default	{
					$AppName = ""			
					break
				}
	}
	Write-Output $AppName	
}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function Get-OfficeVersionString {
[CmdletBinding()]
Param	(
			[int]$OfficeVersion
		)

	switch ($Officeversion) {
		12		{
					$OfficeVersionString = "2007"
					break
				}
		14		{
					$OfficeVersionString = "2010"
					break
				}
		15		{
					$OfficeVersionString = "2013"
					break
				}
		16		{
					$OfficeVersionString = "2016"
					break
				}
		default {
					$OfficeVersionString = ""
					break
				}
	}
	Write-Output $OfficeVersionString
}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function Copy-TrainingFiles {
[CmdletBinding()]
Param	(
			[string]$TrainingFilesPath,
			[int]$OfficeVersion,
			[string]$Appname
		)

	if (($TrainingFilesPath -ne "") -and ($OfficeVersion -ne 0) -and ($Appname -ne "")) {
		$OfficeVersionString = Get-OfficeVersionString $OfficeVersion
	
		if ($Appname -ieq "Windows 7") {
			$SourcePath = Join-Path (Join-Path $TrainingFilesPath ("Office "+$OfficeVersionString)) ($Appname)
		} else {
			$SourcePath = Join-Path (Join-Path $TrainingFilesPath ("Office "+$OfficeVersionString)) ($Appname+" "+$OfficeVersionString)
		}
		Write-verbose "Copy-TrainingFiles : SourcePath : $($SourcePath)"

		$DestPath = $script:userDocumentsFolder
		Write-verbose "Copy-TrainingFiles : DestPath : $($DestPath)"
		try {
			Copy-Item -Path $SourcePath -Destination $DestPath -Recurse -Force -Confirm:$False
		} catch {
			Write-Error "Copy '$($SourcePath)' to '$($DestPath)' FAILED"
		}
	}
}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function Get-ApplicationPath {
    Param ([string]$extension)

    Push-Location
	$path = ""

	$o = New-Object PSObject -Property @{
			"Extension" = $extension
			"AppPath" = $path
		}
    
	try {
		Set-Location "HKLM:\Software\Classes\$($extension)" -ErrorAction Stop
		$default = (Get-ItemProperty -Path $pwd.Path -Name '(Default)').'(Default)'
		
		try {
			Set-Location "HKLM:\Software\Classes\$($default)\shell\open\command" -errorAction Stop
			(Get-ItemProperty -Path $pwd.Path -Name '(Default)').'(Default)' -match '([^"^\s]+)\s*|"([^"]+)"\s*' | Out-Null
			$path = $matches[0].ToString()

			# NOTE: DAS ist wichtig hier, sonst wird IN den String ein Anführungszeichen geschrieben
			$path = $path -Replace '"', ''
			$o.AppPath = $path

		}
		catch {
			#Write-Error $_
		}
	}
	catch {
		#Write-Error $_
	}  
    Pop-Location
	$o
}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function New-AppLinkFile {
	param  ([string]$lnkFileName, 
			[string]$targetFile
			)
			
	$wshShellObject = New-Object -com WScript.Shell
	
	$wshShellLink = $wshShellObject.CreateShortcut($lnkFileName)
	
	$wshShellLink.TargetPath = $targetFile
	$wshShellLink.WindowStyle = 1
	$wshShellLink.IconLocation = $targetFile
	#
	# WorkingDir per Default IMMER Eigene Dokumente
	#
	$wshShellLink.WorkingDirectory = $script:userDocumentsFolder
	$wshShellLink.Save()
}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function Create-ApplikationLink {
	param(	
			[string]$VersionOffice,
			[string]$Appname,
			[string]$AppExtension1,
			[string]$AppExtension2
		  )
	
	$tmpLnkFilePath = ""
		  
	$Info = Get-ApplicationPath $AppExtension1
	#$Info | out-host
	
	if ($Info.AppPath -eq "") {
		$Info = Get-ApplicationPath $AppExtension2
	}
	#$Info | out-host
	if ($Info.AppPath -ne "") {
		$tmpLnkFilename = $Appname + " " + $VersionOffice + ".lnk"
		$tmpLnkFilePath = Join-Path $script:tmpPath $tmpLnkFilename
		
		New-AppLinkFile $tmpLnkFilePath $($Info.AppPath)
	}

	$tmpLnkFilePath
}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function Do-PinToDesktop {
	param(	$AppLnkname)

	if ( ($appLnkName -ne "") -and (Test-Path $AppLnkname) -and (-NOT( Test-Path $(Join-Path $script:userDesktopFolder $(Split-path $AppLnkName -Leaf))))) {
		Copy-Item $AppLnkname $script:userDesktopFolder
	}
}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function  Set-ExcelAppLinkToDesktop {
[CmdletBinding()]
Param	(
			[int]$OfficeVersion
		)
	
	
	$OfficeVersionString = Get-OfficeVersionString $OfficeVersion
	
	$AppExtension1 	= ".xls"
	$AppExtension2 	= ".xlsx"
	$Appname		= "Excel"

	$AppLinkName = Create-ApplikationLink $OfficeVersionString $Appname $AppExtension1 $AppExtension2

	Do-PinToDesktop 	$AppLinkName

}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function  Set-WordAppLinkToDesktop {
[CmdletBinding()]
Param	(
			[int]$OfficeVersion
		)
	
	
	$OfficeVersionString = Get-OfficeVersionString $OfficeVersion
	
	$AppExtension1 	= ".doc"
	$AppExtension2 	= ".docx"
	$Appname		= "Word"

	$AppLinkName = Create-ApplikationLink $OfficeVersionString $Appname $AppExtension1 $AppExtension2

	Do-PinToDesktop 	$AppLinkName

}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function  Set-PowerpointAppLinkToDesktop {
[CmdletBinding()]
Param	(
			[int]$OfficeVersion
		)
	
	
	$OfficeVersionString = Get-OfficeVersionString $OfficeVersion
	
	$AppExtension1 	= ".ppt"
	$AppExtension2 	= ".pptx"
	$Appname		= "PowerPoint"

	$AppLinkName = Create-ApplikationLink $OfficeVersionString $Appname $AppExtension1 $AppExtension2

	Do-PinToDesktop 	$AppLinkName

}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function  Set-AccessAppLinkToDesktop {
[CmdletBinding()]
Param	(
			[int]$OfficeVersion
		)
	
	
	$OfficeVersionString = Get-OfficeVersionString $OfficeVersion
	
	$AppExtension1 	= ".mdb"
	$AppExtension2 	= ".accdb"
	$Appname		= "Access"

	$AppLinkName = Create-ApplikationLink $OfficeVersionString $Appname $AppExtension1 $AppExtension2

	Do-PinToDesktop 	$AppLinkName

}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function  Set-OutlookAppLinkToDesktop {
[CmdletBinding()]
Param	(
			[int]$OfficeVersion
		)
	
	
	$OfficeVersionString = Get-OfficeVersionString $OfficeVersion
	
	$AppExtension1 	= ".pst"
	$AppExtension2 	= ".ost"
	$Appname		= "Outlook"

	$AppLinkName = Create-ApplikationLink $OfficeVersionString $Appname $AppExtension1 $AppExtension2

	Do-PinToDesktop 	$AppLinkName

}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function Set-ApplicationDesktopPin {
[CmdletBinding()]
Param	(
			[int]$OfficeVersion
		)
		
	$Username = $Env:Username
	$AppShortName = $Username.Substring(3,2)
	
	Switch ($AppShortName) {
		"EX"	{
					Set-ExcelAppLinkToDesktop -OfficeVersion $OfficeVersion
					break
				}
		"WO"	{
					Set-WordAppLinkToDesktop -OfficeVersion $OfficeVersion
					break
				}
		"PO"	{
					Set-PowerPointAppLinkToDesktop -OfficeVersion $OfficeVersion
					break
				}
		"AC"	{
					Set-AccessAppLinkToDesktop -OfficeVersion $OfficeVersion
					break
				}
		"OU"	{
					Set-OutlookAppLinkToDesktop -OfficeVersion $OfficeVersion
					break
				}
		"W7"	{
					break
				}
		default	{
					break
				}
	}		

}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function Set-TrainingFilesDesktopPin {
[CmdletBinding()]
Param	(
			[int]$OfficeVersion,
			[string]$Appname
		)
		
	if (($OfficeVersion -ne 0) -and ($Appname -ne "")) {
		$OfficeVersionString = Get-OfficeVersionString $OfficeVersion
		
		if ($Appname -ieq "Windows 7") {
			$Foldername = Join-Path $script:userDocumentsFolder ($Appname)
			$tmpLnkFilename = ("Übungsdateien "+($Appname) + ".lnk"
		} else {
			$Foldername = Join-Path $script:userDocumentsFolder ($Appname+" "+$OfficeVersionString)
			$tmpLnkFilename = ("Übungsdateien "+($Appname+" "+$OfficeVersionString)) + ".lnk"
		}

		$tmpLnkFilePath = Join-Path $script:tmpPath $tmpLnkFilename
		Write-Verbose "Set Folder Pin : $($tmpLnkFilePath) : $($Foldername)"
		New-AppLinkFile $tmpLnkFilePath $Foldername
		Do-PinToDesktop $tmpLnkFilePath
	}
}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
#region MAIN
# #############################################################################
# ##### MAIN
# #############################################################################
#
# ### STEP 1 - Detect Office Version ####
#
$OfficeVersion = Detect-OfficeVersion
Write-Verbose "Detect Office Version $($OfficeVersion)"
#
# ### STEP 2 - Detect TrainingFiles-Path    ####
#
if (Test-Path "D:\Office_Übungsdateien") {
	$TrainingFilesPath = "D:\Office_Übungsdateien"
} elseif (Test-Path "D:\Office_Uebungsdateien") {
	$TrainingFilesPath = "D:\Office_Uebungsdateien"
} elseif (Test-Path "\\wihan.nh\freigabe\Office_Übungsdateien") {
	$TrainingFilesPath = "\\wihan.nh\freigabe\Office_Übungsdateien"
} elseif (Test-Path "\\wihan.nh\freigabe\Office_Uebungsdateien") {
	$TrainingFilesPath = "\\wihan.nh\freigabe\Office_Uebungsdateien"
} else {
	$TrainingFilesPath = ""
}
Write-Verbose "Detect TrainingFiles-Path $($TrainingFilesPath)"
#
# ### STEP 3 - Detect Application    ####
#
$ApplicationName = Detect-Application
Write-Verbose "Detect Application Name $($ApplicationName)"
#
# ### STEP 4 - Clean Desktop    ####
#
Write-Verbose "Clean Desktop : $($script:userDesktopFolder)"
Get-ChildItem $script:userDesktopFolder | Remove-Item -Recurse -Force -Confirm:$False
#
# ### STEP 4 - Clean Documents    ####
#
Write-Verbose "Clean Desktop : $($script:userDocumentsFolder)"
Get-ChildItem $script:userDocumentsFolder | Remove-Item -Recurse -Force -Confirm:$False
#
# ### STEP 5 - Copy Training Files    ####
#
Copy-TrainingFiles -TrainingFilesPath $TrainingFilesPath -OfficeVersion $OfficeVersion -Appname $ApplicationName
#
# ### STEP 6 - Pin AppLink to Desktop    ####
#
Set-ApplicationDesktopPin -OfficeVersion $OfficeVersion
#
# ### STEP 6 - Pin TrainingFile-Path to Desktop    ####
#
Set-TrainingFilesDesktopPin -OfficeVersion $OfficeVersion -Appname $ApplicationName
#
# #############################################################################
# ##### END MAIN
# #############################################################################
#endregion MAIN


