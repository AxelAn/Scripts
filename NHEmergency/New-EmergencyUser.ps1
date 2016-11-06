#################################################################################################
# Name			: 	New-EmergencyUser.ps1
# Description	: 	Main
# Author		: 	Axel Anderson
# License		:	
# Date			: 	01.11.2016 created
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
#Requires –Version 2
Param   (
		)
Set-StrictMode -Version Latest	
#
#################################################################################################
#
#
#region ScriptVariables
$global:RTCred = [System.Management.Automation.PSCredential]::Empty

$script:exchangeServer = ""

#endregion ScriptVariables
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function Test-IsAdmin {
[CmdletBinding()]
Param	()

	$user = [Security.Principal.WindowsIdentity]::GetCurrent();
	$IsAdmin = (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator) 
	
	Write-output $IsAdmin
}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function Get-ExchangeServerInSite-AAN {
	#
	# http://www.mikepfeiffer.net/2010/04/find-exchange-servers-in-the-local-active-directory-site-using-powershell/
	#
	$role = @{
		2  = "MB"
		4  = "CAS"
		16 = "UM"
		32 = "HT"
		64 = "ET"
	}
    $ADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]
    $siteDN = $ADSite::GetComputerSite().GetDirectoryEntry().distinguishedName
    $configNC=([ADSI]"LDAP://RootDse").configurationNamingContext
    $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC")
    $objectClass = "objectClass=msExchExchangeServer"
    $version = "versionNumber>=1937801568"
    $site = "msExchServerSite=$siteDN"
    $search.Filter = "(&($objectClass)($version)($site))"
    $search.PageSize=1000
    [void] $search.PropertiesToLoad.Add("name")
    [void] $search.PropertiesToLoad.Add("msexchcurrentserverroles")
    [void] $search.PropertiesToLoad.Add("networkaddress")
    $search.FindAll() | %{
        $o = New-Object PSObject -Property @{
            Name = $_.Properties.name[0]
            FQDN = $_.Properties.networkaddress |
                %{if ($_ -match "ncacn_ip_tcp") {$_.split(":")[1]}}
            Roles = ""
        }
		$iRoles = $_.Properties.msexchcurrentserverroles[0]
		$o.Roles = ($role.keys | ?{$_ -band $iroles} | %{$role.Get_Item($_)}) -join ", "
		$o
    }
}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function New-AdUser-NH {
[cmdletbinding()]
	Param	(
				[Parameter(Mandatory=$true)][string]$SamAccountName,
				[Parameter(Mandatory=$true)][string]$GivenName,
				[Parameter(Mandatory=$true)][string]$SurName,
				[Parameter(Mandatory=$true)][string]$Password,
				[Parameter(Mandatory=$true)][string]$OUPathDN,
				[Parameter(Mandatory=$true)][string]$GroupDN
			)

	$displayName		=	$givenName + " " + $surName
	$secpasswd 			=	ConvertTo-SecureString $Password -AsPlainText -Force
	
	try {
		$existUser = Get-ADUser -Identity $samAccountName -EA Stop
	} catch {
		$existUser = $null
	}

	if ($existUser -eq $null) {
	
		try {
			New-ADUser `
				-Name						$displayName	`
				-SamAccountName 			$samAccountName	`
				-Type 						"User"	`
				-CannotChangePassword		1	`
				-PasswordNeverExpires 		1	`
				-Enabled					1	`
				-GivenName					$givenName	`
				-Surname					$surname	`
				-DisplayName 				$displayName	`
				-UserPrincipalName 			($samAccountName + "@" + (Get-ADDomain).DNSRoot)		`
				-ProfilePath				""		`
				-AccountPassword 			$secpasswd	`
				-Path						$OUPathDN	`
				-Credential 				$global:RTCred `
				-EA							Stop
				
			Write-Host -f darkgreen "User: $($displayName) angelegt."
		} catch {
			$str = "User: $($displayName) konnte nicht angelegt werden.$($_)"
			Write-Host -f red $str
		}
		try {
			#Clear meint: 'Never Expire'	
			Clear-ADAccountExpiration -Identity $samAccountName -Credential $global:RTCred -EA Stop
			Write-Host -f darkgreen "User: $($displayName) clear Expiration."
		} catch {
			$str = "User: $($displayName) clear Expiration ERROR."
			Write-Host -f red $str
		}
		try {
			Add-ADGroupMember -Identity $GroupDN -Members $samAccountName -Credential $global:RTCred -EA Stop
			Write-Host -f darkgreen "User: $($displayName) add to group $($GroupDN)."
			
		} catch {
			$str = "User: $($displayName) add to group $($GroupDN) ERROR."
			Write-Host -f red $str
		}
		
		
	} else {
		$str = "User: $($displayName) existiert schon."
		Write-Host -f yellow  $str
	}
}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function Enable-Mailbox-NH {
[cmdletbinding()]
	Param	(
				[Parameter(Mandatory=$true)][string]$SamAccountName,
				[Parameter(Mandatory=$true)][string]$Alias,
				[Parameter(Mandatory=$true)][string]$Databasename,
				[switch]$SendAs=$false,
				[string]$SendAsUser="",
				[switch]$SendOnBeHalf=$false,
				[string]$SendOnBeHalfUser=""
			)
	
	$exchangeSession = $null
	Write-Host "~~~~~~ Enable-Mailbox-NH ~~~~~~~~~"
	try	{
		$exchangeSession = New-PSSession `
				-ConfigurationName Microsoft.Exchange `
				-ConnectionUri ("http://"+($script:exchangeServer).FQDN+"/PowerShell/") `
				-Authentication Kerberos `
				-Credential $global:RTCred `
				-ErrorAction Stop
		
		Write-Host -F DarkGreen "Connection successfull" 

		Import-PSSession -Session $exchangeSession -CommandName Enable-Mailbox, Set-Mailbox, Get-Mailbox, Add-ADPermission -AllowClobber | out-null
	} catch {
		Write-Host -f red "Connection not successfull: $($_.Exception.Message)"
		$exchangeSession = $null
	}
	
	if ($exchangeSession) {

		try {
			$ExistMailbox = Get-Mailbox -Identity $SamAccountName -EA Stop
			Write-Host -F Red "Mailbox Existiert"
		} catch {
			$ExistMailbox = $null
			Write-Host -F Green "Mailbox Existiert NICHT"
		}
		
		if ($ExistMailbox -eq $null) {	
			try {

				try {
					$newmailBox = Enable-Mailbox `
									-Identity $SamAccountName	`
									-Alias $Alias	`
									-Database $Databasename `
									-EA Stop
				} catch  {
					try {
						Start-Sleep -MilliSeconds 10000
						$newmailBox = Enable-Mailbox `
										-Identity $SamAccountName	`
										-Alias $Alias	`
										-Database $Databasename `
										-EA Stop
					
					} catch {
						$NewMailbox = $Null
					}		
				}

				#
				# Senden ALS
				#
				if ($NewMailbox) {
					if ($SendAs) {
						Add-ADPermission 	`
								-Identity "$($SendAsUser)"	`
								-User ((Get-ADDomain).DNSRoot+"\"+$SamAccountName)	`
								-AccessRights ExtendedRight 	`
								-ExtendedRights 'Send-as' `
								-EA Stop
						Write-Host -f darkgreen "Senden ALS : $((Get-ADDomain).DNSRoot+"\"+$SamAccountName) zu $($SendAsUser) zugefügt ...."
					}
					# ---------------------------------------------------------------
					#
					# Senden im Auftrag von
					#
					if ($SendOnBeHalf) {
						$userAdd = ($SamAccountName+"@"+(Get-ADDomain).DNSRoot)
					
						try {
							Set-Mailbox $SendOnBeHalfuser –Grantsendonbehalfto @{'+'=$userAdd}
								
							Write-Host -f darkgreen "Senden im Auftrag von : $($userAdd) zu $($SendOnBeHalfuser) zugefügt ...."
						} catch {
							write-host -f red "Senden im Auftrag von : $($userAdd) konnte NICHT zu $($SendOnBeHalfuser) zugefügt werden!"
						}
					}				
					$str = "Mailbox: $($SamAccountName) angelegt."
					Write-host -f darkgreen $str
				} else {
					$str = "Mailbox: $($SamAccountName) konnte nicht angelegt werden."
					Write-host -f red $str
				
				}
			} catch {
				$_ | out-host
				$str = "Mailbox: $($SamAccountName) konnte nicht angelegt werden."
				Write-Host -f red $str
			}
		} else {
			$str = "Mailbox: $($SamAccountName) existiert schon." 
			write-host -f red $str
		}
		
		try { Remove-PSSession -Session $exchangeSession -EA Stop } catch {Write-Host -f red "Remove-Session failed"}
	} 
	Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
Function New-EmergencyOfficeUser {
[CmdletBinding()]
Param	(
			$OUPath,
			$Group,
			$Startuser,
			$MaxUser,
			$PassWord,
			$ExchangeDatabase,
			$OfficeVersion,
			$OfficeVersionFull

		)
	$CSVLines = @()
	#
	# Excel
	#
	$AppName 		= "EX"
	$AppNameFull 	= "Excel"
	$User 			= $OfficeVersion+$AppName
	For ($i=$Startuser; $i -le $MaxUser; $i++) {
		$UserSuffix = $i.ToString().PadLeft(2,"0")

		New-AdUser-NH		-SamAccountName ($User+"-"+$UserSuffix) -GivenName ($AppNameFull+" "+$UserSuffix) -SurName $OfficeVersionFull -Password $PassWord -OUPathDN $OUPath -GroupDN $Group
		$Line = "$($User+'-'+$UserSuffix)"+","+$PassWord+","+"$($AppNameFull+' '+$UserSuffix)"+","+$OfficeVersionFull+","+"$($User+'-'+$UserSuffix)"+"@wihan.nh"+","+"Hannover"
		$CSVLines += $Line
	}
	#
	# Word
	#
	$AppName 		= "WO"
	$AppNameFull 	= "Word"
	$User 			= $OfficeVersion+$AppName
	For ($i=$Startuser; $i -le $MaxUser; $i++) {
		$UserSuffix = $i.ToString().PadLeft(2,"0")

		New-AdUser-NH		-SamAccountName ($User+"-"+$UserSuffix) -GivenName ($AppNameFull+" "+$UserSuffix) -SurName $OfficeVersionFull -Password $PassWord -OUPathDN $OUPath -GroupDN $Group
		$Line = "$($User+'-'+$UserSuffix)"+","+$PassWord+","+"$($AppNameFull+' '+$UserSuffix)"+","+$OfficeVersionFull+","+"$($User+'-'+$UserSuffix)"+"@wihan.nh"+","+"Hannover"
		$CSVLines += $Line
	}
	#
	# PowerPoint
	#
	$AppName 		= "PO"
	$AppNameFull 	= "PowerPoint"
	$User 			= $OfficeVersion+$AppName
	For ($i=$Startuser; $i -le $MaxUser; $i++) {
		$UserSuffix = $i.ToString().PadLeft(2,"0")

		New-AdUser-NH		-SamAccountName ($User+"-"+$UserSuffix) -GivenName ($AppNameFull+" "+$UserSuffix) -SurName $OfficeVersionFull -Password $PassWord -OUPathDN $OUPath -GroupDN $Group
		$Line = "$($User+'-'+$UserSuffix)"+","+$PassWord+","+"$($AppNameFull+' '+$UserSuffix)"+","+$OfficeVersionFull+","+"$($User+'-'+$UserSuffix)"+"@wihan.nh"+","+"Hannover"
		$CSVLines += $Line
	}
	#
	# Access
	#
	$AppName 		= "AC"
	$AppNameFull 	= "Access"
	$User 			= $OfficeVersion+$AppName
	For ($i=$Startuser; $i -le $MaxUser; $i++) {
		$UserSuffix = $i.ToString().PadLeft(2,"0")

		New-AdUser-NH		-SamAccountName ($User+"-"+$UserSuffix) -GivenName ($AppNameFull+" "+$UserSuffix) -SurName $OfficeVersionFull -Password $PassWord -OUPathDN $OUPath -GroupDN $Group
		$Line = "$($User+'-'+$UserSuffix)"+","+$PassWord+","+"$($AppNameFull+' '+$UserSuffix)"+","+$OfficeVersionFull+","+"$($User+'-'+$UserSuffix)"+"@wihan.nh"+","+"Hannover"
		$CSVLines += $Line
	}
	#
	# Outlook
	#
	$AppName 		= "OU"
	$AppNameFull 	= "Outlook"
	$User 			= $OfficeVersion+$AppName
	For ($i=$Startuser; $i -le $MaxUser; $i++) {
		$UserSuffix = $i.ToString().PadLeft(2,"0")

		New-AdUser-NH		-SamAccountName ($User+"-"+$UserSuffix) -GivenName ($AppNameFull+" "+$UserSuffix) -SurName $OfficeVersionFull -Password $PassWord -OUPathDN $OUPath -GroupDN $Group

		Enable-Mailbox-NH	-SamAccountName ($User+"-"+$UserSuffix) -Alias 	  ($User+"-"+$UserSuffix) -Databasename $ExchangeDatabase -SendAs:$False -SendAsUser "" -SendOnBeHalf:$True -SendOnBeHalfUser "Fred Fredsen"
		$Line = "$($User+'-'+$UserSuffix)"+","+$PassWord+","+"$($AppNameFull+' '+$UserSuffix)"+","+$OfficeVersionFull+","+"$($User+'-'+$UserSuffix)"+"@wihan.nh"+","+"Hannover"
		$CSVLines += $Line
	}
	#
	# SPECIAL
	#
	if ($OfficeVersion -eq "O10") {
		$AppName 		= "W7"
		$AppNameFull 	= "Windows 7"
		$User 			= $OfficeVersion+$AppName
		For ($i=$Startuser; $i -le $MaxUser; $i++) {
			$UserSuffix = $i.ToString().PadLeft(2,"0")

			New-AdUser-NH		-SamAccountName ($User+"-"+$UserSuffix) -GivenName ($AppNameFull+" "+$UserSuffix) -SurName $OfficeVersionFull -Password $PassWord -OUPathDN $OUPath -GroupDN $Group
			$Line = "$($User+'-'+$UserSuffix)"+","+$PassWord+","+"$($AppNameFull+' '+$UserSuffix)"+","+$OfficeVersionFull+","+"$($User+'-'+$UserSuffix)"+"@wihan.nh"+","+"Hannover"
			$CSVLines += $Line
		}		
	}
	
	
	Write-Output $CSVLines
}
#
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
#region MAIN
# #############################################################################
# ##### MAIN
# #############################################################################

	if (!(Test-IsAdmin)) {
		Write-Host -F Red "#####################################################################################################################"
		Write-Host -F Red "Bitte führe das Script mit einem Administrator-Konto aus."
		Write-Host -F Red "#####################################################################################################################"
		return -1
	}
	try {
		Import-Module ActiveDirectory -EA Stop
	} catch {
		Write-Host -F Red "#####################################################################################################################"
		Write-Host -F Red "Das Modul ActiveDirectory konnte nicht geladen werden."
		Write-Host -F Red "#####################################################################################################################"
		return -1
	}
	#
	# ~~~ PREPARATION ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	$global:RTCred = [System.Management.Automation.PSCredential]::Empty

	if (Test-path -PathType Leaf ".\XXL-Credentials.ps1" ) {
		. $(Resolve-Path ".\XXL-Credentials.ps1")
	} else {
		$global:RTCred  = Get-Credential
	}

	$script:exchangeServer = Get-ExchangeServerInSite-AAN | where {$_.roles -ilike "*CAS*" } | Select-Object -first 1
	# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	clear-host
	$CSVLines = @()
<#	
	$Lines = New-EmergencyOfficeUser	`
				-OUPath 			"OU=Office 2003,OU=Emergency,OU=Office User,OU=WalkIn,DC=wihan,DC=nh"	`
				-Group  			"CN=Teilnehmer,CN=Users,DC=wihan,DC=nh"	`
				-Startuser 			15 	`
				-MaxUser			28 `
				-PassWord 			"Password1" `
				-ExchangeDatabase	"EmergencyDB" `
				-OfficeVersion		"O03" `
				-OfficeVersionFull	"Office 2003"
	$CSVLines += $Lines
#>
<#			
	$Lines = New-EmergencyOfficeUser	`
				-OUPath 			"OU=Office 2007,OU=Emergency,OU=Office User,OU=WalkIn,DC=wihan,DC=nh"	`
				-Group  			"CN=Teilnehmer,CN=Users,DC=wihan,DC=nh"	`
				-Startuser 			15 	`
				-MaxUser			28 `
				-PassWord 			"Password1" `
				-ExchangeDatabase	"EmergencyDB" `
				-OfficeVersion		"O07" `
				-OfficeVersionFull	"Office 2007"
	$CSVLines += $Lines
#>				
	$Lines = New-EmergencyOfficeUser	`
				-OUPath 			"OU=Office 2010,OU=Emergency,OU=Office User,OU=WalkIn,DC=wihan,DC=nh"	`
				-Group  			"CN=Teilnehmer,CN=Users,DC=wihan,DC=nh"	`
				-Startuser 			15 	`
				-MaxUser			28 `
				-PassWord 			"Password1" `
				-ExchangeDatabase	"EmergencyDB" `
				-OfficeVersion		"O10" `
				-OfficeVersionFull	"Office 2010"
	$CSVLines += $Lines
<#				
	$Lines = New-EmergencyOfficeUser	`
				-OUPath 			"OU=Office 2013,OU=Emergency,OU=Office User,OU=WalkIn,DC=wihan,DC=nh"	`
				-Group  			"CN=Teilnehmer,CN=Users,DC=wihan,DC=nh"	`
				-Startuser 			15 	`
				-MaxUser			28 `
				-PassWord 			"Password1" `
				-ExchangeDatabase	"EmergencyDB" `
				-OfficeVersion		"O13" `
				-OfficeVersionFull	"Office 2013"
	$CSVLines += $Lines
				
	$Lines = New-EmergencyOfficeUser	`
				-OUPath 			"OU=Office 2016,OU=Emergency,OU=Office User,OU=WalkIn,DC=wihan,DC=nh"	`
				-Group  			"CN=Teilnehmer,CN=Users,DC=wihan,DC=nh"	`
				-Startuser 			15 	`
				-MaxUser			28 `
				-PassWord 			"Password1" `
				-ExchangeDatabase	"EmergencyDB" `
				-OfficeVersion		"O16" `
				-OfficeVersionFull	"Office 2016"
	$CSVLines += $Lines
#>

	"username, password, firstname, lastname, email, city" | out-file ".\Moodle.csv" -Force -Encoding UTF8
	$CSVLines | out-file ".\Moodle.csv" -Encoding UTF8 -Append
	Write-Host "Fertig !!!! .... bitte die Datei Moodle.csv im Arbeitsverzeichnis für den Import in Moodle benutzen."
# #############################################################################
# ##### END MAIN
# #############################################################################
#endregion MAIN