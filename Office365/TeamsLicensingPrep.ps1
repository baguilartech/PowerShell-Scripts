<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2019 v5.6.160
	 Created on:   	4/8/2020 2:57 PM
	 Created by:   	Boris Aguilar
	 Organization: 	
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		Takes a text file containing O365 email usernames, separated by new lines, and assigns
		a Microsoft Audio Conferencing Licenses to all the users and sets Toll-free and dial-out policies
#>

#Import Module
ipmo MSOnline
Import-Module "C:\\Program Files\\Common Files\\Skype for Business Online\\Modules\\SkypeOnlineConnector\\SkypeOnlineConnector.psd1"

#Authenticate to MSOLservice
Connect-MSOLService
#File prompt to select the userlist txt file
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$OFD = New-Object System.Windows.Forms.OpenFileDialog
$OFD.filter = "text files (*.*)| *.txt"
$OFD.ShowDialog() | Out-Null
$OFD.filename

If ($OFD.filename -eq '')
{
	Write-Host "You did not choose a file. Try again" -ForegroundColor White -BackgroundColor Red
}

#Create a variable of all users
$users = Get-Content $OFD.filename

#License each user in the $users variable
foreach ($user in $users)
{
	Write-host "Assigning License: $user"
	Set-MsolUserLicense -UserPrincipalName $user -AddLicenses "centricsoftware:MCOMEETADV"
    Write-host "Removing Toll Free Dial-in: $user"
    Set-CsOnlineDialInConferencingUser $user -AllowTollFreeDialIn $false
    Write-host "Removing Dial-out for CPC and PSTN: $user"
    Grant-CSDialOutPolicy -Identity $user -PolicyName "DialoutCPCandPSTNDisabled" 
}

