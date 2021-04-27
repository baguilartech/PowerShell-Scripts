<#        
    .NAME
     CentricSoftwareCarePackage
     
    .SYNOPSIS
     Creates CSIAdmin and installs LMI

    .DESCRIPTION
     This installer checks for the IT administrative account and creates the account if it does not exist. 
     If the account does exist, then the account password is reset for sanity.
     The installer then runs a LogMeIn silent installation.

    .NOTES
    ========================================================================
         Windows PowerShell Source File 
         Created with SAPIEN Technologies PrimalScript 2019
         
         NAME: CentricSoftwareCarePackage
         
         AUTHOR: baguilar@centricsoftware.com, 
         DATE  : 3/31/2020
         
         COMMENT: 
         
    ==========================================================================
#>

# Get the ID and security principal of the current user account
$myWindowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent();
$myWindowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($myWindowsID);

# Get the security principal for the administrator role
$adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator;

# Check to see if we are currently running as an administrator
if ($myWindowsPrincipal.IsInRole($adminRole))
{
	# We are running as an administrator, so change the title and background colour to indicate this
	$Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)";
	$Host.UI.RawUI.BackgroundColor = "DarkBlue";
	Clear-Host;
}
else
{
	# We are not running as an administrator, so relaunch as administrator
	
	# Create a new process object that starts PowerShell
	$newProcess = New-Object System.Diagnostics.ProcessStartInfo "PowerShell";
	
	# Specify the current script path and name as a parameter with added scope and support for scripts with spaces in it's path
	$newProcess.Arguments = "& '" + $script:MyInvocation.MyCommand.Path + "'"
	
	# Indicate that the process should be elevated
	$newProcess.Verb = "runas";
	
	# Start the new process
	[System.Diagnostics.Process]::Start($newProcess);
	
	# Exit from the current, unelevated, process
	exit;
}

#Initialize Logging
$LogPath = ".\"
$LogPathName = Join-Path -Path $LogPath -ChildPath "CSCP-$(Get-Date -Format 'MM-dd-yyyy').log"
Start-Transcript $LogPathName -Append

$ErrorActionPreference = 'Continue'
$VerbosePreference = 'Continue'

#Look for LocalAdmins Group and store the var
$sid2 = 'S-1-5-32-544'
$objSID2 = New-Object System.Security.Principal.SecurityIdentifier($sid2)
$localadminsgroup = (($objSID2.Translate([System.Security.Principal.NTAccount])).Value).Split("\")[1]

#User to search for
$USERNAME = "******"

#Declare LocalUser Object
$ObjLocalUser = $null

#Declare Password
$Password = ConvertTo-SecureString "*********" -AsPlainText -Force

#Find csiadmin
Try
{
	Write-Verbose "Searching for $($USERNAME) in LocalUser DataBase"
	$ObjLocalUser = Get-LocalUser $USERNAME
	Write-Verbose "User $($USERNAME) was found"
}

#Error-checking
Catch [Microsoft.PowerShell.Commands.UserNotFoundException] {
	"User $($USERNAME) was not found" | Write-Warning
}

Catch
{
	"An unspecifed error occured" | Write-Error
	Exit # Stop Powershell! 
}

#Reset the password for sanity.
if ($ObjLocalUser)
{
	Write-Verbose "Resetting $($USERNAME) for sanity"
	Set-LocalUser $USERNAME -Password $Password -ErrorAction SilentlyContinue
	Add-LocalGroupMember -Group $localadminsgroup -Member $USERNAME
	Write-Verbose "Added $($USERNAME) to LocalAdmins"
}

#Create the user if it was not found
If (!$ObjLocalUser)
{
	Write-Verbose "Creating User $($USERNAME)"
	New-LocalUser csiadmin -Password $Password -ErrorAction SilentlyContinue
	Add-LocalGroupMember -Group $localadminsgroup -Member $USERNAME
	Write-Verbose "Added $($USERNAME) to LocalAdmins"
}

Write-Host "Installing LMI Remote Tools for IT"
Start-Process -FilePath ".\LogMeIn.msi" -ArgumentList '/quiet DEPLOYID=01_*********************************** INSTALLMETHOD=5 FQDNDESC=0' -Wait
Write-Host "Done!"
Write-Host -NoNewLine "Closing";

Stop-Transcript

exit;
