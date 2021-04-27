<#	

        Exchange Online Powershell is required for message trace.
		https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell?view=exchange-ps
		
		ReportHTML Moduile is required
        Install-Module -Name ReportHTML
        https://www.powershellgallery.com/packages/ReportHTML/

	.DESCRIPTION
		Compare four weeks of DLWeeklyInactivity report results from your O365 tenant. Removes Weekly reports older than 5 weeks, sends detailed HTML report on unused distribution lists.

#>

#Connection info
$Username = "*************"
$PasswordPath = "C:\Users\**********\Documents\secrets\*********.txt"

#Read the password from the file and convert to SecureString
$SecurePassword = Get-Content $PasswordPath | ConvertTo-SecureString

#Build a Credential Object from the password file and the $username constant
$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $SecurePassword

#Open a session to O365
$ExOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication  Basic -AllowRedirection
import-PSSession $ExOSession -AllowClobber

#Set the constants
$Date = get-date -format MMddyyyy
$ReportsFolder = "C:\Users\**********\Documents\Reports\"
$CompanyLogo = "https://www.freelogodesign.org/Content/img/logo-ex-4.png"
$Table = New-Object 'System.Collections.Generic.List[System.Object]'
$RemovedFilesTable = New-Object 'System.Collections.Generic.List[System.Object]'

#Get report run date for previous weekly reports
$Week1Date = (get-date).AddDays(-21).ToString("MMddyyyy")
$Week2Date = (get-date).AddDays(-14).ToString("MMddyyyy")
$Week3Date = (get-date).AddDays(-7).ToString("MMddyyyy")
$Week4Date = (get-date).ToString("MMddyyyy")

#Clean up weekly reports created more than 35 days ago
$ToOldFiles = ("$ReportsFolder"+"Inactive*"+".txt")
$FilestoRemove = Get-ChildItem -Path $ToOldFiles -Force | Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-35) }
 
Foreach ($File in $FilestoRemove) {
	Remove-Item $File -Force
		if (!$?) {
		$ReportName =  "$($File.name)"
		$Status 	= "Failed to delete the file automatically."
        }
		
		Else {
		$ReportName = "$($File.name)"
		$Status 	= "Successfully deleted the file automatically."
		}
		
		$obj = [PSCustomObject]@{
		'File Name'		   	   = $ReportName
		'Deleted Status'	   = $Status	
		 }
		
		$RemovedFilesTable.add($obj)
}

If (($RemovedFilesTable).count -eq 0)
{
	$RemovedFilesTable = [PSCustomObject]@{
		'Information'  = 'Information: No Inactive Weekly Lists were found to remove.'
	}
}	

#Set up the weekly report file paths
$Week1Path = ("$ReportsFolder"+"Inactive"+"$Week1Date"+".txt")
$Week2Path = ("$ReportsFolder"+"Inactive"+"$Week2Date"+".txt")
$Week3Path = ("$ReportsFolder"+"Inactive"+"$Week3Date"+".txt")
$Week4Path = ("$ReportsFolder"+"Inactive"+"$Week4Date"+".txt")

#Input weekly report files
$Week1Report = Get-Content $Week1Path
$Week2Report = Get-Content $Week2Path
$Week3Report = Get-Content $Week3Path
$Week4Report = Get-Content $Week4Path

#Compare weekly report files
$Week12Results =  Compare-Object -ReferenceObject $Week1Report -DifferenceObject $Week2Report -ExcludeDifferent -IncludeEqual
$Week23Results =  Compare-Object -ReferenceObject $Week12Results.InputObject -DifferenceObject $Week3Report -ExcludeDifferent -IncludeEqual
$Week34Results =  Compare-Object -ReferenceObject $Week23Results.InputObject -DifferenceObject $Week4Report -ExcludeDifferent -IncludeEqual

#Filter slider object out of the results
$MonthlyInactive = $Week34Results.InputObject

#Set export file name for plain text file to be used by quarterly report
$DLMonthlyActivityTxt = ("$ReportsFolder"+"MonthlyInactive"+"$date"+".txt")

#Export the findings to the text file
$MonthlyInactive | Out-File $DLMonthlyActivityTxt

#Get inactive distribution list details and create HTML report
Foreach ($List in $MonthlyInactive) {

		$ListDetails = get-DistributionGroup $List
		$DisplayName = $ListDetails.DisplayName
		$Email = $ListDetails.PrimarySMTPAddress
		$Synced = $ListDetails.IsDirSynced
		$Owner = ($ListDetails.ManagedBy) -join ", "
		$Members = (Get-DistributionGroupMember $List | Sort-Object Name | Select-Object -ExpandProperty Name) -join ", "
		$MeasureMembers = $Members | measure 
		$NumberofMembers = $MeasureMembers.count
		
		$obj = [PSCustomObject]@{
		'Name'				   = $DisplayName
		'Email Address'	       = $Email
		'AD Synced'			   = $Synced
		'Owners'			   = $Owner
		'Members'			   = $Members	
	}
	
	$Table.add($obj)
}

If (($Table).count -eq 0)
{
	$Table = [PSCustomObject]@{
		'Information'  = 'Information: No distribution lists are inactive.'
	}
}

$rpt = New-Object 'System.Collections.Generic.List[System.Object]'
$rpt += get-htmlopenpage -TitleText 'Monthly Inactive Distribution List Report' -LeftLogoString $CompanyLogo 

		$rpt += Get-HTMLContentOpen -HeaderText "Distribution lists that have not been emailed in 4 weeks."
            $rpt += get-htmlcontentdatatable $Table -HideFooter
        $rpt += Get-HTMLContentClose
		$rpt += Get-HTMLContentOpen -HeaderText "Weekly Inactive Reports not created in the past 5 weeks."
		    $rpt += get-htmlcontentdatatable $RemovedFilesTable -HideFooter
	    $rpt += Get-HTMLContentClose
		
$rpt += Get-HTMLClosePage

$rpt += Get-HTMLClosePage
$ReportName = ("DLMonthlyInactiveReport" + "$Date")
Save-HTMLReport -ReportContent $rpt -ShowReport -ReportName $ReportName -ReportPath $ReportsFolder
$MonthlyReport = ("$ReportsFolder"+"$ReportName"+".html")

#Send an email with the findings
$From = "***********"
$To = "************"
$Subject = "Monthly Inactive Distribution List Report"
$Body = "See the attached report for distribution lists that have not been emailed in the past 4 weeks. A .txt file has been saved in the file share to be accessed by the Quarterly DL Inactivity Report script. Do not modify any of the weekly or monthly .txt file master copies in the share."
$SMTPServer = "smtp.office365.com"
$SMTPPort = "587"

Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $Credential -Attachments $MonthlyReport

#Close the session to O365
Remove-PSSession $ExOSession
