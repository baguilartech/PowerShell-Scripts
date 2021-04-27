#Imports the MSOnline Module if you haven't loaded Azure PowerShell
Import-Module MsOnline -ErrorAction SilentlyContinue

#Connects to your Office365 tenant
Connect-MsolService

#Gets the Group via it's displayName and inputs into the $groupid variable. Change 'Test Security Group' to your Group Name
$groupid = Get-MsolGroup | Where-Object {$_.DisplayName -eq "All"}

#Gets the Users where the DisplayName is like 'test' you may need to change this to select a group of users you are after, I personally Export to CSV remove the users that are not needed and Import-Csv back into ps
$users = Get-MsolUser -All | Export-Csv "C:\Users\*****\Documents\******.csv"
