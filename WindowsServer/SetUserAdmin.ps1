$sid2 = 'S-1-5-32-544'
$objSID2 = New-Object System.Security.Principal.SecurityIdentifier($sid2)
$localadminsgroup = (($objSID2.Translate([System.Security.Principal.NTAccount])).Value).Split("\")[1]
$Password = ConvertTo-SecureString "PASSWORD" -AsPlainText -Force
New-LocalUser svc_pdq -Password $Password -ErrorAction SilentlyContinue
Add-LocalGroupMember -Group $localadminsgroup -Member USER
Set-LocalUser USER -Password $Password -PasswordNeverExpires 1 -ErrorAction SilentlyContinue
