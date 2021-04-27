net user USER PASSWORD /add /y
net localgroup administrators USER /add
WMIC USERACCOUNT WHERE Name='USER' SET PasswordExpires=FALSE
