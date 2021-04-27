# Set the values as needed
$WindowsUpdateRegKey = "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU"
$WindowsUpdateRootRegKey = "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\"
$WSUSServer          = "http://www.xxx.yyy.zzz:8530"
$StatServer          = "http://www.xxx.yyy.zzz:8530"
$Enabled             = 1

# Test if the Registry Key doesn't exist already
if(-not (Test-Path $WindowsUpdateRegKey))
{
    # Create the WindowsUpdate\AU key, since it doesn't exist already
    # The -Force parameter will create any non-existing parent keys recursively
    New-Item -Path $WindowsUpdateRegKey -Force
}

# Enable an Intranet-specific WSUS server
Set-ItemProperty -Path $WindowsUpdateRegKey -Name UseWUServer -Value $Enabled -Type DWord
Set-ItemProperty -Path $WindowsUpdateRegKey -Name AUOptions -Value 3 -Type DWord
Set-ItemProperty -Path $WindowsUpdateRegKey -Name AutoInstallMinorUpdates -Value $Enabled -Type DWord
Set-ItemProperty -Path $WindowsUpdateRegKey -Name DetectionFrequency -Value 4 -Type DWord
Set-ItemProperty -Path $WindowsUpdateRegKey -Name DetectionFrequencyEnabled -Value $Enabled -Type DWord
Set-ItemProperty -Path $WindowsUpdateRegKey -Name IncludeRecommendedUpdates -Value $Enabled -Type DWord
Set-ItemProperty -Path $WindowsUpdateRegKey -Name NoAUAsDefaultShutdownOption -Value $Enabled -Type DWord
Set-ItemProperty -Path $WindowsUpdateRegKey -Name NoAUShutdownOption -Value $Enabled -Type DWord
Set-ItemProperty -Path $WindowsUpdateRegKey -Name NoAutoRebootWithLoggedOnUsers -Value $Enabled -Type DWord
Set-ItemProperty -Path $WindowsUpdateRegKey -Name NoAutoUpdate -Value 0 -Type DWord
Set-ItemProperty -Path $WindowsUpdateRegKey -Name ScheduledInstallDay -Value 0 -Type DWord
Set-ItemProperty -Path $WindowsUpdateRegKey -Name ScheduledInstallTime -Value 3 -Type DWord


# Specify the WSUS server
Set-ItemProperty -Path $WindowsUpdateRootRegKey -Name WUServer -Value $WSUSServer -Type String

# Specify the Statistics server
Set-ItemProperty -Path $WindowsUpdateRootRegKey -Name WUStatusServer -Value $StatServer -Type String

# Restart Windows Update Service
Restart-Service wuauserv -Force
