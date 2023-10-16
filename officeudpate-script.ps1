# Install PSWindowsUpdate PowerShell module if needed
if (!(Get-Module -Name PSWindowsUpdate -ListAvailable)) {
    Write-Output "PSWindowsUpdate module not found. Installing module..."
    Install-Module -Name PSWindowsUpdate -Scope AllUsers -Force
    Import-Module -Name PSWindowsUpdate
} else {
    Write-Output "PSWindowsUpdate module already installed."
}

# Check for Office Click-To-Run Products
$officeC2R = Get-ItemProperty HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*,HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Microsoft Office Professional Plus 2019*" -or $_.DisplayName -like "*Microsoft Office Professional Plus 2021*" -or $_.DisplayName -like "*Microsoft Office 365*" -or $_.DisplayName -like "*Microsoft 365*"}

# Update Click-To-Run Office Products (Office 2019, 2021, 365, etc)
if ($officeC2R -ne $null) {
    if (Test-Path "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe") {
        Write-Output "Click-To-Run Office detected. Initiating update."
        & "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user displaylevel=false forceappshutdown=true
    }
    else {
        Write-Output "No Click-To-Run Office detected."
    }
}

# Temporarily disable WSUS
$wsusRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
$wsusValue = Get-ItemPropertyValue -Path '$wsusRegPath' -Name UseWUServer -ErrorAction SilentlyContinue
if ($wsusValue -ne $null) {
    Write-Output "Disabling WSUS"
    Set-ItemProperty -Path $wsusRegPath -Name UseWUServer -Value 0
}

# Temporarily disable Windows Update for Business deferral period
$wufbRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
$wufbValue = (Get-ItemProperty -Path $wufbRegPath -Name DeferQualityUpdatesPeriodInDays -ErrorAction SilentlyContinue).DeferQualityUpdatesPeriodInDays
if ($wufbValue -ne $null) {
    if ($wufbValue -ne 0) {
        Write-Output "Disabling Windows Update for Business deferral period"
        Set-ItemProperty -Path $wufbRegPath -Name DeferQualityUpdatesPeriodInDays -Value 0
    }
    else {
        Write-Output "WUfB deferral period already zero"
    }
}

# Check for Office 2013 and 2016
$office2013 = Get-ItemProperty HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*,HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Microsoft Office Professional Plus 2013*"}
$office2016 = Get-ItemProperty HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*,HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Microsoft Office Professional Plus 2016*"}

# Check if Office 2013 is installed and if the KB5002265 update is installed
if ($office2013 -ne $null) {
    $KB5002265_installed = Get-WindowsUpdate -KBArticleID KB5002265 -IsInstalled

    # If the KB5002265 update is not installed, install it
    if (!$KB5002265_installed) {
        Write-Output "Installing KB5002265 for Office 2013"
        Install-WindowsUpdate -KBArticleID KB5002265 -MicrosoftUpdate -IgnoreReboot -Verbose -Confirm:$false
    }
    else {
        Write-Output "No Outlook 2013 CVE-2023-23397 vulnerability"
    }
}
# Check if Office 2016 is installed and if the KB5002254 update is installed
if ($office2016 -ne $null) {
    $KB5002254_installed = Get-WindowsUpdate -KBArticleID KB5002254 -IsInstalled

    # If the KB5002254 update is not installed, install it
    if (!$KB5002254_installed) {
        Write-Output "Installing KB5002254 for Office 2016"
        Install-WindowsUpdate -KBArticleID KB5002254 -MicrosoftUpdate -IgnoreReboot -Verbose -Confirm:$false
    }
    else {
        Write-Output "No Outlook 2016 CVE-2023-23397 vulnerability"
    }
}

# Return UseWUServer to previous value
if ($wsusValue -ne $null) {
    Write-Output "Enabling WSUS"
    Set-ItemProperty -Path $wsusRegPath -Name UseWUServer -Value $wsusValue
}

# Return DeferQualityUpdatesPeriodInDays to previous value
if ($wufbValue -ne $null) {
    Write-Output "Enabling Windows Update for Business deferral period"
    Set-ItemProperty -Path $wufbRegPath -Name DeferQualityUpdatesPeriodInDays -Value $wufbValue
}

# Reboot if any pending updates
#Get-WURebootStatus -AutoReboot