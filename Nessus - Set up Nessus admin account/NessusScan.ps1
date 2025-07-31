$adminUsername = ""  
$adminPassword = "" 

if (-not (Get-LocalUser -Name $adminUsername -ErrorAction SilentlyContinue)) {
    New-LocalUser -Name $adminUsername -Password (ConvertTo-SecureString $adminPassword -AsPlainText -Force) -FullName "Nessus Admin" -Description "Administrator for Nessus Scanning"
    Add-LocalGroupMember -Group "Administrators" -Member $adminUsername
}

Set-Service -Name "RemoteRegistry" -StartupType Manual
Start-Service -Name "RemoteRegistry"

$registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\system"
$registryName = "LocalAccountTokenFilterPolicy"
$registryValue = 1

if (-not (Get-ItemProperty -Path $registryPath -Name $registryName -ErrorAction SilentlyContinue)) {
    New-ItemProperty -Path $registryPath -Name $registryName -Value $registryValue -PropertyType DWord
} else {
    Set-ItemProperty -Path $registryPath -Name $registryName -Value $registryValue
}


New-NetFirewallRule -DisplayName "File and Printer Sharing - Echo Request" -Direction Inbound -Protocol ICMPv4 -RemoteAddress Any -Action Allow
New-NetFirewallRule -DisplayName "SMB-In" -Direction Inbound -Protocol TCP -LocalPort 445 -Action Allow

New-NetFirewallRule -DisplayName "WMI-In" -Direction Inbound -Protocol TCP -LocalPort 135 -Action Allow
New-NetFirewallRule -DisplayName "DCOM-In" -Direction Inbound -Protocol TCP -LocalPort 135 -Action Allow
New-NetFirewallRule -DisplayName "ASync-In" -Direction Inbound -Protocol TCP -LocalPort 49152-65535 -Action Allow
