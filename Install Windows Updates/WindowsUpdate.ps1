Install-Module -Name PSWindowsUpdate -Force -AllowClobber

Set-ExecutionPolicy RemoteSigned

Import-Module PSWindowsUpdate

Get-WindowsUpdate

Install-WindowsUpdate -AcceptAll -AutoReboot