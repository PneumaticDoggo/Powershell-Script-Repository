if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Write-Host "Installing PSWindowsUpdate module..."
    Install-Module -Name PSWindowsUpdate -Force -SkipPublisherCheck -AllowClobber
}

Import-Module PSWindowsUpdate

Write-Host "Checking for Windows updates..."
$Updates = Get-WindowsUpdate

if ($Updates.Count -eq 0) {
    Write-Host "System up to date"
} else {
    Write-Host "The following updates are available:"
    $Updates | ForEach-Object {
        Write-Host $_.Title
    }

    Write-Host "Installing updates..."
    Install-WindowsUpdate -AcceptAll -AutoReboot
}
