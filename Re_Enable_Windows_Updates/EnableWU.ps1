Write-Host "Re-enabling Windows Updates..." -ForegroundColor Cyan

$registryItems = @(
    @{ Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"; Name = "SetDisableUXWUAccess" },
    @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"; Name = "SettingsPageVisibility" },
    @{ Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"; Name = "SettingsPageVisibility" },
    @{ Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"; Name = "NoAutoUpdate" },
    @{ Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"; Name = "AUOptions" },
    @{ Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"; Name = "DisableWindowsUpdateAccess" }
)

foreach ($item in $registryItems) {
    if (Test-Path $item.Path) {
        try {
            if (Get-ItemProperty -Path $item.Path -Name $item.Name -ErrorAction SilentlyContinue) {
                Write-Host "Removing $($item.Name) from $($item.Path)" -ForegroundColor Yellow
                Remove-ItemProperty -Path $item.Path -Name $item.Name -ErrorAction SilentlyContinue
            }
        } catch {
            Write-Warning "Failed to remove $($item.Name) from $($item.Path): $_"
        }
    }
}

try {
    Write-Host "Configuring Windows Update service..." -ForegroundColor Cyan
    Set-Service -Name wuauserv -StartupType Manual
    Start-Service -Name wuauserv
} catch {
    Write-Warning "Failed to configure or start wuauserv: $_"
}

try {
    gpupdate /force | Out-Null
    Write-Host "Group policy updated." -ForegroundColor Green
} catch {
    Write-Warning "Group policy update failed: $_"
}

Write-Host "Windows Update re-enable complete. Please reboot to finalize." -ForegroundColor Green
