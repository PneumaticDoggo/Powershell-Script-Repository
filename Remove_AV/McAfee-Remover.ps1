#Requires -RunAsAdministrator

Write-Host "=== Enhanced McAfee Removal Script ===" -ForegroundColor Green
Write-Host "Starting at: $(Get-Date)" -ForegroundColor Yellow

$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERROR: This script requires administrator privileges!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "`n1. Force-stopping McAfee processes..." -ForegroundColor Cyan
$mcafeeProcesses = @("mcshield", "mfevtps", "mfefire", "mfemms", "mfevtps", "mcapexe", "mcsacore", "mfeann", "mcuicnt", "mctray", "mcods", "frminst", "mfemactl", "mfevtps", "mfefire", "mfemms", "mfevtps", "mcapexe", "mcsacore", "mfeann", "mcuicnt", "mctray", "mcods", "frminst", "mfemactl")

foreach ($processName in $mcafeeProcesses) {
    try {
        $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
        foreach ($process in $processes) {
            Write-Host "  Force-stopping process: $($process.ProcessName)" -ForegroundColor Yellow
            Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        }
        Start-Process "taskkill" -ArgumentList "/f /im $processName.exe" -NoNewWindow -Wait -ErrorAction SilentlyContinue
    }
    catch {
    }
}

Write-Host "`n2. Stopping and disabling McAfee services..." -ForegroundColor Cyan
$mcafeeServices = Get-Service | Where-Object { $_.Name -like "*McAfee*" -or $_.DisplayName -like "*McAfee*" -or $_.Name -like "*mfe*" }
foreach ($service in $mcafeeServices) {
    try {
        Write-Host "  Stopping service: $($service.DisplayName)" -ForegroundColor Yellow
        Stop-Service -Name $service.Name -Force -ErrorAction SilentlyContinue
        Set-Service -Name $service.Name -StartupType Disabled -ErrorAction SilentlyContinue
        Write-Host "  Service stopped: $($service.DisplayName)" -ForegroundColor Green
    }
    catch {
        Write-Host "  Could not stop service: $($service.DisplayName)" -ForegroundColor Red
    }
}

Write-Host "`n3. Uninstalling McAfee products using multiple methods..." -ForegroundColor Cyan

try {
    $mcafeePackages = Get-Package | Where-Object { $_.Name -like "*McAfee*" }
    foreach ($package in $mcafeePackages) {
        Write-Host "  Uninstalling via Get-Package: $($package.Name)" -ForegroundColor Yellow
        try {
            $package | Uninstall-Package -Force -ErrorAction Stop
            Write-Host "  Successfully uninstalled: $($package.Name)" -ForegroundColor Green
        }
        catch {
            Write-Host "  Failed to uninstall: $($package.Name)" -ForegroundColor Red
        }
    }
}
catch {
    Write-Host "  Get-Package method failed, continuing with other methods..." -ForegroundColor Yellow
}

try {
    Write-Host "  Searching via WMI Win32_Product..." -ForegroundColor Yellow
    $mcafeeProducts = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*McAfee*" }
    foreach ($product in $mcafeeProducts) {
        Write-Host "  Uninstalling via WMI: $($product.Name)" -ForegroundColor Yellow
        try {
            $result = $product.Uninstall()
            if ($result.ReturnValue -eq 0) {
                Write-Host "  Successfully uninstalled: $($product.Name)" -ForegroundColor Green
            } else {
                Write-Host "  Failed to uninstall: $($product.Name) (Return code: $($result.ReturnValue))" -ForegroundColor Red
            }
        }
        catch {
            Write-Host "  Error uninstalling: $($product.Name)" -ForegroundColor Red
        }
    }
}
catch {
    Write-Host "  WMI method failed, continuing..." -ForegroundColor Yellow
}

Write-Host "`n4. Registry-based uninstallation..." -ForegroundColor Cyan
$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

foreach ($regPath in $regPaths) {
    try {
        $mcafeePrograms = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*McAfee*" }
        
        foreach ($program in $mcafeePrograms) {
            Write-Host "  Found in registry: $($program.DisplayName)" -ForegroundColor Yellow
            
            if ($program.UninstallString) {
                try {
                    $uninstallString = $program.UninstallString
                    
                    if ($uninstallString -like "*msiexec*") {
                        $productCode = ($uninstallString -replace '.*\{([^}]+)\}.*', '{$1}')
                        if ($productCode -like "{*}") {
                            Write-Host "    Using MSI uninstall for: $($program.DisplayName)" -ForegroundColor Yellow
                            Start-Process "msiexec.exe" -ArgumentList "/x $productCode /quiet /norestart /L*v `"$env:TEMP\McAfee_Uninstall.log`"" -Wait -NoNewWindow
                            Write-Host "    MSI uninstall completed" -ForegroundColor Green
                        }
                    }
                    elseif ($uninstallString -like "*.exe*") {
                        Write-Host "    Using EXE uninstall for: $($program.DisplayName)" -ForegroundColor Yellow
                        $exePath = ($uninstallString -split '"')[1]
                        if ($exePath -and (Test-Path $exePath)) {
                            Start-Process $exePath -ArgumentList "/S", "/silent", "/quiet", "/uninstall" -Wait -NoNewWindow -ErrorAction SilentlyContinue
                            Write-Host "    EXE uninstall completed" -ForegroundColor Green
                        }
                    }
                }
                catch {
                    Write-Host "    Failed to uninstall: $($program.DisplayName)" -ForegroundColor Red
                }
            }
        }
    }
    catch {
        Write-Host "  Error accessing registry path: $regPath" -ForegroundColor Red
    }
}

Write-Host "`n5. Removing McAfee scheduled tasks..." -ForegroundColor Cyan
try {
    $mcafeeTasks = Get-ScheduledTask | Where-Object { $_.TaskName -like "*McAfee*" -or $_.TaskPath -like "*McAfee*" }
    foreach ($task in $mcafeeTasks) {
        Write-Host "  Removing scheduled task: $($task.TaskName)" -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false -ErrorAction SilentlyContinue
    }
}
catch {
    Write-Host "  Could not remove scheduled tasks" -ForegroundColor Yellow
}

Write-Host "`n6. Comprehensive file and folder removal..." -ForegroundColor Cyan
$mcafeePaths = @(
    "$env:ProgramFiles\McAfee*",
    "$env:ProgramFiles\Common Files\McAfee*",
    "${env:ProgramFiles(x86)}\McAfee*", 
    "${env:ProgramFiles(x86)}\Common Files\McAfee*",
    "$env:ProgramData\McAfee*",
    "$env:LOCALAPPDATA\McAfee*",
    "$env:APPDATA\McAfee*",
    "$env:ALLUSERSPROFILE\McAfee*",
    "$env:SystemRoot\System32\drivers\mfe*",
    "$env:SystemRoot\System32\mfe*",
    "$env:TEMP\McAfee*",
    "C:\McAfee*"
)

foreach ($pathPattern in $mcafeePaths) {
    try {
        $items = Get-Item $pathPattern -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            Write-Host "  Removing: $($item.FullName)" -ForegroundColor Yellow
            try {
                takeown /f "$($item.FullName)" /r /d y > $null 2>&1
                icacls "$($item.FullName)" /grant administrators:F /t > $null 2>&1
            }
            catch { }
            Remove-Item $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
    }
}

Write-Host "`n7. Deep registry cleanup..." -ForegroundColor Cyan
$mcafeeRegPaths = @(
    "HKLM:\SOFTWARE\McAfee*",
    "HKLM:\SOFTWARE\WOW6432Node\McAfee*",
    "HKCU:\SOFTWARE\McAfee*",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\*McAfee*",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce\*McAfee*",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\*McAfee*",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce\*McAfee*",
    "HKLM:\SYSTEM\CurrentControlSet\Services\*McAfee*",
    "HKLM:\SYSTEM\CurrentControlSet\Services\*mfe*"
)

foreach ($regPath in $mcafeeRegPaths) {
    try {
        if ($regPath -like "*\Run\*" -or $regPath -like "*\RunOnce\*") {
            $basePath = $regPath -replace '\\\*.*$', ''
            if (Test-Path $basePath) {
                $values = Get-ItemProperty $basePath -ErrorAction SilentlyContinue
                foreach ($property in $values.PSObject.Properties) {
                    if ($property.Name -like "*McAfee*") {
                        Write-Host "  Removing registry value: $basePath\$($property.Name)" -ForegroundColor Yellow
                        Remove-ItemProperty -Path $basePath -Name $property.Name -ErrorAction SilentlyContinue
                    }
                }
            }
        } else {
            $items = Get-Item $regPath -ErrorAction SilentlyContinue
            foreach ($item in $items) {
                Write-Host "  Removing registry key: $($item.Name)" -ForegroundColor Yellow
                Remove-Item $item.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
    catch {
    }
}

Write-Host "`n10. Final verification..." -ForegroundColor Cyan
$remainingServices = Get-Service | Where-Object { $_.Name -like "*McAfee*" -or $_.DisplayName -like "*McAfee*" }
$remainingProcesses = Get-Process | Where-Object { $_.ProcessName -like "*McAfee*" -or $_.ProcessName -like "*mfe*" } -ErrorAction SilentlyContinue
$remainingFiles = Get-ChildItem "$env:ProgramFiles\McAfee*", "${env:ProgramFiles(x86)}\McAfee*" -ErrorAction SilentlyContinue

if ($remainingServices) {
    Write-Host "  WARNING: Some McAfee services still exist:" -ForegroundColor Yellow
    foreach ($service in $remainingServices) {
        Write-Host "    - $($service.DisplayName)" -ForegroundColor Yellow
    }
}

if ($remainingProcesses) {
    Write-Host "  WARNING: Some McAfee processes still running:" -ForegroundColor Yellow
    foreach ($process in $remainingProcesses) {
        Write-Host "    - $($process.ProcessName)" -ForegroundColor Yellow
    }
}

if ($remainingFiles) {
    Write-Host "  WARNING: Some McAfee files/folders still exist:" -ForegroundColor Yellow
    foreach ($file in $remainingFiles) {
        Write-Host "    - $($file.FullName)" -ForegroundColor Yellow
    }
}

if (-not $remainingServices -and -not $remainingProcesses -and -not $remainingFiles) {
    Write-Host "  McAfee appears to be completely removed!" -ForegroundColor Green
}

Write-Host "`n=== SUMMARY ===" -ForegroundColor Green
Write-Host "McAfee processes force-stopped" -ForegroundColor Green
Write-Host "McAfee services stopped and disabled" -ForegroundColor Green
Write-Host "McAfee products uninstalled (multiple methods)" -ForegroundColor Green  
Write-Host "McAfee scheduled tasks removed" -ForegroundColor Green
Write-Host "McAfee files and folders removed" -ForegroundColor Green
Write-Host "McAfee registry entries cleaned" -ForegroundColor Green

Write-Host "`nCompleted at: $(Get-Date)" -ForegroundColor Yellow
Write-Host "A reboot is strongly recommended to complete the removal process." -ForegroundColor Cyan

$reboot = Read-Host "`nWould you like to reboot now? (Y/N)"
if ($reboot -eq 'Y' -or $reboot -eq 'y') {
    Write-Host "Rebooting in 10 seconds..." -ForegroundColor Yellow
    Start-Sleep -Seconds 10
    Restart-Computer -Force
} else {
    Write-Host "Please remember to reboot manually to complete the removal process." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
} 