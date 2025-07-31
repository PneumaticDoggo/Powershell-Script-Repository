function Get-JoinStatus {
    $dsreg = dsregcmd /status
    return @{
        DomainJoined = ($dsreg -match "DomainJoined\s*:\s*YES")
        AzureAdJoined = ($dsreg -match "AzureAdJoined\s*:\s*YES")
        AzureAdRegistered = ($dsreg -match "AzureAdRegistered\s*:\s*YES")
    }
}

$status = Get-JoinStatus

if ($status.AzureAdRegistered -and -not $status.AzureAdJoined) {
    Write-Host "Device is Entra Registered. Proceeding with fix..."
    
    Write-Host "Removing Entra Registration..."
    dsregcmd /leave
    
    Write-Host "Cleaning up old Entra registration..."
    Remove-Item -Path "C:\ProgramData\Microsoft\Crypto\RSA\S-1-12-1-*" -Force -Recurse -ErrorAction SilentlyContinue
    Remove-Item -Path "C:\Windows\ServiceProfiles\LocalService\AppData\Local\Microsoft\Ngc" -Force -Recurse -ErrorAction SilentlyContinue
    
    Write-Host "Restarting the device in 10 seconds..."
    Start-Sleep -Seconds 10
    Restart-Computer -Force
}


Start-Sleep -Seconds 60

Write-Host "Attempting Hybrid Join..."
dsregcmd /join

Write-Host "Running the Hybrid Join Scheduled Task..."
Start-ScheduledTask -TaskPath "\Microsoft\Windows\Workplace Join\" -TaskName "Automatic-Device-Join"

Start-Sleep -Seconds 30
$status = Get-JoinStatus

if ($status.AzureAdJoined -and $status.DomainJoined) {
    Write-Host "Device is now Hybrid Azure AD Joined!"
} else {
    Write-Host "Hybrid Join failed. Check event logs under Microsoft > Windows > User Device Registration > Admin"
}
