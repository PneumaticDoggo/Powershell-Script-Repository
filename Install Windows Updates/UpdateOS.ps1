[CmdletBinding()]
Param(
    [Parameter(Mandatory = $False)] [ValidateSet('Soft', 'Hard', 'None', 'Delayed')] [String] $Reboot = 'Soft',
    [Parameter(Mandatory = $False)] [Int32] $RebootTimeout = 120,
    [Parameter(Mandatory = $False)] [switch] $ExcludeDrivers,
    [Parameter(Mandatory = $False)] [switch] $ExcludeUpdates
)

Process {

    if ("$env:PROCESSOR_ARCHITEW6432" -ne "ARM64") {
        if (Test-Path "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe") {
            if ($ExcludeDrivers) {
                & "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy bypass -NoProfile -File "$PSCommandPath" -Reboot $Reboot -RebootTimeout $RebootTimeout -ExcludeDrivers
            } elseif ($ExcludeUpdates) {
                & "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy bypass -NoProfile -File "$PSCommandPath" -Reboot $Reboot -RebootTimeout $RebootTimeout -ExcludeUpdates
            } else {
                & "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy bypass -NoProfile -File "$PSCommandPath" -Reboot $Reboot -RebootTimeout $RebootTimeout
            }
            Exit $lastexitcode
        }
    }

    if (-not (Test-Path "$($env:ProgramData)\Microsoft\UpdateOS")) {
        Mkdir "$($env:ProgramData)\Microsoft\UpdateOS"
    }
    Set-Content -Path "$($env:ProgramData)\Microsoft\UpdateOS\UpdateOS.ps1.tag" -Value "Installed"

    Start-Transcript "$($env:ProgramData)\Microsoft\UpdateOS\UpdateOS.log"

    $script:needReboot = $false

    $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
    Write-Output "$ts Opting into Microsoft Update"
    $ServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager"
    $ServiceID = "7971f918-a847-4430-9279-4a52d1efe18d"
    $ServiceManager.AddService2($ServiceId, 7, "") | Out-Null

    $WUDownloader = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateDownloader()
    $WUInstaller = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateInstaller()
    if ($ExcludeDrivers) {
        $queries = @("IsInstalled=0 and Type='Software'")
    }
    elseif ($ExcludeUpdates) {
        $queries = @("IsInstalled=0 and Type='Driver'")
    } else {
        $queries = @("IsInstalled=0 and Type='Software'", "IsInstalled=0 and Type='Driver'")
    }

    $queries | ForEach-Object {

        $WUUpdates = New-Object -ComObject Microsoft.Update.UpdateColl
        $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
        Write-Host "$ts Getting $_ updates."        
        ((New-Object -ComObject Microsoft.Update.Session).CreateupdateSearcher().Search($_)).Updates | ForEach-Object {
            if (!$_.EulaAccepted) { $_.AcceptEula() }
            if ($_.Title -notmatch "Preview") { [void]$WUUpdates.Add($_) }
        }

        if ($WUUpdates.Count -ge 1) {
            $WUInstaller.ForceQuiet = $true
            $WUInstaller.Updates = $WUUpdates
            $WUDownloader.Updates = $WUUpdates
            $UpdateCount = $WUDownloader.Updates.count
            if ($UpdateCount -ge 1) {
                $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
                Write-Output "$ts Downloading $UpdateCount Updates"
                foreach ($update in $WUInstaller.Updates) { Write-Output "$($update.Title)" }
                $Download = $WUDownloader.Download()
            }
            $InstallUpdateCount = $WUInstaller.Updates.count
            if ($InstallUpdateCount -ge 1) {
                $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
                Write-Output "$ts Installing $InstallUpdateCount Updates"
                $Install = $WUInstaller.Install()
                $ResultMeaning = ($Results | Where-Object { $_.ResultCode -eq $Install.ResultCode }).Meaning
                Write-Output $ResultMeaning
                $script:needReboot = $Install.RebootRequired
            } 
        }
        else {
            Write-Output "No Updates Found"
        } 
    }

    $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
    if ($script:needReboot) {
        Write-Host "$ts Windows Update indicated that a reboot is needed."
    }
    else {
        Write-Host "$ts Windows Update indicated that no reboot is required."
    }

    $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
    if ($Reboot -eq "Hard") {
        Write-Host "$ts Exiting with return code 1641 to indicate a hard reboot is needed."
        Stop-Transcript
        Exit 1641
    }
    elseif ($Reboot -eq "Soft") {
        Write-Host "$ts Exiting with return code 3010 to indicate a soft reboot is needed."
        Stop-Transcript
        Exit 3010
    }
    elseif ($Reboot -eq "Delayed") {
        Write-Host "$ts Rebooting with a $RebootTimeout second delay"
        & shutdown.exe /r /t $RebootTimeout /c "Rebooting to complete the installation of Windows updates."
        Exit 0
    }
    else {
        Write-Host "$ts Skipping reboot based on Reboot parameter (None)"
        Exit 0
    }

}
