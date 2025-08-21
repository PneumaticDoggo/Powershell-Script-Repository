$report = [PSCustomObject]@{
    "Computer Name"       = $env:COMPUTERNAME
    "OS Version"          = (Get-CimInstance Win32_OperatingSystem).Caption
    "OS Build"            = (Get-CimInstance Win32_OperatingSystem).BuildNumber
    "CPU"                 = (Get-CimInstance Win32_Processor).Name
    "Cores"               = (Get-CimInstance Win32_Processor).NumberOfCores
    "Logical Processors"  = (Get-CimInstance Win32_Processor).NumberOfLogicalProcessors
    "Total RAM (GB)"      = "{0:N2}" -f ((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB)
    "BIOS Version"        = (Get-CimInstance Win32_BIOS).SMBIOSBIOSVersion
    "System Model"        = (Get-CimInstance Win32_ComputerSystem).Model
    "System Manufacturer" = (Get-CimInstance Win32_ComputerSystem).Manufacturer
}

$disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" |
         Select-Object DeviceID,
                       @{n="Size(GB)";e={"{0:N2}" -f ($_.Size/1GB)}},
                       @{n="Free(GB)";e={"{0:N2}" -f ($_.FreeSpace/1GB)}}

$adapters = Get-CimInstance Win32_NetworkAdapter | Where-Object { $_.PhysicalAdapter -eq $true }
$configs  = Get-CimInstance Win32_NetworkAdapterConfiguration

Write-Host "==== SYSTEM REPORT ====" -ForegroundColor Cyan
$report | Format-List

Write-Host "`n==== DISKS ====" -ForegroundColor Cyan
$disks | Format-Table -AutoSize

Write-Host "`n==== NETWORK ====" -ForegroundColor Cyan
Write-Host "Network Card(s): $($adapters.Count) NIC(s) Installed."

$index = 1
foreach ($nic in $adapters) {
    $cfg = $configs | Where-Object { $_.Index -eq $nic.DeviceID }

    Write-Host ("[{0:D2}]: {1}" -f $index, $nic.Name)

    try {
        $netAdapter = Get-NetAdapter -InterfaceDescription $nic.Name -ErrorAction SilentlyContinue
        if ($netAdapter) {
            Write-Host ("     Connection Name: {0}" -f $netAdapter.Name)
            Write-Host ("     Status:          {0}" -f $netAdapter.Status)
        }
    } catch {}

    if ($cfg -and $cfg.IPEnabled) {
        Write-Host ("     DHCP Enabled:    {0}" -f ($(if ($cfg.DHCPEnabled) {"Yes"} else {"No"})))
        Write-Host ("     DHCP Server:     {0}" -f ($(if ($cfg.DHCPServer) {$cfg.DHCPServer} else {"N/A"})))
        
        if ($cfg.IPAddress) {
            Write-Host "     IP address(es)"
            $i = 1
            foreach ($ip in $cfg.IPAddress) {
                Write-Host ("         [{0:D2}]: {1}" -f $i, $ip)
                $i++
            }
        }
    } else {
        Write-Host "     Status:          Media disconnected"
    }

    $index++
    Write-Host
}

$vmDiskInfo = Get-VM | ForEach-Object {
    $vm = $_
    Get-VMHardDiskDrive -VMName $vm.Name | ForEach-Object {
        $vhd = Get-VHD -Path $_.Path
        [PSCustomObject]@{
            VMName         = $vm.Name
            State          = $vm.State
            CPUCount       = $vm.ProcessorCount
            MemoryAssigned = "{0:N2}" -f ($vm.MemoryAssigned / 1GB)
            MemoryStartup  = "{0:N2}" -f ($vm.MemoryStartup / 1GB)
            Path           = $_.Path
            Controller     = $_.ControllerType
            ControllerNo   = $_.ControllerNumber
            DiskNumber     = $_.ControllerLocation
            DiskSizeGB     = "{0:N2}" -f ($vhd.Size / 1GB)
            DiskUsedGB     = "{0:N2}" -f ($vhd.FileSize / 1GB)
            Generation     = $vm.Generation
            Uptime         = $vm.Uptime
            CreationTime   = $vm.CreationTime
        }
    }
}

$vmDiskInfo | Format-Table -AutoSize

$csvPath = "C:\Temp\VM-DiskReport.csv"
$vmDiskInfo | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "VM + Disk report exported to $csvPath" -ForegroundColor Green
