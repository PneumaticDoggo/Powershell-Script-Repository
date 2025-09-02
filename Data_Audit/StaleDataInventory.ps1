[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$OutputFile = "StaleDataReport.csv"
)

function Format-FileSize {
    param ([int64]$Size)
    $sizes = 'Bytes,KB,MB,GB,TB'
    $sizes = $sizes.Split(',')
    $index = 0
    while ($Size -ge 1KB -and $index -lt ($sizes.Count - 1)) {
        $Size = $Size / 1KB
        $index++
    }
    "{0:N2} {1}" -f $Size, $sizes[$index]
}

function Convert-ToBytes {
    param (
        [string]$Size
    )
    
    $size = $size.Trim().ToUpper()
    if ($size -match '^\d+$') { return [int64]$size } 
    
    $value = [int64]($size -replace '[^0-9.]', '')
    $unit = $size -replace '[0-9.]', ''
    
    switch ($unit.ToUpper()) {
        'KB' { return $value * 1KB }
        'MB' { return $value * 1MB }
        'GB' { return $value * 1GB }
        'TB' { return $value * 1TB }
        default { return $value }
    }
}

function Show-Menu {
    param (
        [string]$Title = 'Select an option',
        [array]$Options,
        [switch]$MultiSelect
    )
    
    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host
    
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Write-Host "$($i+1)) $($Options[$i])"
    }
    Write-Host
    
    if ($MultiSelect) {
        Write-Host "Enter multiple numbers separated by commas (e.g., 1,3,4)"
        Write-Host "Enter 'A' to select all drives"
        $selection = Read-Host "Please make your selection"
        
        if ($selection -eq 'A') {
            return $Options
        }
        
        $selectedIndices = $selection -split ',' | ForEach-Object { $_.Trim() }
        return $selectedIndices | ForEach-Object { $Options[$_ - 1] }
    }
    else {
        $selection = Read-Host "Please make a selection (1-$($Options.Count))"
        return $Options[$selection - 1]
    }
}

function Get-AvailableDrives {
    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Free -gt 0 }
    return $drives | ForEach-Object {
        $freeSpace = Format-FileSize -Size $_.Free
        $usedSpace = Format-FileSize -Size ($_.Used)
        $totalSpace = Format-FileSize -Size ($_.Free + $_.Used)
        "$($_.Name): ($($_.Description)) - Free: $freeSpace, Used: $usedSpace, Total: $totalSpace"
    }
}

function Get-ScanParameters {
    $params = @{}
    
    Write-Host "`nSelect last access time threshold:"
    $timeOptions = @(
        "30 days",
        "60 days",
        "90 days",
        "180 days",
        "365 days",
        "Custom"
    )
    $timeChoice = Show-Menu -Title "Select Time Threshold" -Options $timeOptions
    
    if ($timeChoice -eq "Custom") {
        $days = Read-Host "Enter number of days"
        $params.LastAccessDays = [int]$days
    }
    else {
        $params.LastAccessDays = [int]($timeChoice -replace " days","")
    }
    
    Write-Host "`nSelect minimum file size to consider:"
    $sizeOptions = @(
        "1 MB",
        "10 MB",
        "100 MB",
        "1 GB",
        "Custom"
    )
    $sizeChoice = Show-Menu -Title "Select Minimum File Size" -Options $sizeOptions
    
    if ($sizeChoice -eq "Custom") {
        Write-Host "Enter size (e.g., '500MB' or '2GB'):"
        $customSize = Read-Host
        $params.MinSize = Convert-ToBytes -Size $customSize
    }
    else {
        $params.MinSize = Convert-ToBytes -Size $sizeChoice
    }
    
    return $params
}

function Remove-StaleData {
    param (
        [Parameter(Mandatory = $true)]
        [string]$CsvPath
    )

    if (-not (Test-Path $CsvPath)) {
        Write-Error "CSV file not found: $CsvPath"
        return
    }

    $filesToDelete = Import-Csv -Path $CsvPath

    $totalFiles = $filesToDelete.Count
    $totalSize = ($filesToDelete | Measure-Object -Property SizeInBytes -Sum).Sum
    $totalSizeReadable = Format-FileSize -Size $totalSize

    Write-Host "`nPreparing to delete files:"
    Write-Host "Total files to delete: $totalFiles"
    Write-Host "Total size to be freed: $totalSizeReadable"

    $confirmation = Read-Host "`nAre you sure you want to delete these files? (Y/N)"
    if ($confirmation -ne 'Y') {
        Write-Host "Operation cancelled by user."
        return
    }

    $deletedFiles = 0
    $failedFiles = 0
    $progress = 0

    foreach ($file in $filesToDelete) {
        $progress++
        $percentComplete = ($progress / $totalFiles) * 100

        Write-Progress -Activity "Deleting Stale Files" `
            -Status "Processing $($file.FullPath)" `
            -PercentComplete $percentComplete

        if (Test-Path $file.FullPath) {
            try {
                Remove-Item -Path $file.FullPath -Force
                $deletedFiles++
            }
            catch {
                Write-Warning "Failed to delete: $($file.FullPath)"
                Write-Warning "Error: $_"
                $failedFiles++
            }
        }
        else {
            Write-Warning "File not found: $($file.FullPath)"
            $failedFiles++
        }
    }

    Write-Progress -Activity "Deleting Stale Files" -Completed

    Write-Host "`nDeletion Summary:"
    Write-Host "Successfully deleted: $deletedFiles files"
    Write-Host "Failed to delete: $failedFiles files"
    Write-Host "Total space freed: $totalSizeReadable"
}

function Start-StaleDataScan {
    $driveOptions = Get-AvailableDrives
    $selectedDrives = Show-Menu -Title "Select Drives to Scan" -Options $driveOptions -MultiSelect
    $drivePaths = $selectedDrives | ForEach-Object { "$($_.Substring(0, 1)):\" }
    
    $params = Get-ScanParameters
    
    Write-Host "`nStarting stale data analysis..."
    Write-Host "Selected drives: $($drivePaths -join ', ')"
    Write-Host "Looking for files not accessed in the last $($params.LastAccessDays) days"
    Write-Host "Minimum file size: $(Format-FileSize -Size $params.MinSize)"
    
    $cutOffDate = (Get-Date).AddDays(-$params.LastAccessDays)
    
    $allFiles = @()
    $driveNumber = 1
    
    foreach ($drivePath in $drivePaths) {
        $driveFiles = Start-DriveScan -DrivePath $drivePath `
            -CutOffDate $cutOffDate `
            -MinSize $params.MinSize `
            -TotalDrives $drivePaths.Count `
            -CurrentDriveNumber $driveNumber
        
        $allFiles += $driveFiles
        $driveNumber++
    }
    
    Write-Progress -Activity "Scanning Drives" -Completed
    
    if ($allFiles) {
        $allFiles | Export-Csv -Path $OutputFile -NoTypeInformation
        
        Write-Host "`nAnalysis Complete!"
        
        $drivePaths | ForEach-Object {
            $currentDrive = $_
            $driveFiles = $allFiles | Where-Object { $_.Drive -eq $currentDrive }
            if ($driveFiles) {
                $driveTotal = $driveFiles.Count
                $driveSize = ($driveFiles | Measure-Object -Property SizeInBytes -Sum).Sum
                $driveSizeReadable = Format-FileSize -Size $driveSize
                
                Write-Host "`nDrive $(Split-Path $currentDrive -Qualifier):"
                Write-Host "  Found $driveTotal potentially stale files"
                Write-Host "  Total size of stale files: $driveSizeReadable"
            }
        }
        
        $grandTotal = $allFiles.Count
        $grandSize = ($allFiles | Measure-Object -Property SizeInBytes -Sum).Sum
        $grandSizeReadable = Format-FileSize -Size $grandSize
        
        Write-Host "`nGrand Total:"
        Write-Host "Total stale files across all drives: $grandTotal"
        Write-Host "Total size across all drives: $grandSizeReadable"
        Write-Host "Results have been exported to: $OutputFile"
        
        $openCsv = Read-Host "`nWould you like to open the CSV file? (Y/N)"
        if ($openCsv -eq 'Y') {
            Invoke-Item $OutputFile
        }
    }
    else {
        Write-Host "No stale files found matching the specified criteria."
    }
}

function Start-FolderScan {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )

    if (-not (Test-Path -Path $FolderPath -PathType Container)) {
        Write-Error "Folder not found: $FolderPath"
        return
    }

    $params = Get-ScanParameters
    
    Write-Host "`nStarting stale data analysis..."
    Write-Host "Selected folder: $FolderPath"
    Write-Host "Looking for files not accessed in the last $($params.LastAccessDays) days"
    Write-Host "Minimum file size: $(Format-FileSize -Size $params.MinSize)"
    
    $cutOffDate = (Get-Date).AddDays(-$params.LastAccessDays)
    
    $files = Start-DriveScan -DrivePath $FolderPath `
        -CutOffDate $cutOffDate `
        -MinSize $params.MinSize `
        -TotalDrives 1 `
        -CurrentDriveNumber 1
    
    Write-Progress -Activity "Scanning Folder" -Completed
    
    if ($files) {
        $files | Export-Csv -Path $OutputFile -NoTypeInformation
        
        Write-Host "`nAnalysis Complete!"
        
        $totalFiles = $files.Count
        $totalSize = ($files | Measure-Object -Property SizeInBytes -Sum).Sum
        $totalSizeReadable = Format-FileSize -Size $totalSize
        
        Write-Host "`nSummary:"
        Write-Host "Found $totalFiles potentially stale files"
        Write-Host "Total size of stale files: $totalSizeReadable"
        Write-Host "Results have been exported to: $OutputFile"
        
        $openCsv = Read-Host "`nWould you like to open the CSV file? (Y/N)"
        if ($openCsv -eq 'Y') {
            Invoke-Item $OutputFile
        }
    }
    else {
        Write-Host "No stale files found matching the specified criteria."
    }
}

function Start-DriveScan {
    param (
        [string]$DrivePath,
        [datetime]$CutOffDate,
        [int64]$MinSize,
        [int]$TotalDrives,
        [int]$CurrentDriveNumber
    )
    
    Write-Host "`nScanning path: $DrivePath"
    
    $progress = 0
    $totalFiles = (Get-ChildItem -Path $DrivePath -Recurse -File -ErrorAction SilentlyContinue).Count
    
    $files = Get-ChildItem -Path $DrivePath -Recurse -File -ErrorAction SilentlyContinue | 
        ForEach-Object {
            $progress++
            $overallProgress = (($CurrentDriveNumber - 1) / $TotalDrives * 100) + ($progress / $totalFiles * (100 / $TotalDrives))
            
            Write-Progress -Activity "Scanning Files" `
                -Status "$($_.FullName)" `
                -PercentComplete $overallProgress
            
            if ($_.LastAccessTime -lt $CutOffDate -and $_.Length -ge $MinSize) {
                $_ | Select-Object @{
                    Name = 'Drive'
                    Expression = { $DrivePath }
                },
                @{
                    Name = 'FullPath'
                    Expression = { $_.FullName }
                },
                @{
                    Name = 'SizeInBytes'
                    Expression = { $_.Length }
                },
                @{
                    Name = 'SizeReadable'
                    Expression = { Format-FileSize -Size $_.Length }
                },
                @{
                    Name = 'LastModified'
                    Expression = { $_.LastWriteTime }
                },
                @{
                    Name = 'LastAccessed'
                    Expression = { $_.LastAccessTime }
                },
                @{
                    Name = 'DaysSinceLastAccess'
                    Expression = { [math]::Round((New-TimeSpan -Start $_.LastAccessTime -End (Get-Date)).TotalDays) }
                },
                @{
                    Name = 'FileExtension'
                    Expression = { $_.Extension }
                }
            }
        }
    
    return $files
}

try {
    $mainMenuOptions = @(
        "Scan drives for stale data",
        "Scan specific folder for stale data",
        "Delete files using CSV report",
        "Exit"
    )

    while ($true) {
        $choice = Show-Menu -Title "Data Inventory Tool" -Options $mainMenuOptions

        switch ($choice) {
            "Scan drives for stale data" {
                Start-StaleDataScan
            }
            "Scan specific folder for stale data" {
                $folderPath = Read-Host "`nEnter the folder path to scan"
                if (Test-Path -Path $folderPath -PathType Container) {
                    Start-FolderScan -FolderPath $folderPath
                } else {
                    Write-Error "Invalid folder path: $folderPath"
                }
            }
            "Delete files using CSV report" {
                $csvPath = Read-Host "`nEnter the path to the CSV file (press Enter for $OutputFile)"
                if ([string]::IsNullOrWhiteSpace($csvPath)) {
                    $csvPath = $OutputFile
                }
                Remove-StaleData -CsvPath $csvPath
            }
            "Exit" {
                Write-Host "`nExiting..."
                exit 0
            }
        }

        Write-Host "`nPress Enter to return to main menu..."
        Read-Host
    }
}
catch {
    Write-Error "An error occurred during execution: $_"
    exit 1
}

function Format-FileSize {
    param ([int64]$Size)
    $sizes = 'Bytes,KB,MB,GB,TB'
    $sizes = $sizes.Split(',')
    $index = 0
    while ($Size -ge 1KB -and $index -lt ($sizes.Count - 1)) {
        $Size = $Size / 1KB
        $index++
    }
    "{0:N2} {1}" -f $Size, $sizes[$index]
}

function Convert-ToBytes {
    param (
        [string]$Size
    )
    
    $size = $size.Trim().ToUpper()
    if ($size -match '^\d+$') { return [int64]$size }
    
    $value = [int64]($size -replace '[^0-9.]', '')
    $unit = $size -replace '[0-9.]', ''
    
    switch ($unit.ToUpper()) {
        'KB' { return $value * 1KB }
        'MB' { return $value * 1MB }
        'GB' { return $value * 1GB }
        'TB' { return $value * 1TB }
        default { return $value }
    }
}

function Show-Menu {
    param (
        [string]$Title = 'Select an option',
        [array]$Options,
        [switch]$MultiSelect
    )
    
    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host
    
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Write-Host "$($i+1)) $($Options[$i])"
    }
    Write-Host
    
    if ($MultiSelect) {
        Write-Host "Enter multiple numbers separated by commas (e.g., 1,3,4)"
        Write-Host "Enter 'A' to select all drives"
        $selection = Read-Host "Please make your selection"
        
        if ($selection -eq 'A') {
            return $Options
        }
        
        $selectedIndices = $selection -split ',' | ForEach-Object { $_.Trim() }
        return $selectedIndices | ForEach-Object { $Options[$_ - 1] }
    }
    else {
        $selection = Read-Host "Please make a selection (1-$($Options.Count))"
        return $Options[$selection - 1]
    }
}

function Get-AvailableDrives {
    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Free -gt 0 }
    return $drives | ForEach-Object {
        $freeSpace = Format-FileSize -Size $_.Free
        $usedSpace = Format-FileSize -Size ($_.Used)
        $totalSpace = Format-FileSize -Size ($_.Free + $_.Used)
        "$($_.Name): ($($_.Description)) - Free: $freeSpace, Used: $usedSpace, Total: $totalSpace"
    }
}

function Get-ScanParameters {
    $params = @{}
    
    Write-Host "`nSelect last access time threshold:"
    $timeOptions = @(
        "30 days",
        "60 days",
        "90 days",
        "180 days",
        "365 days",
        "Custom"
    )
    $timeChoice = Show-Menu -Title "Select Time Threshold" -Options $timeOptions
    
    if ($timeChoice -eq "Custom") {
        $days = Read-Host "Enter number of days"
        $params.LastAccessDays = [int]$days
    }
    else {
        $params.LastAccessDays = [int]($timeChoice -replace " days","")
    }
    
    Write-Host "`nSelect minimum file size to consider:"
    $sizeOptions = @(
        "1 MB",
        "10 MB",
        "100 MB",
        "1 GB",
        "Custom"
    )
    $sizeChoice = Show-Menu -Title "Select Minimum File Size" -Options $sizeOptions
    
    if ($sizeChoice -eq "Custom") {
        Write-Host "Enter size (e.g., '500MB' or '2GB'):"
        $customSize = Read-Host
        $params.MinSize = Convert-ToBytes -Size $customSize
    }
    else {
        $params.MinSize = Convert-ToBytes -Size $sizeChoice
    }
    
    return $params
}

function Remove-StaleData {
    param (
        [Parameter(Mandatory = $true)]
        [string]$CsvPath
    )

    if (-not (Test-Path $CsvPath)) {
        Write-Error "CSV file not found: $CsvPath"
        return
    }

    $filesToDelete = Import-Csv -Path $CsvPath

    $totalFiles = $filesToDelete.Count
    $totalSize = ($filesToDelete | Measure-Object -Property SizeInBytes -Sum).Sum
    $totalSizeReadable = Format-FileSize -Size $totalSize

    Write-Host "`nPreparing to delete files:"
    Write-Host "Total files to delete: $totalFiles"
    Write-Host "Total size to be freed: $totalSizeReadable"

    $confirmation = Read-Host "`nAre you sure you want to delete these files? (Y/N)"
    if ($confirmation -ne 'Y') {
        Write-Host "Operation cancelled by user."
        return
    }

    $deletedFiles = 0
    $failedFiles = 0
    $progress = 0

    foreach ($file in $filesToDelete) {
        $progress++
        $percentComplete = ($progress / $totalFiles) * 100

        Write-Progress -Activity "Deleting Stale Files" `
            -Status "Processing $($file.FullPath)" `
            -PercentComplete $percentComplete

        if (Test-Path $file.FullPath) {
            try {
                Remove-Item -Path $file.FullPath -Force
                $deletedFiles++
            }
            catch {
                Write-Warning "Failed to delete: $($file.FullPath)"
                Write-Warning "Error: $_"
                $failedFiles++
            }
        }
        else {
            Write-Warning "File not found: $($file.FullPath)"
            $failedFiles++
        }
    }

    Write-Progress -Activity "Deleting Stale Files" -Completed

    Write-Host "`nDeletion Summary:"
    Write-Host "Successfully deleted: $deletedFiles files"
    Write-Host "Failed to delete: $failedFiles files"
    Write-Host "Total space freed: $totalSizeReadable"
}

function Start-DriveScan {
    param (
        [string]$DrivePath,
        [datetime]$CutOffDate,
        [int64]$MinSize,
        [int]$TotalDrives,
        [int]$CurrentDriveNumber
    )
    
    Write-Host "`nScanning drive: $DrivePath"
    
    # Initialize progress counter
    $progress = 0
    
    # Determine if we need to exclude Windows directory
    if ($DrivePath -eq "C:\") {
        Write-Host "Excluding Windows directory from scan"
        
        # Get total files, excluding Windows folder
        $totalFiles = @(Get-ChildItem -Path $DrivePath -Exclude "Windows" -Recurse -File -ErrorAction SilentlyContinue).Count
        
        # Get all files, excluding Windows folder
        $files = Get-ChildItem -Path $DrivePath -Exclude "Windows" -Recurse -File -ErrorAction SilentlyContinue
    } else {
        # Get total files for non-C: drives
        $totalFiles = @(Get-ChildItem -Path $DrivePath -Recurse -File -ErrorAction SilentlyContinue).Count
        
        # Get all files for non-C: drives
        $files = Get-ChildItem -Path $DrivePath -Recurse -File -ErrorAction SilentlyContinue
    }
    
    # Process the files
    $files = $files | ForEach-Object {
            $progress++
            $overallProgress = (($CurrentDriveNumber - 1) / $TotalDrives * 100) + ($progress / $totalFiles * (100 / $TotalDrives))
            
            Write-Progress -Activity "Scanning Drives" `
                -Status "Drive $CurrentDriveNumber of $TotalDrives - $($_.FullName)" `
                -PercentComplete $overallProgress
            
            if ($_.LastAccessTime -lt $CutOffDate -and $_.Length -ge $MinSize) {
                $_ | Select-Object @{
                    Name = 'Drive'
                    Expression = { $DrivePath }
                },
                @{
                    Name = 'FullPath'
                    Expression = { $_.FullName }
                },
                @{
                    Name = 'SizeInBytes'
                    Expression = { $_.Length }
                },
                @{
                    Name = 'SizeReadable'
                    Expression = { Format-FileSize -Size $_.Length }
                },
                @{
                    Name = 'LastModified'
                    Expression = { $_.LastWriteTime }
                },
                @{
                    Name = 'LastAccessed'
                    Expression = { $_.LastAccessTime }
                },
                @{
                    Name = 'DaysSinceLastAccess'
                    Expression = { [math]::Round((New-TimeSpan -Start $_.LastAccessTime -End (Get-Date)).TotalDays) }
                },
                @{
                    Name = 'FileExtension'
                    Expression = { $_.Extension }
                }
            }
        }
    
    return $files
}

function Start-StaleDataScan {
    $driveOptions = Get-AvailableDrives
    $selectedDrives = Show-Menu -Title "Select Drives to Scan" -Options $driveOptions -MultiSelect
    $drivePaths = $selectedDrives | ForEach-Object { "$($_.Substring(0, 1)):\" }
    
    $params = Get-ScanParameters
    
    Write-Host "`nStarting stale data analysis..."
    Write-Host "Selected drives: $($drivePaths -join ', ')"
    Write-Host "Looking for files not accessed in the last $($params.LastAccessDays) days"
    Write-Host "Minimum file size: $(Format-FileSize -Size $params.MinSize)"
    
    $cutOffDate = (Get-Date).AddDays(-$params.LastAccessDays)
    
    $allFiles = @()
    $driveNumber = 1
    
    foreach ($drivePath in $drivePaths) {
        $driveFiles = Start-DriveScan -DrivePath $drivePath `
            -CutOffDate $cutOffDate `
            -MinSize $params.MinSize `
            -TotalDrives $drivePaths.Count `
            -CurrentDriveNumber $driveNumber
        
        $allFiles += $driveFiles
        $driveNumber++
    }
    
    Write-Progress -Activity "Scanning Drives" -Completed
    
    if ($allFiles) {
        $allFiles | Export-Csv -Path $OutputFile -NoTypeInformation
        
        Write-Host "`nAnalysis Complete!"
        
        $drivePaths | ForEach-Object {
            $currentDrive = $_
            $driveFiles = $allFiles | Where-Object { $_.Drive -eq $currentDrive }
            if ($driveFiles) {
                $driveTotal = $driveFiles.Count
                $driveSize = ($driveFiles | Measure-Object -Property SizeInBytes -Sum).Sum
                $driveSizeReadable = Format-FileSize -Size $driveSize
                
                Write-Host "`nDrive $(Split-Path $currentDrive -Qualifier):"
                Write-Host "  Found $driveTotal potentially stale files"
                Write-Host "  Total size of stale files: $driveSizeReadable"
            }
        }
        
        $grandTotal = $allFiles.Count
        $grandSize = ($allFiles | Measure-Object -Property SizeInBytes -Sum).Sum
        $grandSizeReadable = Format-FileSize -Size $grandSize
        
        Write-Host "`nGrand Total:"
        Write-Host "Total stale files across all drives: $grandTotal"
        Write-Host "Total size across all drives: $grandSizeReadable"
        Write-Host "Results have been exported to: $OutputFile"
        
        $openCsv = Read-Host "`nWould you like to open the CSV file? (Y/N)"
        if ($openCsv -eq 'Y') {
            Invoke-Item $OutputFile
        }
    }
    else {
        Write-Host "No stale files found matching the specified criteria."
    }
}
catch {
    Write-Error "An error occurred during execution: $_"
    exit 1
}