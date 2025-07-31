function Test-RequiredModules {
    $modules = @()
    
    if (Get-Module -ListAvailable -Name SmbShare) {
        Import-Module SmbShare -ErrorAction SilentlyContinue
        $modules += "SmbShare"
    } else {
        Write-Warning "SmbShare module not available. Share permissions will be limited."
    }
    
    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            $modules += "ActiveDirectory"
            Write-Host "Active Directory module loaded successfully." -ForegroundColor Green
        } catch {
            Write-Warning "Active Directory module available but failed to load: $_"
        }
    } else {
        Write-Warning "Active Directory module not available. SID resolution will be limited."
    }
    
    return $modules
}

function Resolve-Identity {
    param(
        [string]$Identity,
        [bool]$UseAD = $false
    )
    
    if ($UseAD -and $Identity -match '^S-1-') {
        try {
            $user = Get-ADObject -Filter "SID -eq '$Identity'" -Properties Name, SamAccountName, ObjectClass
            if ($user) {
                return "$($user.ObjectClass): $($user.Name) ($($user.SamAccountName))"
            }
        } catch {
        }
    }
    
    return $Identity
}

function Get-SharePermissions {
    param (
        [string]$Path,
        [bool]$UseAD = $false,
        [bool]$IncludeSharePerms = $false
    )
    
    $results = @()
    
    try {
        $acl = Get-Acl -Path $Path -ErrorAction Stop
        foreach ($access in $acl.Access) {
            $resolvedIdentity = Resolve-Identity -Identity $access.IdentityReference.ToString() -UseAD $UseAD
            
            $results += [PSCustomObject]@{
                FolderPath     = $Path
                Identity       = $resolvedIdentity
                OriginalSID    = $access.IdentityReference.ToString()
                Rights         = $access.FileSystemRights
                AccessControl  = $access.AccessControlType
                Inherited      = $access.IsInherited
                Type           = "NTFS"
                ShareName      = ""
            }
        }
        
        if ($IncludeSharePerms -and (Get-Command Get-SmbShare -ErrorAction SilentlyContinue)) {
            $shares = Get-SmbShare | Where-Object { $_.Path -eq $Path }
            foreach ($share in $shares) {
                try {
                    $shareAccess = Get-SmbShareAccess -Name $share.Name -ErrorAction SilentlyContinue
                    foreach ($access in $shareAccess) {
                        $resolvedIdentity = Resolve-Identity -Identity $access.AccountName -UseAD $UseAD
                        
                        $results += [PSCustomObject]@{
                            FolderPath     = $Path
                            Identity       = $resolvedIdentity
                            OriginalSID    = $access.AccountName
                            Rights         = $access.AccessRight
                            AccessControl  = $access.AccessControlType
                            Inherited      = $false
                            Type           = "SMB Share"
                            ShareName      = $share.Name
                        }
                    }
                } catch {
                    Write-Warning "Failed to get share permissions for $($share.Name): $_"
                }
            }
        }
        
    } catch {
        Write-Warning "Failed to read permissions for $Path : $_"
    }
    
    return $results
}

function Get-AllShares {
    $shares = @()
    
    if (Get-Command Get-SmbShare -ErrorAction SilentlyContinue) {
        try {
            $smbShares = Get-SmbShare | Where-Object { $_.ShareType -eq 'FileSystemDirectory' -and $_.Name -notmatch '^[A-Z]\$$' }
            foreach ($share in $smbShares) {
                $shares += [PSCustomObject]@{
                    Name = $share.Name
                    Path = $share.Path
                    Description = $share.Description
                    Type = "SMB Share"
                }
            }
        } catch {
            Write-Warning "Failed to enumerate SMB shares: $_"
        }
    }
    
    try {
        $adminShares = Get-SmbShare | Where-Object { $_.Name -match '^[A-Z]\$$' }
        foreach ($share in $adminShares) {
            $shares += [PSCustomObject]@{
                Name = $share.Name
                Path = $share.Path
                Description = "Administrative Share"
                Type = "Admin Share"
            }
        }
    } catch {
        Write-Warning "Failed to enumerate administrative shares: $_"
    }
    
    return $shares
}

function Show-Menu {
    param (
        [string]$Title,
        [array]$Options
    )
    
    Write-Host "`n=== $Title ===" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Write-Host "[$($i + 1)] $($Options[$i])" -ForegroundColor Yellow
    }
    Write-Host "[0] Back/Exit" -ForegroundColor Red
    Write-Host ""
}

function Select-AuditType {
    Write-Host "`n=== Select Audit Type ===" -ForegroundColor Cyan
    Write-Host "[1] Audit specific drives/folders (manual selection)" -ForegroundColor Yellow
    Write-Host "[2] Audit all SMB shares on this server" -ForegroundColor Yellow
    Write-Host "[3] Audit both drives and shares (comprehensive)" -ForegroundColor Yellow
    Write-Host "[0] Exit" -ForegroundColor Red
    
    do {
        $choice = Read-Host "`nEnter your choice (0-3)"
        if ($choice -match '^[0-3]$') {
            return [int]$choice
        } else {
            Write-Host "Invalid choice. Please enter 0, 1, 2, or 3." -ForegroundColor Red
        }
    } while ($true)
}

function Select-Drives {
    $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | Sort-Object DeviceID
    $selectedDrives = @()
    
    if ($drives.Count -eq 0) {
        Write-Host "No local drives found." -ForegroundColor Red
        return $selectedDrives
    }
    
    Write-Host "`n=== Available Local Drives ===" -ForegroundColor Cyan
    for ($i = 0; $i -lt $drives.Count; $i++) {
        $drive = $drives[$i]
        $sizeGB = if ($drive.Size) { '{0:N2}' -f ($drive.Size / 1GB) } else { "Unknown" }
        Write-Host "[$($i + 1)] $($drive.DeviceID) - $($drive.VolumeName) ($sizeGB GB)" -ForegroundColor Yellow
    }
    Write-Host "[A] Select All" -ForegroundColor Green
    Write-Host "[0] Continue without selecting drives" -ForegroundColor Red
    
    Write-Host "`nEnter drive numbers separated by commas (e.g., 1,3,5) or 'A' for all:" -ForegroundColor White
    $input = Read-Host
    
    if ($input.ToUpper() -eq 'A') {
        $selectedDrives = $drives.DeviceID
        Write-Host "Selected all drives: $($selectedDrives -join ', ')" -ForegroundColor Green
    } elseif ($input -ne '0' -and $input.Trim() -ne '') {
        $choices = $input -split ',' | ForEach-Object { $_.Trim() }
        foreach ($choice in $choices) {
            if ($choice -match '^\d+$') {
                $index = [int]$choice - 1
                if ($index -ge 0 -and $index -lt $drives.Count) {
                    $selectedDrives += $drives[$index].DeviceID
                } else {
                    Write-Host "Invalid choice: $choice (out of range)" -ForegroundColor Red
                }
            } else {
                Write-Host "Invalid choice: $choice (not a number)" -ForegroundColor Red
            }
        }
        if ($selectedDrives.Count -gt 0) {
            Write-Host "Selected drives: $($selectedDrives -join ', ')" -ForegroundColor Green
        }
    }
    
    return $selectedDrives
}

function Select-Shares {
    $shares = Get-AllShares
    $selectedShares = @()
    
    if ($shares.Count -eq 0) {
        Write-Host "No shares found on this server." -ForegroundColor Red
        return $selectedShares
    }
    
    Write-Host "`n=== Available Shares ===" -ForegroundColor Cyan
    for ($i = 0; $i -lt $shares.Count; $i++) {
        $share = $shares[$i]
        Write-Host "[$($i + 1)] $($share.Name) -> $($share.Path) [$($share.Type)]" -ForegroundColor Yellow
        if ($share.Description) {
            Write-Host "     Description: $($share.Description)" -ForegroundColor Gray
        }
    }
    Write-Host "[A] Select All" -ForegroundColor Green
    Write-Host "[0] Continue without selecting shares" -ForegroundColor Red
    
    Write-Host "`nEnter share numbers separated by commas (e.g., 1,3,5) or 'A' for all:" -ForegroundColor White
    $input = Read-Host
    
    if ($input.ToUpper() -eq 'A') {
        $selectedShares = $shares.Path
        Write-Host "Selected all shares: $($shares.Count) total" -ForegroundColor Green
    } elseif ($input -ne '0' -and $input.Trim() -ne '') {
        $choices = $input -split ',' | ForEach-Object { $_.Trim() }
        foreach ($choice in $choices) {
            if ($choice -match '^\d+$') {
                $index = [int]$choice - 1
                if ($index -ge 0 -and $index -lt $shares.Count) {
                    $selectedShares += $shares[$index].Path
                } else {
                    Write-Host "Invalid choice: $choice (out of range)" -ForegroundColor Red
                }
            } else {
                Write-Host "Invalid choice: $choice (not a number)" -ForegroundColor Red
            }
        }
        if ($selectedShares.Count -gt 0) {
            Write-Host "Selected $($selectedShares.Count) shares" -ForegroundColor Green
        }
    }
    
    return $selectedShares
}

function Add-CustomPaths {
    $customPaths = @()
    
    Write-Host "`n=== Add Custom Paths ===" -ForegroundColor Cyan
    Write-Host "Enter custom paths (network shares, specific folders, etc.)" -ForegroundColor White
    Write-Host "Examples: \\server\share, C:\SpecificFolder, D:\Data" -ForegroundColor Gray
    Write-Host "Press Enter with empty input when done." -ForegroundColor Gray
    
    do {
        $path = Read-Host "`nEnter path (or press Enter to finish)"
        if ($path.Trim() -ne '') {
            if (Test-Path -Path $path) {
                $customPaths += $path.Trim()
                Write-Host "Added: $path" -ForegroundColor Green
            } else {
                $confirm = Read-Host "Path '$path' is not accessible. Add anyway? (y/N)"
                if ($confirm.ToLower() -eq 'y') {
                    $customPaths += $path.Trim()
                    Write-Host "Added: $path (may not be accessible)" -ForegroundColor Yellow
                }
            }
        }
    } while ($path.Trim() -ne '')
    
    if ($customPaths.Count -gt 0) {
        Write-Host "`nCustom paths added:" -ForegroundColor Green
        $customPaths | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    }
    
    return $customPaths
}

function Get-OutputPath {
    $defaultPath = Join-Path $env:USERPROFILE "Desktop\FileSharePermissionsReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    Write-Host "`n=== Output File Location ===" -ForegroundColor Cyan
    Write-Host "Default location: $defaultPath" -ForegroundColor Gray
    
    $customPath = Read-Host "Enter custom output path (or press Enter for default)"
    
    if ($customPath.Trim() -eq '') {
        $outputPath = $defaultPath
    } else {
        $outputPath = $customPath.Trim()
    }
    
    $directory = Split-Path $outputPath -Parent
    if (!(Test-Path $directory)) {
        try {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
            Write-Host "Created directory: $directory" -ForegroundColor Green
        } catch {
            Write-Host "Failed to create directory: $directory" -ForegroundColor Red
            Write-Host "Using default location instead." -ForegroundColor Yellow
            $outputPath = $defaultPath
        }
    }
    
    Write-Host "Output will be saved to: $outputPath" -ForegroundColor Green
    return $outputPath
}

function Start-PermissionsAudit {
    Write-Host @"

╔══════════════════════════════════════════════════════════════╗
║           Enhanced Folder Permissions Audit Tool            ║
║              Domain Controller Compatible                   ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Cyan

    Write-Host "Checking required modules..." -ForegroundColor Yellow
    $availableModules = Test-RequiredModules
    $useAD = $availableModules -contains "ActiveDirectory"
    $useShares = $availableModules -contains "SmbShare"
    
    Write-Host "Available modules: $($availableModules -join ', ')" -ForegroundColor Green
    if ($useAD) { Write-Host "✓ Active Directory integration enabled" -ForegroundColor Green }
    if ($useShares) { Write-Host "✓ SMB Share permissions enabled" -ForegroundColor Green }
    
    $auditType = Select-AuditType
    
    if ($auditType -eq 0) {
        Write-Host "Exiting..." -ForegroundColor Yellow
        return
    }
    
    $selectedPaths = @()
    
    switch ($auditType) {
        1 { 
            $selectedPaths += Select-Drives
            $selectedPaths += Add-CustomPaths
        }
        2 { 
            if ($useShares) {
                $selectedPaths += Select-Shares
            } else {
                Write-Host "SMB Share module not available. Cannot audit shares." -ForegroundColor Red
                return
            }
        }
        3 { 
            $selectedPaths += Select-Drives
            if ($useShares) {
                $selectedPaths += Select-Shares
            }
            $selectedPaths += Add-CustomPaths
        }
    }
    
    if ($selectedPaths.Count -eq 0) {
        Write-Host "`nNo paths selected. Exiting..." -ForegroundColor Red
        return
    }
    
    $outputPath = Get-OutputPath
    
    Write-Host "`n=== Audit Configuration ===" -ForegroundColor Cyan
    Write-Host "Audit Type: $(switch($auditType) { 1 {'Manual Selection'} 2 {'All Shares'} 3 {'Comprehensive'} })" -ForegroundColor White
    Write-Host "Paths to audit ($($selectedPaths.Count)):" -ForegroundColor White
    $selectedPaths | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    Write-Host "Output file: $outputPath" -ForegroundColor White
    Write-Host "Active Directory: $(if($useAD) {'Enabled'} else {'Disabled'})" -ForegroundColor White
    Write-Host "Share Permissions: $(if($useShares) {'Enabled'} else {'Disabled'})" -ForegroundColor White
    
    $confirm = Read-Host "`nProceed with audit? (Y/n)"
    if ($confirm.ToLower() -eq 'n') {
        Write-Host "Audit cancelled." -ForegroundColor Yellow
        return
    }
    
    Write-Host "`n=== Starting Enhanced Audit ===" -ForegroundColor Cyan
    $results = @()
    $totalPaths = $selectedPaths.Count
    $currentPath = 0
    
    foreach ($share in $selectedPaths) {
        $currentPath++
        $percentComplete = [math]::Round(($currentPath / $totalPaths) * 100)
        Write-Host "[$percentComplete%] Auditing: $share ($currentPath of $totalPaths)" -ForegroundColor Green
        
        if (Test-Path -Path $share) {
            try {
                Write-Host "  Discovering folders..." -ForegroundColor Gray
                $folders = @()
                $folders += Get-Item -Path $share -ErrorAction SilentlyContinue
                $subFolders = Get-ChildItem -Path $share -Recurse -Directory -ErrorAction SilentlyContinue
                $folders += $subFolders
                
                Write-Host "  Found $($folders.Count) folders to audit" -ForegroundColor Gray
                
                $folderCount = 0
                foreach ($folder in $folders) {
                    $folderCount++
                    if ($folderCount % 25 -eq 0) {
                        Write-Host "    Processed $folderCount/$($folders.Count) folders..." -ForegroundColor Gray
                    }
                    
                    $permissions = Get-SharePermissions -Path $folder.FullName -UseAD $useAD -IncludeSharePerms $useShares
                    $results += $permissions
                }
                Write-Host "  ✓ Completed: $folderCount folders, $($results.Count) permissions found so far" -ForegroundColor Green
            } catch {
                Write-Warning "Error processing $share : $_"
            }
        } else {
            Write-Warning "Path not accessible: $share"
        }
    }
    
    Write-Host "`nSaving results to file..." -ForegroundColor Cyan
    try {
        $results | Sort-Object FolderPath, Identity | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
        
        Write-Host @"

╔══════════════════════════════════════════════════════════════╗
║                     Audit Complete!                         ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Green

        Write-Host "Report saved to: $outputPath" -ForegroundColor Green
        Write-Host "Total permissions found: $($results.Count)" -ForegroundColor Yellow
        Write-Host "Paths audited: $($selectedPaths.Count)" -ForegroundColor Yellow
        Write-Host "NTFS permissions: $(($results | Where-Object {$_.Type -eq 'NTFS'}).Count)" -ForegroundColor Yellow
        if ($useShares) {
            Write-Host "Share permissions: $(($results | Where-Object {$_.Type -eq 'SMB Share'}).Count)" -ForegroundColor Yellow
        }
        
        if ($results.Count -gt 0) {
            Write-Host "`nSample results:" -ForegroundColor Cyan
            $results | Select-Object -First 5 | Format-Table -AutoSize
        }
        
    } catch {
        Write-Host "Error saving results: $_" -ForegroundColor Red
    }
}

Start-PermissionsAudit
