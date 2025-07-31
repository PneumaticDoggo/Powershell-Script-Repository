function Ensure-MicrosoftGraphModule {
    $ModuleName = "Microsoft.Graph"

    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Output "The Microsoft Graph module is not installed. Installing now..."

        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -ErrorAction Stop
            Write-Output "Microsoft Graph module installed successfully."
        } catch {
            Write-Output "Failed to install Microsoft Graph module: $_"
            throw
        }
    } else {
        Write-Output "Microsoft Graph module is already installed."
    }
}

$savecsv = $PSScriptroot
Ensure-MicrosoftGraphModule

Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All", "DeviceManagementManagedDevices.Read.All"

function Get-AllDevices {
    Write-Output "Fetching all registered devices..."
    $devices = Get-MgDevice -All -Property "DeviceId,DisplayName,UserPrincipalName,OS,TrustType" | Select-Object DeviceId, DisplayName, UserPrincipalName, OS, TrustType
    Write-Output "Devices fetched: $($devices.Count)"
    return $devices
}

function Get-AllUsers {
    Write-Output "Fetching all users..."
    $users = Get-MgUser -All -Property "UserPrincipalName,AssignedLicenses,AccountEnabled,Mail" | Select-Object UserPrincipalName, AssignedLicenses, AccountEnabled, Mail
    Write-Output "Users fetched: $($users.Count)"
    return $users
}

function Get-LicenseDetails {
    Write-Output "Fetching license details..."
    $licensesPurchased = Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber
    Write-Output "Licenses fetched: $($licensesPurchased.Count)"
    return $licensesPurchased
}

$devices = Get-AllDevices
$users = Get-AllUsers
$licensesPurchased = Get-LicenseDetails

$licenseMap = @{}
foreach ($license in $licensesPurchased) {
    if ($license.SkuId -and $license.SkuPartNumber) {
        $licenseMap[$license.SkuId] = $license.SkuPartNumber
    }
}

$userDetails = @()
foreach ($user in $users) {
    $assignedLicenseNames = @()
    foreach ($license in $user.AssignedLicenses) {
        if ($licenseMap.ContainsKey($license.SkuId)) {
            $assignedLicenseNames += $licenseMap[$license.SkuId]
        } else {
            $assignedLicenseNames += "Unknown License ($($license.SkuId))"
        }
    }

    $userDetails += [PSCustomObject]@{
        UserPrincipalName = $user.UserPrincipalName
        Mail              = if (-not [string]::IsNullOrEmpty($user.Mail)) { $user.Mail } else { "No Mailbox" }
        IsEnabled         = if ($user.AccountEnabled) { "Enabled" } else { "Disabled" }
        LicensesAssigned  = $assignedLicenseNames -join ", "
    }
}

$deviceDetails = $devices | ForEach-Object {
    [PSCustomObject]@{
        DeviceId      = $_.DeviceId
        DeviceName    = $_.DisplayName
        Owner         = $_.UserPrincipalName
        OS            = $_.OS
        JoinType      = $_.TrustType
    }
}

Write-Output "`nUser Details:"
$userDetails | Format-Table -AutoSize

Write-Output "`nDevice Details:"
$deviceDetails | Format-Table -AutoSize

Write-Output "`nPurchased License Overview:"
$licensesPurchased | Format-Table SkuPartNumber, SkuId -AutoSize

$userDetails | Export-Csv -Path "$savecsv\UserDetails.csv" -NoTypeInformation
$deviceDetails | Export-Csv -Path "$savecsv\DeviceDetails.csv" -NoTypeInformation


Write-Output "Details exported to " Write-Output $savecsv


Disconnect-MgGraph