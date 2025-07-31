Import-Module Microsoft.Graph

Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All", "DeviceManagementManagedDevices.Read.All"

function Get-AllDevicesWithOwners {
    Write-Output "Fetching all registered devices..."
    $devices = Get-MgDevice -All -Property "DeviceId,DisplayName,DeviceOwnership,OperatingSystem,TrustType"
    $deviceDetails = @()

    foreach ($device in $devices) {
        $owners = Get-MgDeviceRegisteredOwner -DeviceId $device.DeviceId -All | ForEach-Object { $_.UserPrincipalName } -join ", "
        
        $deviceDetails += [PSCustomObject]@{
            DeviceId      = $device.DeviceId
            DeviceName    = $device.DisplayName
            Owner         = if ($owners) { $owners } else { "No Owners Assigned" }
            OS            = $device.OperatingSystem
            JoinType      = $device.TrustType
        }
    }

    Write-Output "Devices fetched: $($deviceDetails.Count)"
    return $deviceDetails
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

function Get-MailboxType {
    param (
        [string]$UserPrincipalName
    )
    try {
        $recipient = Get-MgUser -UserId $UserPrincipalName -Property "MailboxSettings"
        if ($recipient -and $recipient.MailboxSettings) {
            return "UserMailbox"  # Adjust this as per Graph data structure if additional data is available
        } else {
            return "SharedMailbox"
        }
    } catch {
        return "Unknown"  
    }
}

$devices = Get-AllDevicesWithOwners
$users = Get-AllUsers
$licensesPurchased = Get-LicenseDetails

$licenseMap = @{}
foreach ($license in $licensesPurchased) {
    if ($license.SkuId -and $license.SkuPartNumber) {
        $licenseMap[$license.SkuId] = $license.SkuPartNumber
    }
}

$userDetails = @()
$mailboxDetails = @()

foreach ($user in $users) {
    $assignedLicenseNames = @()
    foreach ($license in $user.AssignedLicenses) {
        if ($licenseMap.ContainsKey($license.SkuId)) {
            $assignedLicenseNames += $licenseMap[$license.SkuId]
        } else {
            $assignedLicenseNames += "Unknown License ($($license.SkuId))"
        }
    }

    $mailboxType = if (-not [string]::IsNullOrEmpty($user.Mail)) {
        Get-MailboxType -UserPrincipalName $user.UserPrincipalName
    } else {
        "No Mailbox"
    }

    $userDetails += [PSCustomObject]@{
        UserPrincipalName = $user.UserPrincipalName
        Mail              = if (-not [string]::IsNullOrEmpty($user.Mail)) { $user.Mail } else { "No Mailbox" }
        IsEnabled         = if ($user.AccountEnabled) { "Enabled" } else { "Disabled" }
        LicensesAssigned  = $assignedLicenseNames -join ", "
    }

    if (-not [string]::IsNullOrEmpty($user.Mail)) {
        $mailboxDetails += [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            Mailbox           = $user.Mail
            MailboxType       = $mailboxType
            Owner             = $user.UserPrincipalName
        }
    }
}

Write-Output "`nUser Details:"
$userDetails | Format-Table -AutoSize

Write-Output "`nDevice Details:"
$devices | Format-Table -AutoSize

Write-Output "`nPurchased License Overview:"
$licensesPurchased | Format-Table SkuPartNumber, SkuId -AutoSize

Write-Output "`nMailbox Details:"
$mailboxDetails | Format-Table -AutoSize

$userDetails | Export-Csv -Path "C:\UserDetails.csv" -NoTypeInformation
$devices | Export-Csv -Path "C:\DeviceDetails.csv" -NoTypeInformation
$mailboxDetails | Export-Csv -Path "C:\MailboxDetails.csv" -NoTypeInformation
Write-Output "Details exported to C:\UserDetails.csv, C:\DeviceDetails.csv, and C:\MailboxDetails.csv"
