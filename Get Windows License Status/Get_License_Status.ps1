$LicenseStatus = (Get-WmiObject -Query "SELECT * FROM SoftwareLicensingProduct WHERE PartialProductKey IS NOT NULL AND Name LIKE 'Windows%'" -Namespace "ROOT\CIMV2").LicenseStatus

$StatusMap = @{
    0 = "Unlicensed"
    1 = "Licensed"
    2 = "OOBGrace"
    3 = "OOTGrace"
    4 = "NonGenuineGrace"
    5 = "Notification"
    6 = "ExtendedGrace"
}

$StatusMessage = $StatusMap[$LicenseStatus]

Write-Output "Windows License Status: $LicenseStatus ($StatusMessage)"

$LogPath = "C:\ProgramData\LicenseStatus.log"
"$(Get-Date) - Windows License Status: $LicenseStatus ($StatusMessage)" | Out-File -Append -FilePath $LogPath
