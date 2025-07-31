function Get-MsiPath {
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "MSI Files (*.msi)|*.msi"
    $OpenFileDialog.Title = "Select the MSI file"
    $OpenFileDialog.ShowDialog() | Out-Null
    return $OpenFileDialog.FileName
}

$msiPath = Get-MsiPath

if (-not (Test-Path $msiPath)) {
    Write-Host "MSI file not found. Exiting..."
    exit 1
}

$adminCreds = Get-Credential -Message "Enter Administrator Credentials for Elevation"

Start-Process msiexec.exe -Credential $adminCreds -ArgumentList "/i", "`"$msiPath`"", "/qn", "/norestart" -NoNewWindow

Write-Host "MSI installation started with elevated privileges."
