[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force

Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted

if (-not (Get-Command Get-WindowsAutopilotInfo -ErrorAction SilentlyContinue)) {
    Install-Script -Name Get-WindowsAutopilotInfo -Force
}

Get-WindowsAutopilotInfo -Online `
    -TenantId "<TenantId>" `
    -AppId "<AppId>" `
    -AppSecret "<AppSecret>" `
    -GroupTag "<GroupTag>"
# Add a grouptag if needed or remove the line if not needed
