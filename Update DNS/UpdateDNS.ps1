$newDNSServer = "ENTER DNS SERVER IP"

$adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }

foreach ($adapter in $adapters) {
    Set-DnsClientServerAddress -InterfaceAlias $adapter.Name -ServerAddresses $newDNSServer
    Write-Output "Updated DNS for adapter: $($adapter.Name) to $newDNSServer"
}
