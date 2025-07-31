# Requires Microsoft.Graph PowerShell SDK (Install-Module Microsoft.Graph)
# Run with an account or app that has:
#   Policy.Read.All, Policy.ReadWrite.AuthenticationMethod, Group.Read.All,
#   DeviceManagementConfiguration.Read.All, DeviceManagementApps.Read.All

# Connect to Microsoft Graph
# Ensure the registered app or account has consented to the following delegated permissions:
#   • Policy.Read.All (covers most policy reads)
#   • Policy.ReadWrite.AuthenticationMethod (required for auth method policy reads)
#   • Group.Read.All (resolve group names)
#   • DeviceManagementConfiguration.Read.All (Intune config & compliance)
#   • DeviceManagementApps.Read.All (if you extend to Intune apps)
Connect-MgGraph -Scopes Policy.Read.All,Policy.ReadWrite.AuthenticationMethod,Group.Read.All,DeviceManagementConfiguration.Read.All,DeviceManagementApps.Read.All

function Resolve-GroupNames {
    param([string[]] $GroupIds)
    if (-not $GroupIds) { return @() }
    $names = foreach ($gid in $GroupIds) {
        try { (Get-MgGroup -GroupId $gid).DisplayName } catch { $gid }
    }
    return $names
}

$caPolicies = Get-MgIdentityConditionalAccessPolicy -ErrorAction Stop
$caReport = $caPolicies | Select-Object -Property 
    @{n='Name';e={$_.DisplayName}},
    @{n='State';e={$_.State}},
    @{n='IncludedGroups';e={ (Resolve-GroupNames ($_.Conditions.Users.Include)) -join '; ' }}
$caReport | Export-Csv -Path './CA-Policies.csv' -NoTypeInformation

try {
    $authResp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy/authenticationMethodConfigurations'
    $authConfigs = $authResp.value
} catch {
    Write-Warning 'Cannot retrieve authentication method configurations; skipping section.'
    $authConfigs = @()
}
$authReport = foreach ($cfg in $authConfigs) {
    [PSCustomObject]@{
        MethodName     = $cfg.Id
        State          = $cfg.State
        IncludedGroups = (Resolve-GroupNames ($cfg.includeTargets | ForEach-Object { $_.Id })) -join '; '
    }
}
$authReport | Export-Csv -Path './Auth-Methods-Report.csv' -NoTypeInformation

$devConfs = Get-MgDeviceManagementDeviceConfiguration -ErrorAction Stop
$devConfReport = foreach ($cfg in $devConfs) {
    $assigns = Get-MgDeviceManagementDeviceConfigurationAssignment -DeviceConfigurationId $cfg.Id
    [PSCustomObject]@{
        PolicyName     = $cfg.DisplayName
        AssignedGroups = (Resolve-GroupNames ($assigns.Target | ForEach-Object { $_.TargetId })) -join '; '
    }
}
$devConfReport | Export-Csv -Path './Intune-DeviceConfig.csv' -NoTypeInformation

$comps = Get-MgDeviceManagementDeviceCompliancePolicy -ErrorAction Stop
$compReport = foreach ($c in $comps) {
    try {
        $assigns = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$($c.Id)/assignments"
        $groupIds = $assigns.value | ForEach-Object { $_.target.groupId }
    } catch {
        Write-Warning "Cannot retrieve assignments for compliance policy '$($c.DisplayName)'"
        $groupIds = @()
    }
    [PSCustomObject]@{
        CompliancePolicy = $c.DisplayName
        AssignedGroups   = (Resolve-GroupNames $groupIds) -join '; '
    }
}
$compReport | Export-Csv -Path './Intune-Compliance.csv' -NoTypeInformation

try {
    $secConfs = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/deviceManagement/securityBaselineConfigurations'
    $secItems = $secConfs.value
} catch {
    Write-Warning 'Security baseline configurations endpoint not available; skipping section.'
    $secItems = @()
}
$secReport = foreach ($p in $secItems) {
    try {
        $assigns = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/securityBaselineConfigurations/$($p.id)/assignments"
        $groupIds = $assigns.value | ForEach-Object { $_.target.groupId }
    } catch {
        Write-Warning "Cannot retrieve assignments for security baseline '$($p.displayName)'"
        $groupIds = @()
    }
    [PSCustomObject]@{
        BaselineName    = $p.displayName
        ProfileType     = $p.platform
        AssignedGroups  = (Resolve-GroupNames $groupIds) -join '; '
    }
}
$secReport | Export-Csv -Path './Security-Baselines.csv' -NoTypeInformation

Write-Host 'Export complete. CSV files saved in current directory.'
