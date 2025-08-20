#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Authentication

param(
    [Parameter(Mandatory = $false, HelpMessage = "Check assignments for specific users")]
    [switch]$CheckUser,
    
    [Parameter(Mandatory = $false, HelpMessage = "User Principal Names to check, comma-separated")]
    [string]$UserPrincipalNames,
    
    [Parameter(Mandatory = $false, HelpMessage = "Check assignments for specific groups")]
    [switch]$CheckGroup,
    
    [Parameter(Mandatory = $false, HelpMessage = "Group names or IDs to check, comma-separated")]
    [string]$GroupNames,
    
    [Parameter(Mandatory = $false, HelpMessage = "Check assignments for specific devices")]
    [switch]$CheckDevice,
    
    [Parameter(Mandatory = $false, HelpMessage = "Device names to check, comma-separated")]
    [string]$DeviceNames,
    
    [Parameter(Mandatory = $false, HelpMessage = "Show all policies and their assignments")]
    [switch]$ShowAllPolicies,
    
    [Parameter(Mandatory = $false, HelpMessage = "Show all 'All Users' assignments")]
    [switch]$ShowAllUsersAssignments,
    
    [Parameter(Mandatory = $false, HelpMessage = "Show all 'All Devices' assignments")]
    [switch]$ShowAllDevicesAssignments,
    
    [Parameter(Mandatory = $false, HelpMessage = "Generate HTML report")]
    [switch]$GenerateHTMLReport,
    
    [Parameter(Mandatory = $false, HelpMessage = "Show policies without assignments")]
    [switch]$ShowPoliciesWithoutAssignments,
    
    [Parameter(Mandatory = $false, HelpMessage = "Check for empty groups in assignments")]
    [switch]$CheckEmptyGroups,
    
    [Parameter(Mandatory = $false, HelpMessage = "Show all Administrative Templates")]
    [switch]$ShowAdminTemplates,
    
    [Parameter(Mandatory = $false, HelpMessage = "Show all failed assignments")]
    [switch]$ShowFailedAssignments,
    
    [Parameter(Mandatory = $false, HelpMessage = "Compare assignments between groups")]
    [switch]$CompareGroups,
    
    [Parameter(Mandatory = $false, HelpMessage = "Groups to compare assignments between, comma-separated")]
    [string]$CompareGroupNames,
    
    [Parameter(Mandatory = $false, HelpMessage = "Export results to CSV")]
    [switch]$ExportToCSV,
    
    [Parameter(Mandatory = $false, HelpMessage = "Path for the exported CSV file")]
    [string]$ExportPath,
    
    [Parameter(Mandatory = $false, HelpMessage = "App ID for authentication")]
    [string]$AppId,
    
    [Parameter(Mandatory = $false, HelpMessage = "Tenant ID for authentication")]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false, HelpMessage = "Certificate Thumbprint for authentication")]
    [string]$CertificateThumbprint,
    
    [Parameter(Mandatory = $false, HelpMessage = "Environment (Global, USGov, USGovDoD)")]
    [ValidateSet("Global", "USGov", "USGovDoD")]
    [string]$Environment = "Global"
)

$parameterMode = $false
$selectedOption = $null

if ($CheckUser) { $parameterMode = $true; $selectedOption = '1' }
elseif ($CheckGroup) { $parameterMode = $true; $selectedOption = '2' }
elseif ($CheckDevice) { $parameterMode = $true; $selectedOption = '3' }
elseif ($ShowAllPolicies) { $parameterMode = $true; $selectedOption = '4' }
elseif ($ShowAllUsersAssignments) { $parameterMode = $true; $selectedOption = '5' }
elseif ($ShowAllDevicesAssignments) { $parameterMode = $true; $selectedOption = '6' }
elseif ($GenerateHTMLReport) { $parameterMode = $true; $selectedOption = '7' }
elseif ($ShowPoliciesWithoutAssignments) { $parameterMode = $true; $selectedOption = '8' }
elseif ($CheckEmptyGroups) { $parameterMode = $true; $selectedOption = '9' }
elseif ($CompareGroups) { $parameterMode = $true; $selectedOption = '10' }
elseif ($ShowFailedAssignments) { $parameterMode = $true; $selectedOption = '11' }

<#
.REQUIRED PERMISSIONS
    - User.Read.All                    (Read user profiles)
    - Group.Read.All                   (Read group information)
    - Device.Read.All                  (Read device information)
    - DeviceManagementApps.Read.All    (Read app management data)
    - DeviceManagementConfiguration.Read.All    (Read device configurations)
    - DeviceManagementManagedDevices.Read.All   (Read device management data)
#>

################################ Prerequisites #####################################################

# Fill in your App ID, Tenant ID, and Certificate Thumbprint
# Use parameter values if provided, otherwise use defaults
$appid = if ($AppId) { $AppId } else { '<YourAppIdHere>' } # App ID of the App Registration
$tenantid = if ($TenantId) { $TenantId } else { '<YourTenantIdHere>' } # Tenant ID of your EntraID
$certThumbprint = if ($CertificateThumbprint) { $CertificateThumbprint } else { '<YourCertificateThumbprintHere>' } # Thumbprint of the certificate associated with the App Registration
# $certName = '<YourCertificateNameHere>' # Name of the certificate associated with the App Registration

####################################################################################################

$localVersion = "3.4.1"

Write-Host "365 Auditing Tool" -ForegroundColor Cyan
Write-Host " | Version" -NoNewline; Write-Host " $localVersion" -ForegroundColor Yellow -NoNewline
Write-Host ""
Write-Host ""

####################################################################################################

if ($PSScriptRoot) {
    $newScriptPath = Join-Path $PSScriptRoot "365Audit.ps1"
}
else {
    $currentDirectory = Get-Location
    $newScriptPath = Join-Path $currentDirectory "365Audit.ps1"
}

$autoUpdate = $false 

try {
    $latestVersion = Invoke-RestMethod -Uri $versionUrl
    
    $local = [System.Version]::new($localVersion)
    $latest = [System.Version]::new($latestVersion)
    
    if ($local -lt $latest) {
        Write-Host "A new version is available: $latestVersion (you are running $localVersion)" -ForegroundColor Yellow
        if ($autoUpdate) {
            Write-Host "AutoUpdate is enabled. Downloading the latest version..." -ForegroundColor Yellow
            try {
                Invoke-WebRequest -Uri $scriptUrl -OutFile $newScriptPath
                Write-Host "The latest version has been downloaded to $newScriptPath" -ForegroundColor Yellow
                Write-Host "Please restart the script to use the updated version." -ForegroundColor Yellow
            }
            catch {
            }
        }
        else {
            Write-Host "Auto-update is disabled." -ForegroundColor Yellow
            Write-Host "" 
        }
    }
    elseif ($local -gt $latest) {
        Write-Host "Note: You are running a pre-release version ($localVersion)" -ForegroundColor Magenta
        Write-Host ""
    }
}
catch {
    Write-Host "Unable to check for updates. Continue with current version..." -ForegroundColor Gray
}

####################################################################################################


$script:GraphEndpoint = $null
$script:GraphEnvironment = $null

function Set-Environment {
    param (
        [Parameter(Mandatory = $false)]
        [string]$EnvironmentName
    )
    
    if ($EnvironmentName) {
        switch ($EnvironmentName) {
            'Global' {
                $script:GraphEndpoint = "https://graph.microsoft.com"
                $script:GraphEnvironment = "Global"
                Write-Host "Environment set to Global" -ForegroundColor Green
                return $script:GraphEnvironment
            }
            'USGov' {
                $script:GraphEndpoint = "https://graph.microsoft.us"
                $script:GraphEnvironment = "USGov"
                Write-Host "Environment set to USGov" -ForegroundColor Green
                return $script:GraphEnvironment
            }
            'USGovDoD' {
                $script:GraphEndpoint = "https://dod-graph.microsoft.us"
                $script:GraphEnvironment = "USGovDoD"
                Write-Host "Environment set to USGovDoD" -ForegroundColor Green
                return $script:GraphEnvironment
            }
            default {
                Write-Host "Invalid environment name. Using interactive selection." -ForegroundColor Yellow
            }
        }
    }
    
    do {
        Write-Host "Select Intune Tenant Environment:" -ForegroundColor Cyan
        Write-Host "  [1] Global" -ForegroundColor White
        Write-Host "  [2] USGov" -ForegroundColor White
        Write-Host "  [3] USGovDoD" -ForegroundColor White
        Write-Host ""
        Write-Host "  [0] Exit" -ForegroundColor White
        Write-Host ""
        Write-Host "Select an option: " -ForegroundColor Yellow -NoNewline

        $selection = Read-Host

        switch ($selection) {
            '1' {
                $script:GraphEndpoint = "https://graph.microsoft.com"
                $script:GraphEnvironment = "Global"
                Write-Host "Environment set to Global" -ForegroundColor Green
                return $script:GraphEnvironment
            }
            '2' {
                $script:GraphEndpoint = "https://graph.microsoft.us"
                $script:GraphEnvironment = "USGov"
                Write-Host "Environment set to USGov" -ForegroundColor Green
                return $script:GraphEnvironment
            }
            '3' {
                $script:GraphEndpoint = "https://dod-graph.microsoft.us"
                $script:GraphEnvironment = "USGovDoD"
                Write-Host "Environment set to USGovDoD" -ForegroundColor Green
                return $script:GraphEnvironment
            }
            '0' {
                exit
            }
            default {
                Write-Host "Invalid choice, please select 1,2,3, or 0" -ForegroundColor Red
            }
        }
    } while ($selection -ne '0')
}


try {
    $requiredPermissions = @(
        @{
            Permission = "User.Read.All"
            Reason     = "Required to read user profile information and check group memberships"
        },
        @{
            Permission = "Group.Read.All"
            Reason     = "Needed to read group information and memberships"
        },
        @{
            Permission = "DeviceManagementConfiguration.Read.All"
            Reason     = "Allows reading Intune device configuration policies and their assignments"
        },
        @{
            Permission = "DeviceManagementApps.Read.All"
            Reason     = "Necessary to read mobile app management policies and app configurations"
        },
        @{
            Permission = "DeviceManagementManagedDevices.Read.All"
            Reason     = "Required to read managed device information and compliance policies"
        },
        @{
            Permission = "Device.Read.All"
            Reason     = "Needed to read device information from Entra ID"
        }
    )

    if (-not $appid -or $appid -eq '<YourAppIdHere>' -or
        -not $tenantid -or $tenantid -eq '<YourTenantIdHere>' -or
        -not $certThumbprint -or $certThumbprint -eq '<YourCertificateThumbprintHere>') {
        Write-Host "App ID, Tenant ID, or Certificate Thumbprint is missing or not set correctly." -ForegroundColor Red
        $manualConnection = Read-Host "Would you like to attempt a manual interactive connection? (y/n)"
        if ($manualConnection -eq 'y') {
            Write-Host "Attempting manual interactive connection (you need privileges to consent permissions)..." -ForegroundColor Yellow
            $permissionsList = ($requiredPermissions | ForEach-Object { $_.Permission }) -join ', '
            if ($parameterMode) {
                Set-Environment -EnvironmentName $Environment
            }
            else {
                Set-Environment  # Prompt for environment selection in interactive mode
            }
            $connectionResult = Connect-MgGraph -Scopes $permissionsList -Environment $script:GraphEnvironment -NoWelcome -ErrorAction Stop
        }
        else {
            Write-Host "Script execution cancelled by user." -ForegroundColor Red
            exit
        }
    }
    else {
        if ($parameterMode) {
            Set-Environment -EnvironmentName $Environment
        }
        else {
            Set-Environment  # Prompt for environment selection in interactive mode
        }
        $connectionResult = Connect-MgGraph -ClientId $appid -TenantId $tenantid -Environment $script:GraphEnvironment -CertificateThumbprint $certThumbprint -NoWelcome -ErrorAction Stop
    }
    Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green

    $context = Get-MgContext
    $currentPermissions = $context.Scopes

    Write-Host "Checking required permissions:" -ForegroundColor Cyan
    $missingPermissions = @()
    foreach ($permissionInfo in $requiredPermissions) {
        $permission = $permissionInfo.Permission
        $reason = $permissionInfo.Reason

        $hasPermission = $currentPermissions -contains $permission -or $currentPermissions -contains $permission.Replace(".Read", ".ReadWrite")

        if ($hasPermission) {
            Write-Host "  [✓] $permission" -ForegroundColor Green
            Write-Host "      Reason: $reason" -ForegroundColor Gray
        }
        else {
            Write-Host "  [✗] $permission" -ForegroundColor Red
            Write-Host "      Reason: $reason" -ForegroundColor Gray
            $missingPermissions += $permission
        }
    }

    if ($missingPermissions.Count -eq 0) {
        Write-Host "All required permissions are present." -ForegroundColor Green
        Write-Host ""
    }
    else {
        Write-Host "WARNING: The following permissions are missing:" -ForegroundColor Red
        $missingPermissions | ForEach-Object { 
            $missingPermission = $_
            $reason = ($requiredPermissions | Where-Object { $_.Permission -eq $missingPermission }).Reason
            Write-Host "  - $missingPermission" -ForegroundColor Yellow
            Write-Host "    Reason: $reason" -ForegroundColor Gray
        }
        Write-Host "The script will continue, but it may not function correctly without these permissions." -ForegroundColor Red
        Write-Host "Please ensure these permissions are granted to the app registration for full functionality." -ForegroundColor Yellow
        
        $continueChoice = Read-Host "Do you want to continue anyway? (y/n)"
        if ($continueChoice -ne 'y') {
            Write-Host "Script execution cancelled by user." -ForegroundColor Red
            exit
        }
    }
}
catch {
    Write-Host "Failed to connect to Microsoft Graph. Error: $_" -ForegroundColor Red
    
    if ($_.Exception.Message -like "*Certificate with thumbprint*was not found*") {
        Write-Host "The specified certificate was not found or has expired. Please check your certificate configuration." -ForegroundColor Yellow
    }
    
    exit
}

function Get-IntuneAssignments {
    param (
        [Parameter(Mandatory = $true)]
        [string]$EntityType,
        
        [Parameter(Mandatory = $true)]
        [string]$EntityId,
        
        [Parameter(Mandatory = $false)]
        [string]$GroupId = $null
    )

    $actualAssignmentsUri = $null

    if ($EntityType -eq "deviceAppManagement/managedAppPolicies") {
        $policyDetailsUri = "$GraphEndpoint/beta/deviceAppManagement/managedAppPolicies/$EntityId"
        try {
            $policyDetailsResponse = Invoke-MgGraphRequest -Uri $policyDetailsUri -Method Get
            $policyODataType = $policyDetailsResponse.'@odata.type'
            $specificPolicyTypePath = switch ($policyODataType) {
                "#microsoft.graph.androidManagedAppProtection" { "androidManagedAppProtections" }
                "#microsoft.graph.iosManagedAppProtection" { "iosManagedAppProtections" }
                "#microsoft.graph.windowsManagedAppProtection" { "windowsManagedAppProtections" }
                default { $null }
            }
            if ($specificPolicyTypePath) {
                $actualAssignmentsUri = "$GraphEndpoint/beta/deviceAppManagement/$specificPolicyTypePath('$EntityId')/assignments"
            }
            else {
                Write-Warning "Could not determine specific App Protection Policy type for $EntityId from OData type '$policyODataType'."
                return [System.Collections.ArrayList]::new() # Return empty ArrayList
            }
        }
        catch {
            Write-Warning "Error fetching details for App Protection Policy '$EntityId': $($_.Exception.Message)"
            return [System.Collections.ArrayList]::new() # Return empty ArrayList
        }
    }
    elseif ($EntityType -eq "mobileAppConfigurations") {
        $actualAssignmentsUri = "$GraphEndpoint/beta/deviceAppManagement/mobileAppConfigurations('$EntityId')/assignments"
    }
    elseif ($EntityType -like "deviceAppManagement/*ManagedAppProtections") {
        $actualAssignmentsUri = "$GraphEndpoint/beta/$EntityType('$EntityId')/assignments" # EntityType already includes deviceAppManagement
    }
    else {
        # General device management entities
        $actualAssignmentsUri = "$GraphEndpoint/beta/deviceManagement/$EntityType('$EntityId')/assignments"
    }

    if (-not $actualAssignmentsUri) {
        Write-Warning "Could not determine a valid assignments URI for EntityType '$EntityType' and EntityId '$EntityId'."
        return [System.Collections.ArrayList]::new() # Return empty ArrayList
    }

    $assignmentsToReturn = [System.Collections.ArrayList]::new()
    try {
        $allAssignmentsForEntity = [System.Collections.ArrayList]::new()
        $currentAssignmentsPageUri = $actualAssignmentsUri
        do {
            $pagedAssignmentResponse = Invoke-MgGraphRequest -Uri $currentAssignmentsPageUri -Method Get
            if ($pagedAssignmentResponse -and $null -ne $pagedAssignmentResponse.value) {
                $allAssignmentsForEntity.AddRange($pagedAssignmentResponse.value)
            }
            $currentAssignmentsPageUri = $pagedAssignmentResponse.'@odata.nextLink'
        } while (![string]::IsNullOrEmpty($currentAssignmentsPageUri))

        $assignmentList = if ($allAssignmentsForEntity) { $allAssignmentsForEntity } else { @() }

        foreach ($assignment in $assignmentList) {
            $currentAssignmentReason = $null
            $currentTargetGroupId = $null # Initialize to null

            if ($assignment.target -and $assignment.target.'@odata.type') {
                $odataType = $assignment.target.'@odata.type'
                
                if ($odataType -eq '#microsoft.graph.groupAssignmentTarget') {
                    $currentTargetGroupId = $assignment.target.groupId
                    if ($GroupId) {
                        if ($currentTargetGroupId -eq $GroupId) {
                            $currentAssignmentReason = "Direct Assignment"
                        }
                    }
                    else {
                        $currentAssignmentReason = "Group Assignment"
                    }
                }
                elseif ($odataType -eq '#microsoft.graph.exclusionGroupAssignmentTarget') {
                    $currentTargetGroupId = $assignment.target.groupId
                    if ($GroupId) {
                        if ($currentTargetGroupId -eq $GroupId) {
                            $currentAssignmentReason = "Direct Exclusion"
                        }
                    }
                    else {
                        $currentAssignmentReason = "Group Exclusion"
                    }
                }
                elseif (-not $GroupId) {
                    $currentAssignmentReason = switch ($odataType) {
                        '#microsoft.graph.allLicensedUsersAssignmentTarget' { "All Users" }
                        '#microsoft.graph.allDevicesAssignmentTarget' { "All Devices" }
                        default { $null }
                    }
                }
            }
            else {
                Write-Warning "Assignment item for EntityId '$EntityId' (URI: $actualAssignmentsUri) is missing 'target' or 'target.@odata.type' property. Assignment data: $($assignment | ConvertTo-Json -Depth 3)"
            }
            
            if ($currentAssignmentReason) {
                $null = $assignmentsToReturn.Add(@{
                        Reason  = $currentAssignmentReason
                        GroupId = $currentTargetGroupId
                        Apps    = $null # 'Apps' property is not directly available from general assignments endpoint
                    })
            }
        }
    }
    catch {
        Write-Warning "Error fetching assignments from '$actualAssignmentsUri': $($_.Exception.Message)"
    }
    
    return $assignmentsToReturn
}

function Get-IntuneEntities {
    param (
        [Parameter(Mandatory = $true)]
        [string]$EntityType,
        
        [Parameter(Mandatory = $false)]
        [string]$Filter = "",
        
        [Parameter(Mandatory = $false)]
        [string]$Select = "",
        
        [Parameter(Mandatory = $false)]
        [string]$Expand = ""
    )

    if ($EntityType -like "deviceAppManagement/*" -or $EntityType -eq "deviceManagement/templates" -or $EntityType -eq "deviceManagement/intents") {
        $baseUri = "$GraphEndpoint/beta"
        $actualEntityType = $EntityType
    }
    else {
        $baseUri = "$GraphEndpoint/beta/deviceManagement"
        $actualEntityType = "$EntityType"
    }
    
    $currentUri = "$baseUri/$actualEntityType"
    if ($Filter) { $currentUri += "?`$filter=$Filter" }
    if ($Select) { $currentUri += $(if ($Filter) { "&" }else { "?" }) + "`$select=$Select" }
    if ($Expand) { $currentUri += $(if ($Filter -or $Select) { "&" }else { "?" }) + "`$expand=$Expand" }

    $entities = [System.Collections.ArrayList]::new() # Initialize as ArrayList

    do {
        try {
            $response = Invoke-MgGraphRequest -Uri $currentUri -Method Get -ErrorAction Stop
            if ($null -ne $response -and $null -ne $response.value) {
                if ($response.value -is [array]) {
                    $entities.AddRange($response.value)
                }
                else {
                    $entities.Add($response.value)
                }
            }
            $currentUri = $response.'@odata.nextLink'
        }
        catch {
            Write-Warning "Error fetching entities for $EntityType from $currentUri : $($_.Exception.Message)"
            $currentUri = $null # Stop pagination on error
        }
    } while ($currentUri)

    return $entities
}

function Get-GroupInfo {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupId
    )

    try {
        $groupUri = "$GraphEndpoint/v1.0/groups/$GroupId"
        $group = Invoke-MgGraphRequest -Uri $groupUri -Method Get
        return @{
            Id          = $group.id
            DisplayName = $group.displayName
            Success     = $true
        }
    }
    catch {
        return @{
            Id          = $GroupId
            DisplayName = "Unknown Group"
            Success     = $false
        }
    }
}

function Get-DeviceInfo {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DeviceName
    )

    $deviceUri = "$GraphEndpoint/v1.0/devices?`$filter=displayName eq '$DeviceName'"
    $deviceResponse = Invoke-MgGraphRequest -Uri $deviceUri -Method Get
    
    if ($deviceResponse.value) {
        return @{
            Id          = $deviceResponse.value[0].id
            DisplayName = $deviceResponse.value[0].displayName
            Success     = $true
        }
    }
    
    return @{
        Id          = $null
        DisplayName = $DeviceName
        Success     = $false
    }
}

function Get-UserInfo {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName
    )

    try {
        $userUri = "$GraphEndpoint/v1.0/users/$UserPrincipalName"
        $user = Invoke-MgGraphRequest -Uri $userUri -Method Get
        return @{
            Id                = $user.id
            UserPrincipalName = $user.userPrincipalName
            Success           = $true
        }
    }
    catch {
        return @{
            Id                = $null
            UserPrincipalName = $UserPrincipalName
            Success           = $false
        }
    }
}

function Get-GroupMemberships {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ObjectId,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet("User", "Device")]
        [string]$ObjectType
    )

    $uri = "$GraphEndpoint/v1.0/$($ObjectType.ToLower())s/$ObjectId/transitiveMemberOf?`$select=id,displayName"
    $response = Invoke-MgGraphRequest -Uri $uri -Method Get
    
    return $response.value
}

function Get-AssignmentInfo {
    param (
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [array]$Assignments
    )

    if ($null -eq $Assignments -or $Assignments.Count -eq 0) {
        return @{
            Type   = "None"
            Target = "Not Assigned"
        }
    }

    $assignment = $Assignments[0]  # Take the first assignment
    $type = switch ($assignment.Reason) {
        "All Users" { "All Users"; break }
        "All Devices" { "All Devices"; break }
        "Group Assignment" { "Group"; break }
        default { "None" }
    }

    $target = switch ($type) {
        "All Users" { "All Users" }
        "All Devices" { "All Devices" }
        "Group" {
            if ($assignment.GroupId) {
                $groupInfo = Get-GroupInfo -GroupId $assignment.GroupId
                $groupInfo.DisplayName
            }
            else {
                "Unknown Group"
            }
        }
        default { "Not Assigned" }
    }

    return @{
        Type   = $type
        Target = $target
    }
}

function Get-AssignmentFailures {
    Write-Host "Fetching assignment failures..." -ForegroundColor Green
    
    $failedAssignments = [System.Collections.ArrayList]::new()
    $headers = @{
        'Authorization' = "Bearer $($global:graphApiToken)"
        'Content-Type'  = 'application/json'
    }
    
    
    Write-Host "Checking device configuration policy failures..." -ForegroundColor Yellow
    try {
        $configPoliciesUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations"
        $configPolicies = (Invoke-MgGraphRequest -Uri $configPoliciesUri -Method GET).value
        
        foreach ($policy in $configPolicies) {
            $statusUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations('$($policy.id)')/deviceStatuses"
            $statuses = (Invoke-MgGraphRequest -Uri $statusUri -Method GET).value
            
            $failures = $statuses | Where-Object { 
                $_.status -in @("error", "conflict", "notApplicable") -or
                $_.complianceGracePeriodExpirationDateTime -and 
                [DateTime]$_.complianceGracePeriodExpirationDateTime -lt [DateTime]::Now
            }
            
            foreach ($failure in $failures) {
                $null = $failedAssignments.Add([PSCustomObject]@{
                        Type             = "Device Configuration"
                        PolicyName       = $policy.displayName
                        Target           = "Device: $($failure.deviceDisplayName)"
                        ErrorCode        = "$($failure.status)"
                        ErrorDescription = if ($failure.userPrincipalName) { "$($failure.userPrincipalName)" } else { "No additional details" }
                        LastAttempt      = $failure.lastReportedDateTime
                    })
            }
        }
    }
    catch {
        Write-Host "Error fetching device configuration failures: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "Checking compliance policy failures..." -ForegroundColor Yellow
    try {
        $compliancePoliciesUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies"
        $compliancePolicies = (Invoke-MgGraphRequest -Uri $compliancePoliciesUri -Method GET).value
        
        foreach ($policy in $compliancePolicies) {
            $statusUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies('$($policy.id)')/deviceStatuses"
            $statuses = (Invoke-MgGraphRequest -Uri $statusUri -Method GET).value
            
            $failures = $statuses | Where-Object { 
                $_.status -in @("error", "conflict", "notApplicable", "nonCompliant")
            }
            
            foreach ($failure in $failures) {
                $null = $failedAssignments.Add([PSCustomObject]@{
                        Type             = "Compliance Policy"
                        PolicyName       = $policy.displayName
                        Target           = "Device: $($failure.deviceDisplayName)"
                        ErrorCode        = "$($failure.status)"
                        ErrorDescription = if ($failure.userPrincipalName) { "$($failure.userPrincipalName)" } else { "No additional details" }
                        LastAttempt      = $failure.lastReportedDateTime
                    })
            }
        }
    }
    catch {
        Write-Host "Error fetching compliance policy failures: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    return $failedAssignments
}

function Show-SaveFileDialog {
    param (
        [string]$DefaultFileName
    )
    
    Add-Type -AssemblyName System.Windows.Forms
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveFileDialog.FileName = $DefaultFileName
    $saveFileDialog.Title = "Save Policy Report"
    
    if ($saveFileDialog.ShowDialog() -eq 'OK') {
        return $saveFileDialog.FileName
    }
    return $null
}

function Export-PolicyData {
    param (
        [Parameter(Mandatory = $true)]
        [System.Collections.ArrayList]$ExportData,
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    
    if ($extension -eq '.xlsx') {
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Host "The ImportExcel module is required for Excel export. Would you like to install it? (y/n)" -ForegroundColor Yellow
            $install = Read-Host
            if ($install -eq 'y') {
                try {
                    Install-Module -Name ImportExcel -Force -Scope CurrentUser
                    Write-Host "ImportExcel module installed successfully." -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to install ImportExcel module. Falling back to CSV export." -ForegroundColor Red
                    $FilePath = [System.IO.Path]::ChangeExtension($FilePath, '.csv')
                    $ExportData | Export-Csv -Path $FilePath -NoTypeInformation
                    Write-Host "Results exported to $FilePath" -ForegroundColor Green
                    return
                }
            }
            else {
                Write-Host "Falling back to CSV export." -ForegroundColor Yellow
                $FilePath = [System.IO.Path]::ChangeExtension($FilePath, '.csv')
                $ExportData | Export-Csv -Path $FilePath -NoTypeInformation
                Write-Host "Results exported to $FilePath" -ForegroundColor Green
                return
            }
        }

        try {
            $ExportData | Export-Excel -Path $FilePath -AutoSize -AutoFilter -WorksheetName "Intune Assignments" -TableName "IntuneAssignments"
            Write-Host "Results exported to $FilePath" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to export to Excel. Falling back to CSV export." -ForegroundColor Red
            $FilePath = [System.IO.Path]::ChangeExtension($FilePath, '.csv')
            $ExportData | Export-Csv -Path $FilePath -NoTypeInformation
            Write-Host "Results exported to $FilePath" -ForegroundColor Green
        }
    }
    else {
        $ExportData | Export-Csv -Path $FilePath -NoTypeInformation
        Write-Host "Results exported to $FilePath" -ForegroundColor Green
    }
}

function Add-ExportData {
    param (
        [System.Collections.ArrayList]$ExportData,
        [string]$Category,
        [object[]]$Items,
        [Parameter(Mandatory = $false)]
        [object]$AssignmentReason = "N/A"
    )
    
    foreach ($item in $Items) {
        $itemName = if ($item.displayName) { $item.displayName } else { $item.name }
        
        $reason = if ($AssignmentReason -is [scriptblock]) {
            & $AssignmentReason $item
        }
        elseif ($item.AssignmentReason) {
            $item.AssignmentReason
        }
        elseif ($item.AssignmentSummary) {
            $item.AssignmentSummary
        }
        else {
            $AssignmentReason
        }
        
        $null = $ExportData.Add([PSCustomObject]@{
                Category         = $Category
                Item             = "$itemName (ID: $($item.id))"
                AssignmentReason = $reason
            })
    }
}

function Add-AppExportData {
    param (
        [System.Collections.ArrayList]$ExportData,
        [string]$Category,
        [object[]]$Apps,
        [string]$AssignmentReason = "N/A"
    )
    
    foreach ($app in $Apps) {
        $appName = if ($app.displayName) { $app.displayName } else { $app.name }
        $null = $ExportData.Add([PSCustomObject]@{
                Category         = $Category
                Item             = "$appName (ID: $($app.id))"
                AssignmentReason = "$AssignmentReason - $($app.AssignmentIntent)"
            })
    }
}

function Show-Menu {    
    Write-Host "Assignment Checks:" -ForegroundColor Cyan
    Write-Host "  [1] Check User(s) Assignments" -ForegroundColor White
    Write-Host "  [2] Check Group(s) Assignments" -ForegroundColor White
    Write-Host "  [3] Check Device(s) Assignments" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Policy Overview:" -ForegroundColor Cyan
    Write-Host "  [4] Show All Policies and Their Assignments" -ForegroundColor White
    Write-Host "  [5] Show All 'All Users' Assignments" -ForegroundColor White
    Write-Host "  [6] Show All 'All Devices' Assignments" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Advanced Options:" -ForegroundColor Cyan
    Write-Host "  [7] Generate HTML Report" -ForegroundColor White
    Write-Host "  [8] Show Policies Without Assignments" -ForegroundColor White
    Write-Host "  [9] Check for Empty Groups in Assignments" -ForegroundColor White
    Write-Host "  [10] Compare Assignments Between Groups" -ForegroundColor White
    Write-Host "  [11] Show All Failed Assignments" -ForegroundColor White
    Write-Host ""
    
    Write-Host "System:" -ForegroundColor Cyan
    Write-Host "  [0] Exit" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Select an option: " -ForegroundColor Yellow -NoNewline
}

function Export-ResultsIfRequested {
    param (
        [System.Collections.ArrayList]$ExportData,
        [string]$DefaultFileName,
        [switch]$ForceExport,
        [string]$CustomExportPath
    )
    
    if ($ForceExport -or $ExportToCSV) {
        $exportPath = if ($CustomExportPath) {
            $CustomExportPath
        }
        else {
            Show-SaveFileDialog -DefaultFileName $DefaultFileName
        }
        
        if ($exportPath) {
            Export-PolicyData -ExportData $ExportData -FilePath $exportPath
        }
    }
    else {
        $export = Read-Host "`nWould you like to export the results to CSV? (y/n)"
        if ($export -eq 'y') {
            $exportPath = Show-SaveFileDialog -DefaultFileName $DefaultFileName
            if ($exportPath) {
                Export-PolicyData -ExportData $ExportData -FilePath $exportPath
            }
        }
    }
}


do {
    if (-not $parameterMode) {
        Show-Menu
        $selection = Read-Host
    }
    else {
        $selection = $selectedOption
    }

    switch ($selection) {
        '1' {
            Write-Host "User selection chosen" -ForegroundColor Green

            if ($parameterMode -and $UserPrincipalNames) {
                $upnInput = $UserPrincipalNames
            }
            else {
                Write-Host "Please enter User Principal Name(s), separated by commas (,): " -ForegroundColor Cyan
                $upnInput = Read-Host
            }
    
            if ([string]::IsNullOrWhiteSpace($upnInput)) {
                Write-Host "No UPN provided. Please try again with a valid UPN." -ForegroundColor Red
                if ($parameterMode) { exit 1 } else { continue }
            }
    
            $upns = $upnInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    
            if ($upns.Count -eq 0) {
                Write-Host "No valid UPNs provided. Please try again with at least one valid UPN." -ForegroundColor Red
                if ($parameterMode) { exit 1 } else { continue }
            }

            $exportData = [System.Collections.ArrayList]::new()

            foreach ($upn in $upns) {
                Write-Host "Checking following UPN: $upn" -ForegroundColor Yellow

                $userInfo = Get-UserInfo -UserPrincipalName $upn
                if (-not $userInfo.Success) {
                    Write-Host "User not found: $upn" -ForegroundColor Red
                    Write-Host "Please verify the User Principal Name is correct." -ForegroundColor Yellow
                    continue
                }

                try {
                    $groupMemberships = Get-GroupMemberships -ObjectId $userInfo.Id -ObjectType "User"
                    Write-Host "User Group Memberships: $($groupMemberships.displayName -join ', ')" -ForegroundColor Green
                }
                catch {
                    Write-Host "Error fetching group memberships for user: $upn" -ForegroundColor Red
                    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
                    continue
                }

                Write-Host "Fetching Intune Profiles and Applications for the user ... (this takes a few seconds)" -ForegroundColor Yellow

                $relevantPolicies = @{
                    DeviceConfigs               = @()
                    SettingsCatalog             = @()
                    AdminTemplates              = @()
                    CompliancePolicies          = @()
                    AppProtectionPolicies       = @()
                    AppConfigurationPolicies    = @()
                    AppsRequired                = @()
                    AppsAvailable               = @()
                    AppsUninstall               = @()
                    PlatformScripts             = @()
                    HealthScripts               = @()
                    AntivirusProfiles           = @()
                    DiskEncryptionProfiles      = @()
                    FirewallProfiles            = @()
                    EndpointDetectionProfiles   = @()
                    AttackSurfaceProfiles       = @()
                    DeploymentProfiles          = @()
                    ESPProfiles                 = @()
                    CloudPCProvisioningPolicies = @()
                    CloudPCUserSettings         = @()
                }

                Write-Host "Fetching Device Configurations..." -ForegroundColor Yellow
                $deviceConfigs = Get-IntuneEntities -EntityType "deviceConfigurations"
                foreach ($config in $deviceConfigs) {
                    $assignments = Get-IntuneAssignments -EntityType "deviceConfigurations" -EntityId $config.id
                    foreach ($assignment in $assignments) {
                        if ($assignment.Reason -eq "All Users" -or
                            ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId)) {
                            $config | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                            $relevantPolicies.DeviceConfigs += $config
                            break
                        }
                        elseif ($assignment.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignment.GroupId) {
                            $config | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                            $relevantPolicies.DeviceConfigs += $config
                            break
                        }
                    }
                }

                Write-Host "Fetching Settings Catalog Policies..." -ForegroundColor Yellow
                $settingsCatalog = Get-IntuneEntities -EntityType "configurationPolicies"
                foreach ($policy in $settingsCatalog) {
                    $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                    foreach ($assignment in $assignments) {
                        if ($assignment.Reason -eq "All Users" -or
                            ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId)) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                            $relevantPolicies.SettingsCatalog += $policy
                            break
                        }
                        elseif ($assignment.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignment.GroupId) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                            $relevantPolicies.SettingsCatalog += $policy
                            break
                        }
                    }
                }

                Write-Host "Fetching Administrative Templates..." -ForegroundColor Yellow
                $adminTemplates = Get-IntuneEntities -EntityType "groupPolicyConfigurations"
                foreach ($template in $adminTemplates) {
                    $assignments = Get-IntuneAssignments -EntityType "groupPolicyConfigurations" -EntityId $template.id
                    foreach ($assignment in $assignments) {
                        if ($assignment.Reason -eq "All Users" -or
                            ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId)) {
                            $template | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                            $relevantPolicies.AdminTemplates += $template
                            break
                        }
                        elseif ($assignment.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignment.GroupId) {
                            $template | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                            $relevantPolicies.AdminTemplates += $template
                            break
                        }
                    }
                }

                Write-Host "Fetching Compliance Policies..." -ForegroundColor Yellow
                $compliancePolicies = Get-IntuneEntities -EntityType "deviceCompliancePolicies"
                foreach ($policy in $compliancePolicies) {
                    $assignments = Get-IntuneAssignments -EntityType "deviceCompliancePolicies" -EntityId $policy.id
                    foreach ($assignment in $assignments) {
                        if ($assignment.Reason -eq "All Users" -or
                            ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId)) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                            $relevantPolicies.CompliancePolicies += $policy
                            break
                        }
                        elseif ($assignment.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignment.GroupId) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                            $relevantPolicies.CompliancePolicies += $policy
                            break
                        }
                    }
                }

                Write-Host "Fetching App Protection Policies..." -ForegroundColor Yellow
                $appProtectionPolicies = Get-IntuneEntities -EntityType "deviceAppManagement/managedAppPolicies"
                foreach ($policy in $appProtectionPolicies) {
                    $policyType = $policy.'@odata.type'
                    $assignmentsUri = switch ($policyType) {
                        "#microsoft.graph.androidManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/androidManagedAppProtections('$($policy.id)')/assignments" }
                        "#microsoft.graph.iosManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/iosManagedAppProtections('$($policy.id)')/assignments" }
                        "#microsoft.graph.windowsManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/windowsManagedAppProtections('$($policy.id)')/assignments" }
                        default { $null }
                    }

                    if ($assignmentsUri) {
                        try {
                            $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get
                            $assignments = @()
                            foreach ($assignment in $assignmentResponse.value) {
                                $assignmentReason = $null
                                switch ($assignment.target.'@odata.type') {
                                    '#microsoft.graph.allLicensedUsersAssignmentTarget' {
                                        $assignmentReason = "All Users"
                                    }
                                    '#microsoft.graph.groupAssignmentTarget' {
                                        if (!$GroupId -or $assignment.target.groupId -eq $GroupId) {
                                            $assignmentReason = "Group Assignment"
                                        }
                                    }
                                    '#microsoft.graph.exclusionGroupAssignmentTarget' {
                                        if (!$GroupId -or $assignment.target.groupId -eq $GroupId) {
                                            $assignmentReason = "Group Exclusion"
                                        }
                                    }
                                }

                                if ($assignmentReason) {
                                    $assignments += @{
                                        Reason  = $assignmentReason
                                        GroupId = $assignment.target.groupId
                                    }
                                }
                            }

                            if ($assignments.Count -gt 0) {
                                $assignmentSummary = $assignments | ForEach-Object {
                                    if ($_.Reason -eq "Group Assignment" -or $_.Reason -eq "Group Exclusion") {
                                        $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                                        $color = if ($_.Reason -eq "Group Exclusion") { "Red" } else { "White" }
                                        "$($_.Reason) - $($groupInfo.DisplayName)"
                                    }
                                    else {
                                        $_.Reason
                                    }
                                }
                                $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                                $relevantPolicies.AppProtectionPolicies += $policy
                            }
                        }
                        catch {
                            Write-Host "Error fetching assignments for policy $($policy.displayName): $($_.Exception.Message)" -ForegroundColor Red
                        }
                    }
                }

                Write-Host "Fetching App Configuration Policies..." -ForegroundColor Yellow
                $appConfigPolicies = Get-IntuneEntities -EntityType "deviceAppManagement/mobileAppConfigurations"
                foreach ($policy in $appConfigPolicies) {
                    $assignments = Get-IntuneAssignments -EntityType "mobileAppConfigurations" -EntityId $policy.id
                    foreach ($assignment in $assignments) {
                        if ($assignment.Reason -eq "All Users" -or 
                            ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId)) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                            $relevantPolicies.AppConfigurationPolicies += $policy
                            break
                        }
                    }
                }

                Write-Host "Fetching Applications..." -ForegroundColor Yellow
                $appUri = "$GraphEndpoint/beta/deviceAppManagement/mobileApps?`$filter=isAssigned eq true"
                $appResponse = Invoke-MgGraphRequest -Uri $appUri -Method Get
                $allApps = $appResponse.value
                while ($appResponse.'@odata.nextLink') {
                    $appResponse = Invoke-MgGraphRequest -Uri $appResponse.'@odata.nextLink' -Method Get
                    $allApps += $appResponse.value
                }
                $totalApps = $allApps.Count
                $currentApp = 0

                foreach ($app in $allApps) {
                    if ($app.isFeatured -or $app.isBuiltIn) {
                        continue
                    }

                    $currentApp++
                    Write-Host "`rFetching Application $currentApp of $totalApps" -NoNewline
                    $appId = $app.id
                    $assignmentsUri = "$GraphEndpoint/beta/deviceAppManagement/mobileApps('$appId')/assignments"
                    $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get

                    $isExcluded = $false
                    $isIncluded = $false

                    foreach ($assignment in $assignmentResponse.value) {
                        if ($assignment.target.'@odata.type' -eq '#microsoft.graph.exclusionGroupAssignmentTarget' -and
                            $groupMemberships.id -contains $assignment.target.groupId) {
                            $isExcluded = $true
                            break
                        }
                        elseif ($assignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget' -or
                            ($assignment.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget' -and
                            $groupMemberships.id -contains $assignment.target.groupId)) {
                            $isIncluded = $true
                        }
                    }

                    if ($isIncluded -and -not $isExcluded) {
                        $appWithReason = $app.PSObject.Copy()
                        $appWithReason | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Included" -Force
                        switch ($assignment.intent) {
                            "required" { $relevantPolicies.AppsRequired += $appWithReason; break }
                            "available" { $relevantPolicies.AppsAvailable += $appWithReason; break }
                            "uninstall" { $relevantPolicies.AppsUninstall += $appWithReason; break }
                        }
                    }
                    elseif ($isExcluded) {
                        $appWithReason = $app.PSObject.Copy()
                        $appWithReason | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                        switch ($assignment.intent) {
                            "required" { $relevantPolicies.AppsRequired += $appWithReason; break }
                            "available" { $relevantPolicies.AppsAvailable += $appWithReason; break }
                            "uninstall" { $relevantPolicies.AppsUninstall += $appWithReason; break }
                        }
                    }
                }
                Write-Host "`rFetching Application $totalApps of $totalApps" -NoNewline
                Start-Sleep -Milliseconds 100
                Write-Host ""  # Move to the next line after the loop

                Write-Host "Fetching Platform Scripts..." -ForegroundColor Yellow
                $platformScripts = Get-IntuneEntities -EntityType "deviceManagementScripts"
                foreach ($script in $platformScripts) {
                    $assignments = Get-IntuneAssignments -EntityType "deviceManagementScripts" -EntityId $script.id
                    foreach ($assignment in $assignments) {
                        if ($assignment.Reason -eq "All Users" -or 
                            ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId)) {
                            $script | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                            $relevantPolicies.PlatformScripts += $script
                            break
                        }
                    }
                }

                Write-Host "Fetching Proactive Remediation Scripts..." -ForegroundColor Yellow
                $healthScripts = Get-IntuneEntities -EntityType "deviceHealthScripts"
                foreach ($script in $healthScripts) {
                    $assignments = Get-IntuneAssignments -EntityType "deviceHealthScripts" -EntityId $script.id
                    foreach ($assignment in $assignments) {
                        if ($assignment.Reason -eq "All Users" -or 
                            ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId)) {
                            $script | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                            $relevantPolicies.HealthScripts += $script
                            break
                        }
                    }
                }

                Write-Host "Fetching Antivirus Policies..." -ForegroundColor Yellow
                $antivirusPoliciesFound = [System.Collections.ArrayList]::new()
                $processedAntivirusIds = [System.Collections.Generic.HashSet[string]]::new()

                $configPoliciesForAntivirus = Get-IntuneEntities -EntityType "configurationPolicies"
                $matchingConfigPoliciesAntivirus = $configPoliciesForAntivirus | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAntivirus' }

                if ($matchingConfigPoliciesAntivirus) {
                    foreach ($policy in $matchingConfigPoliciesAntivirus) {
                        if ($processedAntivirusIds.Add($policy.id)) {
                            $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                            foreach ($assignmentDetail in $assignments) {
                                if ($assignmentDetail.Reason -eq "All Users" -or
                                    ($assignmentDetail.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetail.GroupId)) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetail.Reason -Force
                                    [void]$antivirusPoliciesFound.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetail.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetail.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$antivirusPoliciesFound.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }

                $allIntentsForAntivirus = Get-IntuneEntities -EntityType "deviceManagement/intents"
                $matchingIntentsAntivirus = $allIntentsForAntivirus | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAntivirus' }
                
                if ($matchingIntentsAntivirus) {
                    foreach ($policy in $matchingIntentsAntivirus) {
                        if ($processedAntivirusIds.Add($policy.id)) {
                            $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                            $assignments = $assignmentsResponse.value
                            foreach ($assignment in $assignments) {
                                $assignmentDetails = @{
                                    Reason  = switch ($assignment.target.'@odata.type') {
                                        '#microsoft.graph.allLicensedUsersAssignmentTarget' { "All Users" }
                                        '#microsoft.graph.allDevicesAssignmentTarget' { "All Devices" }
                                        '#microsoft.graph.groupAssignmentTarget' { "Group Assignment" }
                                        '#microsoft.graph.exclusionGroupAssignmentTarget' { "Group Exclusion" }
                                        default { "Unknown" }
                                    }
                                    GroupId = if ($assignment.target.'@odata.type' -match "groupAssignmentTarget") { $assignment.target.groupId } else { $null }
                                }
                                if ($assignmentDetails.Reason -eq "All Users" -or
                                    ($assignmentDetails.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetails.GroupId)) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetails.Reason -Force
                                    [void]$antivirusPoliciesFound.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetails.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetails.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$antivirusPoliciesFound.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }
                $relevantPolicies.AntivirusProfiles = $antivirusPoliciesFound
                
                Write-Host "Fetching Disk Encryption Policies..." -ForegroundColor Yellow
                $diskEncryptionPoliciesFound = [System.Collections.ArrayList]::new()
                $processedDiskEncryptionIds = [System.Collections.Generic.HashSet[string]]::new()

                $configPoliciesForDiskEncryption = Get-IntuneEntities -EntityType "configurationPolicies"
                $matchingConfigPoliciesDiskEnc = $configPoliciesForDiskEncryption | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityDiskEncryption' }
                
                if ($matchingConfigPoliciesDiskEnc) {
                    foreach ($policy in $matchingConfigPoliciesDiskEnc) {
                        if ($processedDiskEncryptionIds.Add($policy.id)) {
                            $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                            foreach ($assignmentDetail in $assignments) {
                                if ($assignmentDetail.Reason -eq "All Users" -or
                                    ($assignmentDetail.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetail.GroupId)) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetail.Reason -Force
                                    [void]$diskEncryptionPoliciesFound.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetail.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetail.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$diskEncryptionPoliciesFound.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }

                $allIntentsForDiskEncryption = Get-IntuneEntities -EntityType "deviceManagement/intents"
                $matchingIntentsDiskEnc = $allIntentsForDiskEncryption | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityDiskEncryption' }

                if ($matchingIntentsDiskEnc) {
                    foreach ($policy in $matchingIntentsDiskEnc) {
                        if ($processedDiskEncryptionIds.Add($policy.id)) {
                            $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                            $assignments = $assignmentsResponse.value
                            
                            foreach ($assignment in $assignments) {
                                $assignmentDetails = @{
                                    Reason  = switch ($assignment.target.'@odata.type') {
                                        '#microsoft.graph.allLicensedUsersAssignmentTarget' { "All Users" }
                                        '#microsoft.graph.allDevicesAssignmentTarget' { "All Devices" }
                                        '#microsoft.graph.groupAssignmentTarget' { "Group Assignment" }
                                        '#microsoft.graph.exclusionGroupAssignmentTarget' { "Group Exclusion" }
                                        default { "Unknown" }
                                    }
                                    GroupId = if ($assignment.target.'@odata.type' -match "groupAssignmentTarget") { $assignment.target.groupId } else { $null }
                                }

                                if ($assignmentDetails.Reason -eq "All Users" -or
                                    ($assignmentDetails.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetails.GroupId)) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetails.Reason -Force
                                    [void]$diskEncryptionPoliciesFound.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetails.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetails.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$diskEncryptionPoliciesFound.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }
                $relevantPolicies.DiskEncryptionProfiles = $diskEncryptionPoliciesFound
                
                Write-Host "Fetching Firewall Policies..." -ForegroundColor Yellow
                $firewallPoliciesFound = [System.Collections.ArrayList]::new()
                $processedFirewallIds = [System.Collections.Generic.HashSet[string]]::new()

                $configPoliciesForFirewall = Get-IntuneEntities -EntityType "configurationPolicies"
                $matchingConfigPoliciesFirewall = $configPoliciesForFirewall | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityFirewall' }

                if ($matchingConfigPoliciesFirewall) {
                    foreach ($policy in $matchingConfigPoliciesFirewall) {
                        if ($processedFirewallIds.Add($policy.id)) {
                            $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                            foreach ($assignmentDetail in $assignments) {
                                if ($assignmentDetail.Reason -eq "All Users" -or
                                    ($assignmentDetail.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetail.GroupId)) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetail.Reason -Force
                                    [void]$firewallPoliciesFound.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetail.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetail.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$firewallPoliciesFound.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }

                $allIntentsForFirewall = Get-IntuneEntities -EntityType "deviceManagement/intents"
                $matchingIntentsFirewall = $allIntentsForFirewall | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityFirewall' }

                if ($matchingIntentsFirewall) {
                    foreach ($policy in $matchingIntentsFirewall) {
                        if ($processedFirewallIds.Add($policy.id)) {
                            $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                            $assignments = $assignmentsResponse.value
                            foreach ($assignment in $assignments) {
                                $assignmentDetails = @{
                                    Reason  = switch ($assignment.target.'@odata.type') {
                                        '#microsoft.graph.allLicensedUsersAssignmentTarget' { "All Users" }
                                        '#microsoft.graph.allDevicesAssignmentTarget' { "All Devices" }
                                        '#microsoft.graph.groupAssignmentTarget' { "Group Assignment" }
                                        '#microsoft.graph.exclusionGroupAssignmentTarget' { "Group Exclusion" }
                                        default { "Unknown" }
                                    }
                                    GroupId = if ($assignment.target.'@odata.type' -match "groupAssignmentTarget") { $assignment.target.groupId } else { $null }
                                }
                                if ($assignmentDetails.Reason -eq "All Users" -or
                                    ($assignmentDetails.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetails.GroupId)) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetails.Reason -Force
                                    [void]$firewallPoliciesFound.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetails.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetails.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$firewallPoliciesFound.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }
                $relevantPolicies.FirewallProfiles = $firewallPoliciesFound
                
                Write-Host "Fetching Endpoint Detection and Response Policies..." -ForegroundColor Yellow
                $edrPoliciesFound = [System.Collections.ArrayList]::new()
                $processedEDRIds = [System.Collections.Generic.HashSet[string]]::new()

                $configPoliciesForEDR = Get-IntuneEntities -EntityType "configurationPolicies"
                $matchingConfigPoliciesEDR = $configPoliciesForEDR | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityEndpointDetectionAndResponse' }

                if ($matchingConfigPoliciesEDR) {
                    foreach ($policy in $matchingConfigPoliciesEDR) {
                        if ($processedEDRIds.Add($policy.id)) {
                            $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                            foreach ($assignmentDetail in $assignments) {
                                if ($assignmentDetail.Reason -eq "All Users" -or
                                    ($assignmentDetail.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetail.GroupId)) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetail.Reason -Force
                                    [void]$edrPoliciesFound.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetail.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetail.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$edrPoliciesFound.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }

                $allIntentsForEDR = Get-IntuneEntities -EntityType "deviceManagement/intents"
                $matchingIntentsEDR = $allIntentsForEDR | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityEndpointDetectionAndResponse' }

                if ($matchingIntentsEDR) {
                    foreach ($policy in $matchingIntentsEDR) {
                        if ($processedEDRIds.Add($policy.id)) {
                            $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                            $assignments = $assignmentsResponse.value
                            foreach ($assignment in $assignments) {
                                $assignmentDetails = @{
                                    Reason  = switch ($assignment.target.'@odata.type') {
                                        '#microsoft.graph.allLicensedUsersAssignmentTarget' { "All Users" }
                                        '#microsoft.graph.allDevicesAssignmentTarget' { "All Devices" }
                                        '#microsoft.graph.groupAssignmentTarget' { "Group Assignment" }
                                        '#microsoft.graph.exclusionGroupAssignmentTarget' { "Group Exclusion" }
                                        default { "Unknown" }
                                    }
                                    GroupId = if ($assignment.target.'@odata.type' -match "groupAssignmentTarget") { $assignment.target.groupId } else { $null }
                                }
                                if ($assignmentDetails.Reason -eq "All Users" -or
                                    ($assignmentDetails.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetails.GroupId)) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetails.Reason -Force
                                    [void]$edrPoliciesFound.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetails.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetails.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$edrPoliciesFound.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }
                $relevantPolicies.EndpointDetectionProfiles = $edrPoliciesFound
                
                Write-Host "Fetching Attack Surface Reduction Policies..." -ForegroundColor Yellow
                $asrPoliciesFound = [System.Collections.ArrayList]::new()
                $processedASRIds = [System.Collections.Generic.HashSet[string]]::new()

                $configPoliciesForASR = Get-IntuneEntities -EntityType "configurationPolicies"
                $matchingConfigPoliciesASR = $configPoliciesForASR | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAttackSurfaceReductionRules' }

                if ($matchingConfigPoliciesASR) {
                    foreach ($policy in $matchingConfigPoliciesASR) {
                        if ($processedASRIds.Add($policy.id)) {
                            $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                            foreach ($assignmentDetail in $assignments) {
                                if ($assignmentDetail.Reason -eq "All Users" -or
                                    ($assignmentDetail.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetail.GroupId)) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetail.Reason -Force
                                    [void]$asrPoliciesFound.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetail.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetail.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$asrPoliciesFound.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }

                $allIntentsForASR = Get-IntuneEntities -EntityType "deviceManagement/intents"
                $matchingIntentsASR = $allIntentsForASR | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAttackSurfaceReductionRules' }

                if ($matchingIntentsASR) {
                    foreach ($policy in $matchingIntentsASR) {
                        if ($processedASRIds.Add($policy.id)) {
                            $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                            $assignments = $assignmentsResponse.value
                            foreach ($assignment in $assignments) {
                                $assignmentDetails = @{
                                    Reason  = switch ($assignment.target.'@odata.type') {
                                        '#microsoft.graph.allLicensedUsersAssignmentTarget' { "All Users" }
                                        '#microsoft.graph.allDevicesAssignmentTarget' { "All Devices" }
                                        '#microsoft.graph.groupAssignmentTarget' { "Group Assignment" }
                                        '#microsoft.graph.exclusionGroupAssignmentTarget' { "Group Exclusion" }
                                        default { "Unknown" }
                                    }
                                    GroupId = if ($assignment.target.'@odata.type' -match "groupAssignmentTarget") { $assignment.target.groupId } else { $null }
                                }
                                if ($assignmentDetails.Reason -eq "All Users" -or
                                    ($assignmentDetails.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetails.GroupId)) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetails.Reason -Force
                                    [void]$asrPoliciesFound.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetails.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetails.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$asrPoliciesFound.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }
                $relevantPolicies.AttackSurfaceProfiles = $asrPoliciesFound

                Write-Host "Fetching Windows 365 Cloud PC Provisioning Policies..." -ForegroundColor Yellow
                try {
                    $cloudPCProvisioningPolicies = Get-IntuneEntities -EntityType "virtualEndpoint/provisioningPolicies"
                    foreach ($policy in $cloudPCProvisioningPolicies) {
                        $assignments = Get-IntuneAssignments -EntityType "virtualEndpoint/provisioningPolicies" -EntityId $policy.id
                        foreach ($assignment in $assignments) {
                            if ($assignment.Reason -eq "All Users" -or
                                ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId)) {
                                $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                                $relevantPolicies.CloudPCProvisioningPolicies += $policy
                                break
                            }
                            elseif ($assignment.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignment.GroupId) {
                                $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                $relevantPolicies.CloudPCProvisioningPolicies += $policy
                                break
                            }
                        }
                    }
                }
                catch {
                    Write-Warning "Unable to fetch Windows 365 Cloud PC Provisioning Policies: $($_.Exception.Message)"
                }

                Write-Host "Fetching Windows 365 Cloud PC User Settings..." -ForegroundColor Yellow
                try {
                    $cloudPCUserSettings = Get-IntuneEntities -EntityType "virtualEndpoint/userSettings"
                    foreach ($setting in $cloudPCUserSettings) {
                        $assignments = Get-IntuneAssignments -EntityType "virtualEndpoint/userSettings" -EntityId $setting.id
                        foreach ($assignment in $assignments) {
                            if ($assignment.Reason -eq "All Users" -or
                                ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId)) {
                                $setting | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                                $relevantPolicies.CloudPCUserSettings += $setting
                                break
                            }
                            elseif ($assignment.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignment.GroupId) {
                                $setting | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                $relevantPolicies.CloudPCUserSettings += $setting
                                break
                            }
                        }
                    }
                }
                catch {
                    Write-Warning "Unable to fetch Windows 365 Cloud PC User Settings: $($_.Exception.Message)"
                }

                Write-Host "`nAssignments for User: $upn" -ForegroundColor Green

                Write-Host "`n------- Device Configurations -------" -ForegroundColor Cyan
                if ($relevantPolicies.DeviceConfigs.Count -eq 0) {
                    Write-Host "No Device Configurations found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "Configuration Name", "Configuration ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($config in $relevantPolicies.DeviceConfigs) {
                        $configName = if ([string]::IsNullOrWhiteSpace($config.name)) { $config.displayName } else { $config.name }
                        if ($configName.Length -gt 47) {
                            $configName = $configName.Substring(0, 44) + "..."
                        }
                        
                        $configId = $config.id
                        if ($configId.Length -gt 37) {
                            $configId = $configId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $config.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $configName, $configId, $assignment
                        if ($assignment -eq "Excluded") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }

                Write-Host "`n------- Settings Catalog Policies -------" -ForegroundColor Cyan
                if ($relevantPolicies.SettingsCatalog.Count -eq 0) {
                    Write-Host "No Settings Catalog Policies found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "Policy Name", "Policy ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($policy in $relevantPolicies.SettingsCatalog) {
                        $policyName = if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                        if ($policyName.Length -gt 47) {
                            $policyName = $policyName.Substring(0, 44) + "..."
                        }
                        
                        $policyId = $policy.id
                        if ($policyId.Length -gt 37) {
                            $policyId = $policyId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $policy.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $policyName, $policyId, $assignment
                        if ($assignment -eq "Excluded") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }

                Write-Host "`n------- Administrative Templates -------" -ForegroundColor Cyan
                if ($relevantPolicies.AdminTemplates.Count -eq 0) {
                    Write-Host "No Administrative Templates found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "Template Name", "Template ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($template in $relevantPolicies.AdminTemplates) {
                        $templateName = if ([string]::IsNullOrWhiteSpace($template.name)) { $template.displayName } else { $template.name }
                        if ($templateName.Length -gt 47) {
                            $templateName = $templateName.Substring(0, 44) + "..."
                        }
                        
                        $templateId = $template.id
                        if ($templateId.Length -gt 37) {
                            $templateId = $templateId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $template.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $templateName, $templateId, $assignment
                        if ($assignment -eq "Excluded") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }

                Write-Host "`n------- Compliance Policies -------" -ForegroundColor Cyan
                if ($relevantPolicies.CompliancePolicies.Count -eq 0) {
                    Write-Host "No Compliance Policies found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "Policy Name", "Policy ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($policy in $relevantPolicies.CompliancePolicies) {
                        $policyName = if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                        if ($policyName.Length -gt 47) {
                            $policyName = $policyName.Substring(0, 44) + "..."
                        }
                        
                        $policyId = $policy.id
                        if ($policyId.Length -gt 37) {
                            $policyId = $policyId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $policy.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $policyName, $policyId, $assignment
                        if ($assignment -eq "Excluded") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }

                Write-Host "`n------- App Protection Policies -------" -ForegroundColor Cyan
                if ($relevantPolicies.AppProtectionPolicies.Count -eq 0) {
                    Write-Host "No App Protection Policies found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-40} {1,-30} {2,-20} {3,-30}" -f "Policy Name", "Policy ID", "Type", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($policy in $relevantPolicies.AppProtectionPolicies) {
                        $policyName = $policy.displayName
                        if ($policyName.Length -gt 37) {
                            $policyName = $policyName.Substring(0, 34) + "..."
                        }
                        
                        $policyId = $policy.id
                        if ($policyId.Length -gt 27) {
                            $policyId = $policyId.Substring(0, 24) + "..."
                        }
                        
                        $policyType = switch ($policy.'@odata.type') {
                            "#microsoft.graph.androidManagedAppProtection" { "Android" }
                            "#microsoft.graph.iosManagedAppProtection" { "iOS" }
                            "#microsoft.graph.windowsManagedAppProtection" { "Windows" }
                            default { "Unknown" }
                        }
                        
                        $assignment = $policy.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-40} {1,-30} {2,-20} {3,-30}" -f $policyName, $policyId, $policyType, $assignment
                        if ($assignment -eq "Excluded") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }

                Write-Host "`n------- App Configuration Policies -------" -ForegroundColor Cyan
                if ($relevantPolicies.AppConfigurationPolicies.Count -eq 0) {
                    Write-Host "No App Configuration Policies found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "Policy Name", "Policy ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($policy in $relevantPolicies.AppConfigurationPolicies) {
                        $policyName = if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                        if ($policyName.Length -gt 47) {
                            $policyName = $policyName.Substring(0, 44) + "..."
                        }
                        
                        $policyId = $policy.id
                        if ($policyId.Length -gt 37) {
                            $policyId = $policyId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $policy.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $policyName, $policyId, $assignment
                        if ($assignment -eq "Excluded") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }

                Write-Host "`n------- Platform Scripts -------" -ForegroundColor Cyan
                if ($relevantPolicies.PlatformScripts.Count -eq 0) {
                    Write-Host "No Platform Scripts found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "Script Name", "Script ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($script in $relevantPolicies.PlatformScripts) {
                        $scriptName = if ([string]::IsNullOrWhiteSpace($script.name)) { $script.displayName } else { $script.name }
                        if ($scriptName.Length -gt 47) {
                            $scriptName = $scriptName.Substring(0, 44) + "..."
                        }
                        
                        $scriptId = $script.id
                        if ($scriptId.Length -gt 37) {
                            $scriptId = $scriptId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $script.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $scriptName, $scriptId, $assignment
                        if ($assignment -eq "Excluded") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }

                Write-Host "`n------- Proactive Remediation Scripts -------" -ForegroundColor Cyan
                if ($relevantPolicies.HealthScripts.Count -eq 0) {
                    Write-Host "No Proactive Remediation Scripts found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "Script Name", "Script ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($script in $relevantPolicies.HealthScripts) {
                        $scriptName = if ([string]::IsNullOrWhiteSpace($script.name)) { $script.displayName } else { $script.name }
                        if ($scriptName.Length -gt 47) {
                            $scriptName = $scriptName.Substring(0, 44) + "..."
                        }
                        
                        $scriptId = $script.id
                        if ($scriptId.Length -gt 37) {
                            $scriptId = $scriptId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $script.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $scriptName, $scriptId, $assignment
                        if ($assignment -eq "Excluded") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }

                Write-Host "`n------- Required Apps -------" -ForegroundColor Cyan
                if ($relevantPolicies.AppsRequired.Count -eq 0) {
                    Write-Host "No Required Apps found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "App Name", "App ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($app in $relevantPolicies.AppsRequired) {
                        $appName = $app.displayName
                        if ($appName.Length -gt 47) {
                            $appName = $appName.Substring(0, 44) + "..."
                        }
                        
                        $appId = $app.id
                        if ($appId.Length -gt 37) {
                            $appId = $appId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $app.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $appName, $appId, $assignment
                        if ($assignment -like "*Exclusion*") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }

                Write-Host "`n------- Available Apps -------" -ForegroundColor Cyan
                if ($relevantPolicies.AppsAvailable.Count -eq 0) {
                    Write-Host "No Available Apps found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "App Name", "App ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($app in $relevantPolicies.AppsAvailable) {
                        $appName = $app.displayName
                        if ($appName.Length -gt 47) {
                            $appName = $appName.Substring(0, 44) + "..."
                        }
                        
                        $appId = $app.id
                        if ($appId.Length -gt 37) {
                            $appId = $appId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $app.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $appName, $appId, $assignment
                        if ($assignment -like "*Exclusion*") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }

                Write-Host "`n------- Uninstall Apps -------" -ForegroundColor Cyan
                if ($relevantPolicies.AppsUninstall.Count -eq 0) {
                    Write-Host "No Uninstall Apps found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "App Name", "App ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($app in $relevantPolicies.AppsUninstall) {
                        $appName = $app.displayName
                        if ($appName.Length -gt 47) {
                            $appName = $appName.Substring(0, 44) + "..."
                        }
                        
                        $appId = $app.id
                        if ($appId.Length -gt 37) {
                            $appId = $appId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $app.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $appName, $appId, $assignment
                        if ($assignment -like "*Exclusion*") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }
                
                Write-Host "`n------- Endpoint Security - Antivirus Profiles -------" -ForegroundColor Cyan
                if ($relevantPolicies.AntivirusProfiles.Count -eq 0) {
                    Write-Host "No Antivirus Profiles found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "Profile Name", "Profile ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($profile in $relevantPolicies.AntivirusProfiles) {
                        $profileName = if (-not [string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.displayName } elseif (-not [string]::IsNullOrWhiteSpace($profile.name)) { $profile.name } else { "Unnamed Profile" }
                        if ($profileName.Length -gt 47) {
                            $profileName = $profileName.Substring(0, 44) + "..."
                        }
                        
                        $profileId = $profile.id
                        if ($profileId.Length -gt 37) {
                            $profileId = $profileId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $profile.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $profileName, $profileId, $assignment
                        if ($assignment -eq "Excluded") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }
                
                Write-Host "`n------- Endpoint Security - Disk Encryption Profiles -------" -ForegroundColor Cyan
                if ($relevantPolicies.DiskEncryptionProfiles.Count -eq 0) {
                    Write-Host "No Disk Encryption Profiles found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "Profile Name", "Profile ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($profile in $relevantPolicies.DiskEncryptionProfiles) {
                        $profileName = if (-not [string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.displayName } elseif (-not [string]::IsNullOrWhiteSpace($profile.name)) { $profile.name } else { "Unnamed Profile" }
                        if ($profileName.Length -gt 47) {
                            $profileName = $profileName.Substring(0, 44) + "..."
                        }
                        
                        $profileId = $profile.id
                        if ($profileId.Length -gt 37) {
                            $profileId = $profileId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $profile.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $profileName, $profileId, $assignment
                        if ($assignment -eq "Excluded") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }
                
                Write-Host "`n------- Endpoint Security - Firewall Profiles -------" -ForegroundColor Cyan
                if ($relevantPolicies.FirewallProfiles.Count -eq 0) {
                    Write-Host "No Firewall Profiles found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "Profile Name", "Profile ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($profile in $relevantPolicies.FirewallProfiles) {
                        $profileName = if (-not [string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.displayName } elseif (-not [string]::IsNullOrWhiteSpace($profile.name)) { $profile.name } else { "Unnamed Profile" }
                        if ($profileName.Length -gt 47) {
                            $profileName = $profileName.Substring(0, 44) + "..."
                        }
                        
                        $profileId = $profile.id
                        if ($profileId.Length -gt 37) {
                            $profileId = $profileId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $profile.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $profileName, $profileId, $assignment
                        if ($assignment -eq "Excluded") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }
                
                Write-Host "`n------- Endpoint Security - Endpoint Detection and Response Profiles -------" -ForegroundColor Cyan
                if ($relevantPolicies.EndpointDetectionProfiles.Count -eq 0) {
                    Write-Host "No Endpoint Detection and Response Profiles found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "Profile Name", "Profile ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($profile in $relevantPolicies.EndpointDetectionProfiles) {
                        $profileName = if (-not [string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.displayName } elseif (-not [string]::IsNullOrWhiteSpace($profile.name)) { $profile.name } else { "Unnamed Profile" }
                        if ($profileName.Length -gt 47) {
                            $profileName = $profileName.Substring(0, 44) + "..."
                        }
                        
                        $profileId = $profile.id
                        if ($profileId.Length -gt 37) {
                            $profileId = $profileId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $profile.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $profileName, $profileId, $assignment
                        if ($assignment -eq "Excluded") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }
                
                Write-Host "`n------- Endpoint Security - Attack Surface Reduction Profiles -------" -ForegroundColor Cyan
                if ($relevantPolicies.AttackSurfaceProfiles.Count -eq 0) {
                    Write-Host "No Attack Surface Reduction Profiles found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "Profile Name", "Profile ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($profile in $relevantPolicies.AttackSurfaceProfiles) {
                        $profileName = if (-not [string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.displayName } elseif (-not [string]::IsNullOrWhiteSpace($profile.name)) { $profile.name } else { "Unnamed Profile" }
                        if ($profileName.Length -gt 47) {
                            $profileName = $profileName.Substring(0, 44) + "..."
                        }
                        
                        $profileId = $profile.id
                        if ($profileId.Length -gt 37) {
                            $profileId = $profileId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $profile.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $profileName, $profileId, $assignment
                        if ($assignment -eq "Excluded") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }

                Write-Host "`n------- Windows 365 Cloud PC Provisioning Policies -------" -ForegroundColor Cyan
                if ($relevantPolicies.CloudPCProvisioningPolicies.Count -eq 0) {
                    Write-Host "No Windows 365 Cloud PC Provisioning Policies found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "Policy Name", "Policy ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($policy in $relevantPolicies.CloudPCProvisioningPolicies) {
                        $policyName = if (-not [string]::IsNullOrWhiteSpace($policy.displayName)) { $policy.displayName } elseif (-not [string]::IsNullOrWhiteSpace($policy.name)) { $policy.name } else { "Unnamed Policy" }
                        if ($policyName.Length -gt 47) {
                            $policyName = $policyName.Substring(0, 44) + "..."
                        }
                        
                        $policyId = $policy.id
                        if ($policyId.Length -gt 37) {
                            $policyId = $policyId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $policy.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $policyName, $policyId, $assignment
                        if ($assignment -eq "Excluded") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }

                Write-Host "`n------- Windows 365 Cloud PC User Settings -------" -ForegroundColor Cyan
                if ($relevantPolicies.CloudPCUserSettings.Count -eq 0) {
                    Write-Host "No Windows 365 Cloud PC User Settings found" -ForegroundColor Gray
                }
                else {
                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "Setting Name", "Setting ID", "Assignment"
                    $separator = "-" * 120
                    Write-Host $separator
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $separator
                    
                    foreach ($setting in $relevantPolicies.CloudPCUserSettings) {
                        $settingName = if (-not [string]::IsNullOrWhiteSpace($setting.displayName)) { $setting.displayName } elseif (-not [string]::IsNullOrWhiteSpace($setting.name)) { $setting.name } else { "Unnamed Setting" }
                        if ($settingName.Length -gt 47) {
                            $settingName = $settingName.Substring(0, 44) + "..."
                        }
                        
                        $settingId = $setting.id
                        if ($settingId.Length -gt 37) {
                            $settingId = $settingId.Substring(0, 34) + "..."
                        }
                        
                        $assignment = $setting.AssignmentReason
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $settingName, $settingId, $assignment
                        if ($assignment -eq "Excluded") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $separator
                }

                Add-ExportData -ExportData $exportData -Category "User" -Items @([PSCustomObject]@{
                        displayName      = $upn
                        id               = $userInfo.Id
                        AssignmentReason = "N/A"
                    }

                    Add-ExportData -ExportData $exportData -Category "Device Configuration" -Items $relevantPolicies.DeviceConfigs -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Settings Catalog Policy" -Items $relevantPolicies.SettingsCatalog -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Administrative Template" -Items $relevantPolicies.AdminTemplates -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Compliance Policy" -Items $relevantPolicies.CompliancePolicies -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "App Protection Policy" -Items $relevantPolicies.AppProtectionPolicies -AssignmentReason { param($item) $item.AssignmentSummary }
                    Add-ExportData -ExportData $exportData -Category "App Configuration Policy" -Items $relevantPolicies.AppConfigurationPolicies -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Platform Scripts" -Items $relevantPolicies.PlatformScripts -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Proactive Remediation Scripts" -Items $relevantPolicies.HealthScripts -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Autopilot Deployment Profile" -Items $relevantPolicies.DeploymentProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Enrollment Status Page" -Items $relevantPolicies.ESPProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Windows 365 Cloud PC Provisioning Policy" -Items $relevantPolicies.CloudPCProvisioningPolicies -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Windows 365 Cloud PC User Setting" -Items $relevantPolicies.CloudPCUserSettings -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Antivirus" -Items $relevantPolicies.AntivirusProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Disk Encryption" -Items $relevantPolicies.DiskEncryptionProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Firewall" -Items $relevantPolicies.FirewallProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - EDR" -Items $relevantPolicies.EndpointDetectionProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - ASR" -Items $relevantPolicies.AttackSurfaceProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                )
            }

            Export-ResultsIfRequested -ExportData $exportData -DefaultFileName "IntuneUserAssignments.csv" -ForceExport:$ExportToCSV -CustomExportPath $ExportPath
        }
        '2' {
            Write-Host "Group selection chosen" -ForegroundColor Green

            Write-Host "Please enter Group names or Object IDs, separated by commas (,): " -ForegroundColor Cyan
            Write-Host "Example: 'Marketing Team, 12345678-1234-1234-1234-123456789012'" -ForegroundColor Gray
            $groupInput = Read-Host

            if ([string]::IsNullOrWhiteSpace($groupInput)) {
                Write-Host "No group information provided. Please try again." -ForegroundColor Red
                continue
            }

            $groupInputs = $groupInput -split ',' | ForEach-Object { $_.Trim() }
            $exportData = [System.Collections.ArrayList]::new()

            foreach ($input in $groupInputs) {
                Write-Host "`nProcessing input: $input" -ForegroundColor Yellow

                $groupId = $null
                $groupName = $null

                if ($input -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
                    $groupInfo = Get-GroupInfo -GroupId $input
                    if (-not $groupInfo.Success) {
                        Write-Host "No group found with ID: $input" -ForegroundColor Red
                        continue
                    }
                    $groupId = $groupInfo.Id
                    $groupName = $groupInfo.DisplayName
                }
                else {
                    $groupUri = "$GraphEndpoint/v1.0/groups?`$filter=displayName eq '$input'"
                    $groupResponse = Invoke-MgGraphRequest -Uri $groupUri -Method Get

                    if ($groupResponse.value.Count -eq 0) {
                        Write-Host "No group found with name: $input" -ForegroundColor Red
                        continue
                    }
                    elseif ($groupResponse.value.Count -gt 1) {
                        Write-Host "Multiple groups found with name: $input. Please use the Object ID instead:" -ForegroundColor Red
                        foreach ($group in $groupResponse.value) {
                            Write-Host "  - $($group.displayName) (ID: $($group.id))" -ForegroundColor Yellow
                        }
                        continue
                    }

                    $groupId = $groupResponse.value[0].id
                    $groupName = $groupResponse.value[0].displayName
                }

                Write-Host "Found group: $groupName (ID: $groupId)" -ForegroundColor Green
                Write-Host "Fetching Intune Profiles and Applications for the group ... (this takes a few seconds)" -ForegroundColor Yellow

                $relevantPolicies = @{
                    DeviceConfigs               = @()
                    SettingsCatalog             = @()
                    AdminTemplates              = @()
                    CompliancePolicies          = @()
                    AppProtectionPolicies       = @()
                    AppConfigurationPolicies    = @()
                    AppsRequired                = @()
                    AppsAvailable               = @()
                    AppsUninstall               = @()
                    PlatformScripts             = @()
                    HealthScripts               = @()
                    AntivirusProfiles           = @()
                    DiskEncryptionProfiles      = @()
                    FirewallProfiles            = @()
                    EndpointDetectionProfiles   = @()
                    AttackSurfaceProfiles       = @()
                    DeploymentProfiles          = @()
                    ESPProfiles                 = @()
                    CloudPCProvisioningPolicies = @()
                    CloudPCUserSettings         = @()
                }

                Write-Host "Fetching Device Configurations..." -ForegroundColor Yellow
                $deviceConfigs = Get-IntuneEntities -EntityType "deviceConfigurations"
                foreach ($config in $deviceConfigs) {
                    $directAssignments = Get-IntuneAssignments -EntityType "deviceConfigurations" -EntityId $config.id -GroupId $groupId
                    if ($directAssignments.Count -gt 0) {
                        $assignmentReason = $directAssignments[0].Reason
                        if ($assignmentReason -eq "Direct Assignment" -or $assignmentReason -eq "Direct Exclusion") {
                            $config | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentReason -Force
                            $relevantPolicies.DeviceConfigs += $config
                        }
                    }
                }

                Write-Host "Fetching Settings Catalog Policies..." -ForegroundColor Yellow
                $settingsCatalog = Get-IntuneEntities -EntityType "configurationPolicies"
                foreach ($policy in $settingsCatalog) {
                    if ($policy.templateReference -and $policy.templateReference.templateFamily -like "endpointSecurity*") {
                        continue
                    }
                    $directAssignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id -GroupId $groupId
                    if ($directAssignments.Count -gt 0) {
                        $assignmentReason = $directAssignments[0].Reason
                        if ($assignmentReason -eq "Direct Assignment" -or $assignmentReason -eq "Direct Exclusion") {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentReason -Force
                            $relevantPolicies.SettingsCatalog += $policy
                        }
                    }
                }

                Write-Host "Fetching Administrative Templates..." -ForegroundColor Yellow
                $adminTemplates = Get-IntuneEntities -EntityType "groupPolicyConfigurations"
                foreach ($template in $adminTemplates) {
                    $directAssignments = Get-IntuneAssignments -EntityType "groupPolicyConfigurations" -EntityId $template.id -GroupId $groupId
                    if ($directAssignments.Count -gt 0) {
                        $assignmentReason = $directAssignments[0].Reason
                        if ($assignmentReason -eq "Direct Assignment" -or $assignmentReason -eq "Direct Exclusion") {
                            $template | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentReason -Force
                            $relevantPolicies.AdminTemplates += $template
                        }
                    }
                }

                Write-Host "Fetching Compliance Policies..." -ForegroundColor Yellow
                $compliancePolicies = Get-IntuneEntities -EntityType "deviceCompliancePolicies"
                foreach ($policy in $compliancePolicies) {
                    $directAssignments = Get-IntuneAssignments -EntityType "deviceCompliancePolicies" -EntityId $policy.id -GroupId $groupId
                    if ($directAssignments.Count -gt 0) {
                        $assignmentReason = $directAssignments[0].Reason
                        if ($assignmentReason -eq "Direct Assignment" -or $assignmentReason -eq "Direct Exclusion") {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentReason -Force
                            $relevantPolicies.CompliancePolicies += $policy
                        }
                    }
                }

                Write-Host "Fetching App Protection Policies..." -ForegroundColor Yellow
                $appProtectionPolicies = Get-IntuneEntities -EntityType "deviceAppManagement/managedAppPolicies"
                foreach ($policy in $appProtectionPolicies) {
                    $policyType = $policy.'@odata.type'
                    $assignmentsUri = switch ($policyType) {
                        "#microsoft.graph.androidManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/androidManagedAppProtections('$($policy.id)')/assignments" }
                        "#microsoft.graph.iosManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/iosManagedAppProtections('$($policy.id)')/assignments" }
                        "#microsoft.graph.windowsManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/windowsManagedAppProtections('$($policy.id)')/assignments" }
                        default { $null }
                    }

                    if ($assignmentsUri) {
                        try {
                            $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get
                            $directAssignments = Get-IntuneAssignments -EntityType "deviceAppManagement/managedAppPolicies" -EntityId $policy.id -GroupId $groupId
                            if ($directAssignments.Count -gt 0) {
                                $assignmentReason = $directAssignments[0].Reason
                                if ($assignmentReason -eq "Direct Assignment" -or $assignmentReason -eq "Direct Exclusion") {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentReason -Force
                                    $relevantPolicies.AppProtectionPolicies += $policy
                                }
                            }
                        }
                        catch {
                            Write-Host "Error fetching assignments for policy $($policy.displayName): $($_.Exception.Message)" -ForegroundColor Red
                        }
                    }
                }

                Write-Host "Fetching App Configuration Policies..." -ForegroundColor Yellow
                $appConfigPolicies = Get-IntuneEntities -EntityType "deviceAppManagement/mobileAppConfigurations"
                foreach ($policy in $appConfigPolicies) {
                    $directAssignments = Get-IntuneAssignments -EntityType "mobileAppConfigurations" -EntityId $policy.id -GroupId $groupId
                    if ($directAssignments.Count -gt 0) {
                        $assignmentReason = $directAssignments[0].Reason
                        if ($assignmentReason -eq "Direct Assignment" -or $assignmentReason -eq "Direct Exclusion") {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentReason -Force
                            $relevantPolicies.AppConfigurationPolicies += $policy
                        }
                    }
                }

                Write-Host "Fetching Antivirus Policies for group..." -ForegroundColor Yellow
                $antivirusPoliciesFoundGroup = [System.Collections.ArrayList]::new()
                $processedAntivirusIdsGroup = [System.Collections.Generic.HashSet[string]]::new()

                $configPoliciesForAntivirusGroup = Get-IntuneEntities -EntityType "configurationPolicies"
                $matchingConfigPoliciesAntivirusGroup = $configPoliciesForAntivirusGroup | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAntivirus' }

                $endpointSecurityCategories = @(
                    @{ Name = "Antivirus"; Key = "AntivirusProfiles"; TemplateFamily = "endpointSecurityAntivirus"; UserFriendlyType = "Antivirus Profile" },
                    @{ Name = "Disk Encryption"; Key = "DiskEncryptionProfiles"; TemplateFamily = "endpointSecurityDiskEncryption"; UserFriendlyType = "Disk Encryption Profile" },
                    @{ Name = "Firewall"; Key = "FirewallProfiles"; TemplateFamily = "endpointSecurityFirewall"; UserFriendlyType = "Firewall Profile" },
                    @{ Name = "Endpoint Detection and Response"; Key = "EndpointDetectionProfiles"; TemplateFamily = "endpointSecurityEndpointDetectionAndResponse"; UserFriendlyType = "EDR Profile" },
                    @{ Name = "Attack Surface Reduction"; Key = "AttackSurfaceProfiles"; TemplateFamily = "endpointSecurityAttackSurfaceReductionRules"; UserFriendlyType = "ASR Profile" }
                )

                foreach ($esCategory in $endpointSecurityCategories) {
                    Write-Host "Fetching $($esCategory.Name) Policies for group..." -ForegroundColor Yellow
                    $processedEsPolicyIds = [System.Collections.Generic.HashSet[string]]::new() # Track IDs per category to avoid duplicates from configPolicies and intents

                    $allConfigEsPolicies = Get-IntuneEntities -EntityType "configurationPolicies" # Fetch all, then filter
                    $matchingConfigEsPolicies = $allConfigEsPolicies | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq $esCategory.TemplateFamily }
                    if ($matchingConfigEsPolicies) {
                        foreach ($policy in $matchingConfigEsPolicies) {
                            if ($processedEsPolicyIds.Add($policy.id)) {
                                $directAssignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id -GroupId $groupId
                                if ($directAssignments.Count -gt 0) {
                                    $assignmentReason = $directAssignments[0].Reason
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentReason -Force
                                    $relevantPolicies[$esCategory.Key] += $policy
                                }
                            }
                        }
                    }

                    $allIntentEsPolicies = Get-IntuneEntities -EntityType "deviceManagement/intents" # Fetch all, then filter
                    $matchingIntentEsPolicies = $allIntentEsPolicies | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq $esCategory.TemplateFamily }
                    if ($matchingIntentEsPolicies) {
                        foreach ($policy in $matchingIntentEsPolicies) {
                            if ($processedEsPolicyIds.Add($policy.id)) {
                                try {
                                    $allIntentAssignments = [System.Collections.ArrayList]::new()
                                    $currentIntentAssignmentsUri = "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments"
                                    do {
                                        $intentAssignmentsResponsePage = Invoke-MgGraphRequest -Uri $currentIntentAssignmentsUri -Method Get
                                        if ($intentAssignmentsResponsePage -and $null -ne $intentAssignmentsResponsePage.value) {
                                            $allIntentAssignments.AddRange($intentAssignmentsResponsePage.value)
                                        }
                                        $currentIntentAssignmentsUri = $intentAssignmentsResponsePage.'@odata.nextLink'
                                    } while (![string]::IsNullOrEmpty($currentIntentAssignmentsUri))

                                    $directGroupAssignment = $allIntentAssignments | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget' -and $_.target.groupId -eq $groupId }
                                    $directGroupExclusion = $allIntentAssignments | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.exclusionGroupAssignmentTarget' -and $_.target.groupId -eq $groupId }

                                    $assignmentReason = $null
                                    if ($directGroupExclusion) {
                                        $assignmentReason = "Direct Exclusion"
                                    }
                                    elseif ($directGroupAssignment) {
                                        $assignmentReason = "Direct Assignment"
                                    }

                                    if ($assignmentReason) {
                                        $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentReason -Force
                                        $relevantPolicies[$esCategory.Key] += $policy
                                    }
                                }
                                catch {
                                    Write-Warning "Error fetching assignments for ES Intent $($policy.displayName) (ID: $($policy.id)): $($_.Exception.Message)"
                                }
                            }
                        }
                    }
                }

                Write-Host "Fetching Applications..." -ForegroundColor Yellow
                $appUri = "$GraphEndpoint/beta/deviceAppManagement/mobileApps?`$filter=isAssigned eq true"
                $appResponse = Invoke-MgGraphRequest -Uri $appUri -Method Get
                $allApps = $appResponse.value
                while ($appResponse.'@odata.nextLink') {
                    $appResponse = Invoke-MgGraphRequest -Uri $appResponse.'@odata.nextLink' -Method Get
                    $allApps += $appResponse.value
                }
                $totalApps = $allApps.Count

                foreach ($app in $allApps) {
                    if ($app.isFeatured -or $app.isBuiltIn) {
                        continue
                    }

                    $appId = $app.id
                    $allAppAssignments = [System.Collections.ArrayList]::new()
                    $currentAppAssignmentsUri = "$GraphEndpoint/beta/deviceAppManagement/mobileApps('$appId')/assignments"
                    do {
                        $appAssignmentsResponsePage = Invoke-MgGraphRequest -Uri $currentAppAssignmentsUri -Method Get
                        if ($appAssignmentsResponsePage -and $null -ne $appAssignmentsResponsePage.value) {
                            $allAppAssignments.AddRange($appAssignmentsResponsePage.value)
                        }
                        $currentAppAssignmentsUri = $appAssignmentsResponsePage.'@odata.nextLink'
                    } while (![string]::IsNullOrEmpty($currentAppAssignmentsUri))
                    
                    $relevantAppAssignmentReason = $null
                    $intentForGroup = $null

                    foreach ($assignmentItem in $allAppAssignments) {
                        if ($assignmentItem.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget' -and $assignmentItem.target.groupId -eq $groupId) {
                            $relevantAppAssignmentReason = "Direct Assignment"
                            $intentForGroup = $assignmentItem.intent
                            break
                        }
                        elseif ($assignmentItem.target.'@odata.type' -eq '#microsoft.graph.exclusionGroupAssignmentTarget' -and $assignmentItem.target.groupId -eq $groupId) {
                            $relevantAppAssignmentReason = "Group Exclusion"
                            $intentForGroup = $assignmentItem.intent # Intent might still be relevant for excluded apps
                            break
                        }
                    }

                    if ($relevantAppAssignmentReason) {
                        $appWithReason = $app.PSObject.Copy()
                        $appWithReason | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $relevantAppAssignmentReason -Force
                        if ($intentForGroup) {
                            switch ($intentForGroup) {
                                "required" { $relevantPolicies.AppsRequired += $appWithReason }
                                "available" { $relevantPolicies.AppsAvailable += $appWithReason }
                                "uninstall" { $relevantPolicies.AppsUninstall += $appWithReason }
                            }
                        }
                    }
                }

                Write-Host "Fetching Platform Scripts..." -ForegroundColor Yellow
                $platformScripts = Get-IntuneEntities -EntityType "deviceManagementScripts"
                foreach ($script in $platformScripts) {
                    $directAssignments = Get-IntuneAssignments -EntityType "deviceManagementScripts" -EntityId $script.id -GroupId $groupId
                    if ($directAssignments.Count -gt 0) {
                        $assignmentReason = $directAssignments[0].Reason
                        if ($assignmentReason -eq "Direct Assignment" -or $assignmentReason -eq "Direct Exclusion") {
                            $script | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentReason -Force
                            $relevantPolicies.PlatformScripts += $script
                        }
                    }
                }

                Write-Host "Fetching Proactive Remediation Scripts..." -ForegroundColor Yellow
                $healthScripts = Get-IntuneEntities -EntityType "deviceHealthScripts"
                foreach ($script in $healthScripts) {
                    $directAssignments = Get-IntuneAssignments -EntityType "deviceHealthScripts" -EntityId $script.id -GroupId $groupId
                    if ($directAssignments.Count -gt 0) {
                        $assignmentReason = $directAssignments[0].Reason
                        if ($assignmentReason -eq "Direct Assignment" -or $assignmentReason -eq "Direct Exclusion") {
                            $script | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentReason -Force
                            $relevantPolicies.HealthScripts += $script
                        }
                    }
                }

                Write-Host "Fetching Autopilot Deployment Profiles..." -ForegroundColor Yellow
                $autoProfiles = Get-IntuneEntities -EntityType "windowsAutopilotDeploymentProfiles"
                foreach ($profile in $autoProfiles) {
                    $directAssignments = Get-IntuneAssignments -EntityType "windowsAutopilotDeploymentProfiles" -EntityId $profile.id -GroupId $groupId
                    if ($directAssignments.Count -gt 0) {
                        $assignmentReason = $directAssignments[0].Reason
                        if ($assignmentReason -eq "Direct Assignment" -or $assignmentReason -eq "Direct Exclusion") {
                            $profile | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentReason -Force
                            $relevantPolicies.DeploymentProfiles += $profile
                        }
                    }
                }

                Write-Host "Fetching Enrollment Status Page Profiles..." -ForegroundColor Yellow
                $enrollmentConfigs = Get-IntuneEntities -EntityType "deviceEnrollmentConfigurations"
                $espProfiles = $enrollmentConfigs | Where-Object { $_.'@odata.type' -match 'EnrollmentCompletionPageConfiguration' }
                foreach ($esp in $espProfiles) {
                    $directAssignments = Get-IntuneAssignments -EntityType "deviceEnrollmentConfigurations" -EntityId $esp.id -GroupId $groupId
                    if ($directAssignments.Count -gt 0) {
                        $assignmentReason = $directAssignments[0].Reason
                        if ($assignmentReason -eq "Direct Assignment" -or $assignmentReason -eq "Direct Exclusion") {
                            $esp | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentReason -Force
                            $relevantPolicies.ESPProfiles += $esp
                        }
                    }
                }

                Write-Host "Fetching Windows 365 Cloud PC Provisioning Policies..." -ForegroundColor Yellow
                try {
                    $cloudPCProvisioningPolicies = Get-IntuneEntities -EntityType "virtualEndpoint/provisioningPolicies"
                    foreach ($policy in $cloudPCProvisioningPolicies) {
                        $directAssignments = Get-IntuneAssignments -EntityType "virtualEndpoint/provisioningPolicies" -EntityId $policy.id -GroupId $groupId
                        if ($directAssignments.Count -gt 0) {
                            $assignmentReason = $directAssignments[0].Reason
                            if ($assignmentReason -eq "Direct Assignment" -or $assignmentReason -eq "Direct Exclusion") {
                                $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentReason -Force
                                $relevantPolicies.CloudPCProvisioningPolicies += $policy
                            }
                        }
                    }
                }
                catch {
                    Write-Warning "Unable to fetch Windows 365 Cloud PC Provisioning Policies: $($_.Exception.Message)"
                }

                Write-Host "Fetching Windows 365 Cloud PC User Settings..." -ForegroundColor Yellow
                try {
                    $cloudPCUserSettings = Get-IntuneEntities -EntityType "virtualEndpoint/userSettings"
                    foreach ($setting in $cloudPCUserSettings) {
                        $directAssignments = Get-IntuneAssignments -EntityType "virtualEndpoint/userSettings" -EntityId $setting.id -GroupId $groupId
                        if ($directAssignments.Count -gt 0) {
                            $assignmentReason = $directAssignments[0].Reason
                            if ($assignmentReason -eq "Direct Assignment" -or $assignmentReason -eq "Direct Exclusion") {
                                $setting | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentReason -Force
                                $relevantPolicies.CloudPCUserSettings += $setting
                            }
                        }
                    }
                }
                catch {
                    Write-Warning "Unable to fetch Windows 365 Cloud PC User Settings: $($_.Exception.Message)"
                }

                function Format-PolicyTable {
                    param (
                        [string]$Title,
                        [object[]]$Policies,
                        [scriptblock]$GetName,
                        [scriptblock]$GetExtra = { param($p) "" }
                    )
                    $localTableSeparator = "-" * 120 # Use a local variable for separator

                    $headerSeparator = "-" * ($Title.Length + 16)
                    Write-Host "`n$headerSeparator" -ForegroundColor Cyan
                    Write-Host "------- $Title -------" -ForegroundColor Cyan
                    Write-Host "$headerSeparator" -ForegroundColor Cyan
                    
                    if ($Policies.Count -eq 0) {
                        Write-Host "No $Title found for this group." -ForegroundColor Gray
                        Write-Host $localTableSeparator -ForegroundColor Gray
                        Write-Host ""
                        return
                    }

                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "Policy Name", "ID", "Assignment"
                    
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $localTableSeparator -ForegroundColor Gray
                    
                    foreach ($policy in $Policies) {
                        $name = & $GetName $policy
                        $extra = & $GetExtra $policy
                        
                        if ($name.Length -gt 47) { $name = $name.Substring(0, 44) + "..." }
                        
                        $id = $policy.id
                        if ($id.Length -gt 37) { $id = $id.Substring(0, 34) + "..." }
                        
                        $assignment = if ($policy.AssignmentReason) { $policy.AssignmentReason } else { "N/A" }
                        if ($assignment.Length -gt 27) { $assignment = $assignment.Substring(0, 24) + "..." }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $name, $id, $assignment
                        if ($assignment -eq "Direct Exclusion") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    Write-Host $localTableSeparator -ForegroundColor Gray
                }

                Format-PolicyTable -Title "Device Configurations" -Policies $relevantPolicies.DeviceConfigs -GetName {
                    param($config)
                    if ([string]::IsNullOrWhiteSpace($config.name)) { $config.displayName } else { $config.name }
                }

                Format-PolicyTable -Title "Settings Catalog Policies" -Policies $relevantPolicies.SettingsCatalog -GetName {
                    param($policy)
                    if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                }

                Format-PolicyTable -Title "Administrative Templates" -Policies $relevantPolicies.AdminTemplates -GetName {
                    param($template)
                    if ([string]::IsNullOrWhiteSpace($template.name)) { $template.displayName } else { $template.name }
                }

                Format-PolicyTable -Title "Compliance Policies" -Policies $relevantPolicies.CompliancePolicies -GetName {
                    param($policy)
                    if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                }

                Format-PolicyTable -Title "App Protection Policies" -Policies $relevantPolicies.AppProtectionPolicies -GetName {
                    param($policy)
                    $policy.displayName
                } -GetExtra {
                    param($policy)
                    @{
                        Label = 'Platform'
                        Value = switch ($policy.'@odata.type') {
                            "#microsoft.graph.androidManagedAppProtection" { "Android" }
                            "#microsoft.graph.iosManagedAppProtection" { "iOS" }
                            "#microsoft.graph.windowsManagedAppProtection" { "Windows" }
                            default { "Unknown" }
                        }
                    }
                }

                Format-PolicyTable -Title "App Configuration Policies" -Policies $relevantPolicies.AppConfigurationPolicies -GetName {
                    param($policy)
                    if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                }

                Format-PolicyTable -Title "Platform Scripts" -Policies $relevantPolicies.PlatformScripts -GetName {
                    param($script)
                    if ([string]::IsNullOrWhiteSpace($script.name)) { $script.displayName } else { $script.name }
                }

                Format-PolicyTable -Title "Proactive Remediation Scripts" -Policies $relevantPolicies.HealthScripts -GetName {
                    param($script)
                    if ([string]::IsNullOrWhiteSpace($script.name)) { $script.displayName } else { $script.name }
                }

                Format-PolicyTable -Title "Autopilot Deployment Profiles" -Policies $relevantPolicies.DeploymentProfiles -GetName {
                    param($profile)
                    if ([string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.name } else { $profile.displayName }
                }

                Format-PolicyTable -Title "Enrollment Status Page Profiles" -Policies $relevantPolicies.ESPProfiles -GetName {
                    param($profile)
                    if ([string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.name } else { $profile.displayName }
                }

                Format-PolicyTable -Title "Windows 365 Cloud PC Provisioning Policies" -Policies $relevantPolicies.CloudPCProvisioningPolicies -GetName {
                    param($policy)
                    if ([string]::IsNullOrWhiteSpace($policy.displayName)) { $policy.name } else { $policy.displayName }
                }

                Format-PolicyTable -Title "Windows 365 Cloud PC User Settings" -Policies $relevantPolicies.CloudPCUserSettings -GetName {
                    param($setting)
                    if ([string]::IsNullOrWhiteSpace($setting.displayName)) { $setting.name } else { $setting.displayName }
                }

                Format-PolicyTable -Title "Required Apps" -Policies $relevantPolicies.AppsRequired -GetName {
                    param($app)
                    $app.displayName
                }

                Format-PolicyTable -Title "Available Apps" -Policies $relevantPolicies.AppsAvailable -GetName {
                    param($app)
                    $app.displayName
                }

                Format-PolicyTable -Title "Uninstall Apps" -Policies $relevantPolicies.AppsUninstall -GetName {
                    param($app)
                    $app.displayName
                }

                Format-PolicyTable -Title "Endpoint Security - Antivirus Profiles" -Policies $relevantPolicies.AntivirusProfiles -GetName { param($profile) if (-not [string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.displayName } elseif (-not [string]::IsNullOrWhiteSpace($profile.name)) { $profile.name } else { "Unnamed Profile" } }
                
                Format-PolicyTable -Title "Endpoint Security - Disk Encryption Profiles" -Policies $relevantPolicies.DiskEncryptionProfiles -GetName { param($profile) if (-not [string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.displayName } elseif (-not [string]::IsNullOrWhiteSpace($profile.name)) { $profile.name } else { "Unnamed Profile" } }
                
                Format-PolicyTable -Title "Endpoint Security - Firewall Profiles" -Policies $relevantPolicies.FirewallProfiles -GetName { param($profile) if (-not [string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.displayName } elseif (-not [string]::IsNullOrWhiteSpace($profile.name)) { $profile.name } else { "Unnamed Profile" } }
                
                Format-PolicyTable -Title "Endpoint Security - EDR Profiles" -Policies $relevantPolicies.EndpointDetectionProfiles -GetName { param($profile) if (-not [string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.displayName } elseif (-not [string]::IsNullOrWhiteSpace($profile.name)) { $profile.name } else { "Unnamed Profile" } }
                
                Format-PolicyTable -Title "Endpoint Security - ASR Profiles" -Policies $relevantPolicies.AttackSurfaceProfiles -GetName { param($profile) if (-not [string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.displayName } elseif (-not [string]::IsNullOrWhiteSpace($profile.name)) { $profile.name } else { "Unnamed Profile" } }

                Add-ExportData -ExportData $exportData -Category "Device" -Items @([PSCustomObject]@{
                        displayName      = $deviceName
                        id               = $deviceInfo.Id
                        AssignmentReason = "N/A"
                    }

                    Add-ExportData -ExportData $exportData -Category "Device Configuration" -Items $relevantPolicies.DeviceConfigs -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Settings Catalog Policy" -Items $relevantPolicies.SettingsCatalog -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Administrative Template" -Items $relevantPolicies.AdminTemplates -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Compliance Policy" -Items $relevantPolicies.CompliancePolicies -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "App Protection Policy" -Items $relevantPolicies.AppProtectionPolicies -AssignmentReason { param($item) $item.AssignmentSummary }
                    Add-ExportData -ExportData $exportData -Category "App Configuration Policy" -Items $relevantPolicies.AppConfigurationPolicies -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Platform Scripts" -Items $relevantPolicies.PlatformScripts -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Proactive Remediation Scripts" -Items $relevantPolicies.HealthScripts -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Autopilot Deployment Profile" -Items $relevantPolicies.DeploymentProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Enrollment Status Page" -Items $relevantPolicies.ESPProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Windows 365 Cloud PC Provisioning Policy" -Items $relevantPolicies.CloudPCProvisioningPolicies -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Windows 365 Cloud PC User Setting" -Items $relevantPolicies.CloudPCUserSettings -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Antivirus" -Items $relevantPolicies.AntivirusProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Disk Encryption" -Items $relevantPolicies.DiskEncryptionProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Firewall" -Items $relevantPolicies.FirewallProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - EDR" -Items $relevantPolicies.EndpointDetectionProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - ASR" -Items $relevantPolicies.AttackSurfaceProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                )
            }

            Export-ResultsIfRequested -ExportData $exportData -DefaultFileName "IntuneDeviceAssignments.csv" -ForceExport:$ExportToCSV -CustomExportPath $ExportPath
        }
        '3' {
            Write-Host "Device selection chosen" -ForegroundColor Green

            if ($parameterMode -and $DeviceNames) {
                $deviceInput = $DeviceNames
            }
            else {
                Write-Host "Please enter Device Name(s), separated by commas (,): " -ForegroundColor Cyan
                $deviceInput = Read-Host
            }

            if ([string]::IsNullOrWhiteSpace($deviceInput)) {
                Write-Host "No device name provided. Please try again." -ForegroundColor Red
                if ($parameterMode) { exit 1 } else { continue }
            }

            $deviceNames = $deviceInput -split ',' | ForEach-Object { $_.Trim() }
            $exportData = [System.Collections.ArrayList]::new()

            foreach ($deviceName in $deviceNames) {
                Write-Host "`nProcessing device: $deviceName" -ForegroundColor Yellow

                $deviceInfo = Get-DeviceInfo -DeviceName $deviceName
                if (-not $deviceInfo.Success) {
                    Write-Host "Device not found: $deviceName" -ForegroundColor Red
                    Write-Host "Please verify the device name is correct." -ForegroundColor Yellow
                    continue
                }

                try {
                    $groupMemberships = Get-GroupMemberships -ObjectId $deviceInfo.Id -ObjectType "Device"
                    Write-Host "Device Group Memberships: $($groupMemberships.displayName -join ', ')" -ForegroundColor Green
                }
                catch {
                    Write-Host "Error fetching group memberships for device: $deviceName" -ForegroundColor Red
                    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
                    continue
                }

                Write-Host "Fetching Intune Profiles and Applications for the device ... (this takes a few seconds)" -ForegroundColor Yellow

                $relevantPolicies = @{
                    DeviceConfigs               = @()
                    SettingsCatalog             = @()
                    AdminTemplates              = @()
                    CompliancePolicies          = @()
                    AppProtectionPolicies       = @()
                    AppConfigurationPolicies    = @()
                    AppsRequired                = @()
                    AppsAvailable               = @()
                    AppsUninstall               = @()
                    PlatformScripts             = @()
                    HealthScripts               = @()
                    AntivirusProfiles           = @()
                    DiskEncryptionProfiles      = @()
                    FirewallProfiles            = @()
                    EndpointDetectionProfiles   = @()
                    AttackSurfaceProfiles       = @()
                    CloudPCProvisioningPolicies = @()
                    CloudPCUserSettings         = @()
                }

                Write-Host "Fetching Device Configurations..." -ForegroundColor Yellow
                $deviceConfigs = Get-IntuneEntities -EntityType "deviceConfigurations"
                foreach ($config in $deviceConfigs) {
                    $assignments = Get-IntuneAssignments -EntityType "deviceConfigurations" -EntityId $config.id
                    foreach ($assignment in $assignments) {
                        if ($assignment.Reason -ne "All Users" -and
                            ($assignment.Reason -eq "All Devices" -or
                            ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId))) {
                            $config | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                            $relevantPolicies.DeviceConfigs += $config
                            break
                        }
                    }
                }

                Write-Host "Fetching Settings Catalog Policies..." -ForegroundColor Yellow
                $settingsCatalog = Get-IntuneEntities -EntityType "configurationPolicies"
                foreach ($policy in $settingsCatalog) {
                    $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                    foreach ($assignment in $assignments) {
                        if ($assignment.Reason -ne "All Users" -and
                            ($assignment.Reason -eq "All Devices" -or
                            ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId))) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                            $relevantPolicies.SettingsCatalog += $policy
                            break
                        }
                    }
                }

                Write-Host "Fetching Administrative Templates..." -ForegroundColor Yellow
                $adminTemplates = Get-IntuneEntities -EntityType "groupPolicyConfigurations"
                foreach ($template in $adminTemplates) {
                    $assignments = Get-IntuneAssignments -EntityType "groupPolicyConfigurations" -EntityId $template.id
                    foreach ($assignment in $assignments) {
                        if ($assignment.Reason -ne "All Users" -and
                            ($assignment.Reason -eq "All Devices" -or
                            ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId))) {
                            $template | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                            $relevantPolicies.AdminTemplates += $template
                            break
                        }
                    }
                }

                Write-Host "Fetching Compliance Policies..." -ForegroundColor Yellow
                $compliancePolicies = Get-IntuneEntities -EntityType "deviceCompliancePolicies"
                foreach ($policy in $compliancePolicies) {
                    $assignments = Get-IntuneAssignments -EntityType "deviceCompliancePolicies" -EntityId $policy.id
                    foreach ($assignment in $assignments) {
                        if ($assignment.Reason -ne "All Users" -and
                            ($assignment.Reason -eq "All Devices" -or
                            ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId))) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                            $relevantPolicies.CompliancePolicies += $policy
                            break
                        }
                    }
                }

                Write-Host "Fetching App Protection Policies..." -ForegroundColor Yellow
                $appProtectionPolicies = Get-IntuneEntities -EntityType "deviceAppManagement/managedAppPolicies"
                foreach ($policy in $appProtectionPolicies) {
                    $policyType = $policy.'@odata.type'
                    $assignmentsUri = switch ($policyType) {
                        "#microsoft.graph.androidManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/androidManagedAppProtections('$($policy.id)')/assignments" }
                        "#microsoft.graph.iosManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/iosManagedAppProtections('$($policy.id)')/assignments" }
                        "#microsoft.graph.windowsManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/windowsManagedAppProtections('$($policy.id)')/assignments" }
                        default { $null }
                    }

                    if ($assignmentsUri) {
                        try {
                            $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get
                            $assignments = @()
                            foreach ($assignment in $assignmentResponse.value) {
                                $assignmentReason = $null
                                switch ($assignment.target.'@odata.type') {
                                    '#microsoft.graph.allLicensedUsersAssignmentTarget' {
                                        $assignmentReason = "All Users"
                                    }
                                    '#microsoft.graph.allDevicesAssignmentTarget' {
                                        $assignmentReason = "All Devices"
                                    }
                                    '#microsoft.graph.groupAssignmentTarget' {
                                        if ($groupMemberships.id -contains $assignment.target.groupId) {
                                            $assignmentReason = "Group Assignment"
                                        }
                                    }
                                    '#microsoft.graph.exclusionGroupAssignmentTarget' {
                                        if ($groupMemberships.id -contains $assignment.target.groupId) {
                                            $assignmentReason = "Group Exclusion"
                                        }
                                    }
                                }

                                if ($assignmentReason -and $assignmentReason -ne "All Users") {
                                    $assignments += @{
                                        Reason  = $assignmentReason
                                        GroupId = $assignment.target.groupId
                                    }
                                }
                            }

                            if ($assignments.Count -gt 0) {
                                $assignmentSummary = $assignments | Where-Object { $_.Reason -ne "All Users" } | ForEach-Object {
                                    if ($_.Reason -eq "Group Assignment") {
                                        $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                                        "$($_.Reason) - $($groupInfo.DisplayName)"
                                    }
                                    else {
                                        $_.Reason
                                    }
                                }
                                $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                                $relevantPolicies.AppProtectionPolicies += $policy
                            }
                        }
                        catch {
                            Write-Host "Error fetching assignments for policy $($policy.displayName): $($_.Exception.Message)" -ForegroundColor Red
                        }
                    }
                }

                Write-Host "Fetching App Configuration Policies..." -ForegroundColor Yellow
                $appConfigPolicies = Get-IntuneEntities -EntityType "deviceAppManagement/mobileAppConfigurations"
                foreach ($policy in $appConfigPolicies) {
                    $assignments = Get-IntuneAssignments -EntityType "mobileAppConfigurations" -EntityId $policy.id
                    foreach ($assignment in $assignments) {
                        if ($assignment.Reason -ne "All Users" -and
                            ($assignment.Reason -eq "All Devices" -or
                            ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId))) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                            $relevantPolicies.AppConfigurationPolicies += $policy
                            break
                        }
                    }
                }

                Write-Host "Fetching Platform Scripts..." -ForegroundColor Yellow
                $platformScripts = Get-IntuneEntities -EntityType "deviceManagementScripts"
                foreach ($script in $platformScripts) {
                    $assignments = Get-IntuneAssignments -EntityType "deviceManagementScripts" -EntityId $script.id
                    foreach ($assignment in $assignments) {
                        if ($assignment.Reason -ne "All Users" -and
                            ($assignment.Reason -eq "All Devices" -or
                            ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId))) {
                            $script | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                            $relevantPolicies.PlatformScripts += $script
                            break
                        }
                    }
                }

                Write-Host "Fetching Proactive Remediation Scripts..." -ForegroundColor Yellow
                $healthScripts = Get-IntuneEntities -EntityType "deviceHealthScripts"
                foreach ($script in $healthScripts) {
                    $assignments = Get-IntuneAssignments -EntityType "deviceHealthScripts" -EntityId $script.id
                    foreach ($assignment in $assignments) {
                        if ($assignment.Reason -ne "All Users" -and
                            ($assignment.Reason -eq "All Devices" -or
                            ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId))) {
                            $script | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                            $relevantPolicies.HealthScripts += $script
                            break
                        }
                    }
                }

                Write-Host "Fetching Autopilot Deployment Profiles..." -ForegroundColor Yellow
                $autoProfiles = Get-IntuneEntities -EntityType "windowsAutopilotDeploymentProfiles"
                foreach ($profile in $autoProfiles) {
                    $assignments = Get-IntuneAssignments -EntityType "windowsAutopilotDeploymentProfiles" -EntityId $profile.id
                    foreach ($assignment in $assignments) {
                        if (($assignment.Reason -eq "All Devices") -or
                            ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId)) {
                            $profile | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                            $relevantPolicies.DeploymentProfiles += $profile
                            break
                        }
                        elseif ($assignment.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignment.GroupId) {
                            $profile | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                            $relevantPolicies.DeploymentProfiles += $profile
                            break
                        }
                    }
                }

                Write-Host "Fetching Enrollment Status Page Profiles..." -ForegroundColor Yellow
                $enrollmentConfigs = Get-IntuneEntities -EntityType "deviceEnrollmentConfigurations"
                $espProfiles = $enrollmentConfigs | Where-Object { $_.'@odata.type' -match 'EnrollmentCompletionPageConfiguration' }
                foreach ($esp in $espProfiles) {
                    $assignments = Get-IntuneAssignments -EntityType "deviceEnrollmentConfigurations" -EntityId $esp.id
                    foreach ($assignment in $assignments) {
                        if (($assignment.Reason -eq "All Devices") -or
                            ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId)) {
                            $esp | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                            $relevantPolicies.ESPProfiles += $esp
                            break
                        }
                        elseif ($assignment.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignment.GroupId) {
                            $esp | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                            $relevantPolicies.ESPProfiles += $esp
                            break
                        }
                    }
                }

                Write-Host "Fetching Windows 365 Cloud PC Provisioning Policies..." -ForegroundColor Yellow
                try {
                    $cloudPCProvisioningPolicies = Get-IntuneEntities -EntityType "virtualEndpoint/provisioningPolicies"
                    foreach ($policy in $cloudPCProvisioningPolicies) {
                        $assignments = Get-IntuneAssignments -EntityType "virtualEndpoint/provisioningPolicies" -EntityId $policy.id
                        foreach ($assignment in $assignments) {
                            if (($assignment.Reason -eq "All Devices") -or
                                ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId)) {
                                $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                                $relevantPolicies.CloudPCProvisioningPolicies += $policy
                                break
                            }
                            elseif ($assignment.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignment.GroupId) {
                                $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                $relevantPolicies.CloudPCProvisioningPolicies += $policy
                                break
                            }
                        }
                    }
                }
                catch {
                    Write-Warning "Unable to fetch Windows 365 Cloud PC Provisioning Policies: $($_.Exception.Message)"
                }

                Write-Host "Fetching Windows 365 Cloud PC User Settings..." -ForegroundColor Yellow
                try {
                    $cloudPCUserSettings = Get-IntuneEntities -EntityType "virtualEndpoint/userSettings"
                    foreach ($setting in $cloudPCUserSettings) {
                        $assignments = Get-IntuneAssignments -EntityType "virtualEndpoint/userSettings" -EntityId $setting.id
                        foreach ($assignment in $assignments) {
                            if (($assignment.Reason -eq "All Devices") -or
                                ($assignment.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignment.GroupId)) {
                                $setting | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignment.Reason -Force
                                $relevantPolicies.CloudPCUserSettings += $setting
                                break
                            }
                            elseif ($assignment.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignment.GroupId) {
                                $setting | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                $relevantPolicies.CloudPCUserSettings += $setting
                                break
                            }
                        }
                    }
                }
                catch {
                    Write-Warning "Unable to fetch Windows 365 Cloud PC User Settings: $($_.Exception.Message)"
                }

                Write-Host "Fetching Antivirus Policies" -ForegroundColor Yellow
                $antivirusPoliciesFoundDevice = [System.Collections.ArrayList]::new()
                $processedAntivirusIdsDevice = [System.Collections.Generic.HashSet[string]]::new()

                $configPoliciesForAntivirusDevice = Get-IntuneEntities -EntityType "configurationPolicies"
                $matchingConfigPoliciesAntivirusDevice = $configPoliciesForAntivirusDevice | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAntivirus' }

                if ($matchingConfigPoliciesAntivirusDevice) {
                    foreach ($policy in $matchingConfigPoliciesAntivirusDevice) {
                        if ($processedAntivirusIdsDevice.Add($policy.id)) {
                            $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                            foreach ($assignmentDetail in $assignments) {
                                if (($assignmentDetail.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetail.GroupId) -or
                                    ($assignmentDetail.Reason -eq "All Devices")) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetail.Reason -Force
                                    [void]$antivirusPoliciesFoundDevice.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetail.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetail.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$antivirusPoliciesFoundDevice.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }

                $allIntentsForAntivirusDevice = Get-IntuneEntities -EntityType "deviceManagement/intents"
                $matchingIntentsAntivirusDevice = $allIntentsForAntivirusDevice | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAntivirus' }

                if ($matchingIntentsAntivirusDevice) {
                    foreach ($policy in $matchingIntentsAntivirusDevice) {
                        if ($processedAntivirusIdsDevice.Add($policy.id)) {
                            $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                            $assignments = $assignmentsResponse.value
                            
                            foreach ($assignment in $assignments) {
                                $assignmentDetails = @{
                                    Reason  = switch ($assignment.target.'@odata.type') {
                                        '#microsoft.graph.allDevicesAssignmentTarget' { "All Devices" }
                                        '#microsoft.graph.groupAssignmentTarget' { "Group Assignment" }
                                        '#microsoft.graph.exclusionGroupAssignmentTarget' { "Group Exclusion" }
                                        default { "Unknown" }
                                    }
                                    GroupId = if ($assignment.target.'@odata.type' -match "groupAssignmentTarget") { $assignment.target.groupId } else { $null }
                                }

                                if (($assignmentDetails.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetails.GroupId) -or
                                    ($assignmentDetails.Reason -eq "All Devices")) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetails.Reason -Force
                                    [void]$antivirusPoliciesFoundDevice.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetails.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetails.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$antivirusPoliciesFoundDevice.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }
                $relevantPolicies.AntivirusProfiles = $antivirusPoliciesFoundDevice

                Write-Host "Fetching Disk Encryption Policies." -ForegroundColor Yellow
                $diskEncryptionPoliciesFoundDevice = [System.Collections.ArrayList]::new()
                $processedDiskEncryptionIdsDevice = [System.Collections.Generic.HashSet[string]]::new()

                $configPoliciesForDiskEncDevice = Get-IntuneEntities -EntityType "configurationPolicies"
                $matchingConfigPoliciesDiskEncDevice = $configPoliciesForDiskEncDevice | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityDiskEncryption' }
                
                if ($matchingConfigPoliciesDiskEncDevice) {
                    foreach ($policy in $matchingConfigPoliciesDiskEncDevice) {
                        if ($processedDiskEncryptionIdsDevice.Add($policy.id)) {
                            $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                            foreach ($assignmentDetail in $assignments) {
                                if (($assignmentDetail.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetail.GroupId) -or
                                    ($assignmentDetail.Reason -eq "All Devices")) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetail.Reason -Force
                                    [void]$diskEncryptionPoliciesFoundDevice.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetail.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetail.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$diskEncryptionPoliciesFoundDevice.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }

                $allIntentsForDiskEncDevice = Get-IntuneEntities -EntityType "deviceManagement/intents"
                $matchingIntentsDiskEncDevice = $allIntentsForDiskEncDevice | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityDiskEncryption' }

                if ($matchingIntentsDiskEncDevice) {
                    foreach ($policy in $matchingIntentsDiskEncDevice) {
                        if ($processedDiskEncryptionIdsDevice.Add($policy.id)) {
                            $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                            $assignments = $assignmentsResponse.value
                            foreach ($assignment in $assignments) {
                                $assignmentDetails = @{
                                    Reason  = switch ($assignment.target.'@odata.type') {
                                        '#microsoft.graph.allDevicesAssignmentTarget' { "All Devices" }
                                        '#microsoft.graph.groupAssignmentTarget' { "Group Assignment" }
                                        '#microsoft.graph.exclusionGroupAssignmentTarget' { "Group Exclusion" }
                                        default { "Unknown" }
                                    }
                                    GroupId = if ($assignment.target.'@odata.type' -match "groupAssignmentTarget") { $assignment.target.groupId } else { $null }
                                }
                                if (($assignmentDetails.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetails.GroupId) -or
                                    ($assignmentDetails.Reason -eq "All Devices")) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetails.Reason -Force
                                    [void]$diskEncryptionPoliciesFoundDevice.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetails.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetails.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$diskEncryptionPoliciesFoundDevice.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }
                $relevantPolicies.DiskEncryptionProfiles = $diskEncryptionPoliciesFoundDevice

                Write-Host "Fetching Firewall Policies" -ForegroundColor Yellow
                $firewallPoliciesFoundDevice = [System.Collections.ArrayList]::new()
                $processedFirewallIdsDevice = [System.Collections.Generic.HashSet[string]]::new()

                $configPoliciesForFirewallDevice = Get-IntuneEntities -EntityType "configurationPolicies"
                $matchingConfigPoliciesFirewallDevice = $configPoliciesForFirewallDevice | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityFirewall' }
                
                if ($matchingConfigPoliciesFirewallDevice) {
                    foreach ($policy in $matchingConfigPoliciesFirewallDevice) {
                        if ($processedFirewallIdsDevice.Add($policy.id)) {
                            $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                            foreach ($assignmentDetail in $assignments) {
                                if (($assignmentDetail.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetail.GroupId) -or
                                    ($assignmentDetail.Reason -eq "All Devices")) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetail.Reason -Force
                                    [void]$firewallPoliciesFoundDevice.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetail.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetail.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$firewallPoliciesFoundDevice.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }

                $allIntentsForFirewallDevice = Get-IntuneEntities -EntityType "deviceManagement/intents"
                $matchingIntentsFirewallDevice = $allIntentsForFirewallDevice | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityFirewall' }

                if ($matchingIntentsFirewallDevice) {
                    foreach ($policy in $matchingIntentsFirewallDevice) {
                        if ($processedFirewallIdsDevice.Add($policy.id)) {
                            $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                            $assignments = $assignmentsResponse.value
                            foreach ($assignment in $assignments) {
                                $assignmentDetails = @{
                                    Reason  = switch ($assignment.target.'@odata.type') {
                                        '#microsoft.graph.allDevicesAssignmentTarget' { "All Devices" }
                                        '#microsoft.graph.groupAssignmentTarget' { "Group Assignment" }
                                        '#microsoft.graph.exclusionGroupAssignmentTarget' { "Group Exclusion" }
                                        default { "Unknown" }
                                    }
                                    GroupId = if ($assignment.target.'@odata.type' -match "groupAssignmentTarget") { $assignment.target.groupId } else { $null }
                                }
                                if (($assignmentDetails.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetails.GroupId) -or
                                    ($assignmentDetails.Reason -eq "All Devices")) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetails.Reason -Force
                                    [void]$firewallPoliciesFoundDevice.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetails.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetails.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$firewallPoliciesFoundDevice.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }
                $relevantPolicies.FirewallProfiles = $firewallPoliciesFoundDevice

                Write-Host "Fetching EDR Policies" -ForegroundColor Yellow
                $edrPoliciesFoundDevice = [System.Collections.ArrayList]::new()
                $processedEDRIdsDevice = [System.Collections.Generic.HashSet[string]]::new()

                $configPoliciesForEDRDevice = Get-IntuneEntities -EntityType "configurationPolicies"
                $matchingConfigPoliciesEDRDevice = $configPoliciesForEDRDevice | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityEndpointDetectionAndResponse' }

                if ($matchingConfigPoliciesEDRDevice) {
                    foreach ($policy in $matchingConfigPoliciesEDRDevice) {
                        if ($processedEDRIdsDevice.Add($policy.id)) {
                            $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                            foreach ($assignmentDetail in $assignments) {
                                if (($assignmentDetail.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetail.GroupId) -or
                                    ($assignmentDetail.Reason -eq "All Devices")) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetail.Reason -Force
                                    [void]$edrPoliciesFoundDevice.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetail.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetail.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$edrPoliciesFoundDevice.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }

                $allIntentsForEDRDevice = Get-IntuneEntities -EntityType "deviceManagement/intents"
                $matchingIntentsEDRDevice = $allIntentsForEDRDevice | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityEndpointDetectionAndResponse' }
                
                if ($matchingIntentsEDRDevice) {
                    foreach ($policy in $matchingIntentsEDRDevice) {
                        if ($processedEDRIdsDevice.Add($policy.id)) {
                            $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                            $assignments = $assignmentsResponse.value
                            foreach ($assignment in $assignments) {
                                $assignmentDetails = @{
                                    Reason  = switch ($assignment.target.'@odata.type') {
                                        '#microsoft.graph.allDevicesAssignmentTarget' { "All Devices" }
                                        '#microsoft.graph.groupAssignmentTarget' { "Group Assignment" }
                                        '#microsoft.graph.exclusionGroupAssignmentTarget' { "Group Exclusion" }
                                        default { "Unknown" }
                                    }
                                    GroupId = if ($assignment.target.'@odata.type' -match "groupAssignmentTarget") { $assignment.target.groupId } else { $null }
                                }
                                if (($assignmentDetails.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetails.GroupId) -or
                                    ($assignmentDetails.Reason -eq "All Devices")) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetails.Reason -Force
                                    [void]$edrPoliciesFoundDevice.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetails.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetails.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$edrPoliciesFoundDevice.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }
                $relevantPolicies.EndpointDetectionProfiles = $edrPoliciesFoundDevice

                Write-Host "Fetching ASR Policies" -ForegroundColor Yellow
                $asrPoliciesFoundDevice = [System.Collections.ArrayList]::new()
                $processedASRIdsDevice = [System.Collections.Generic.HashSet[string]]::new()

                $configPoliciesForASRDevice = Get-IntuneEntities -EntityType "configurationPolicies"
                $matchingConfigPoliciesASRDevice = $configPoliciesForASRDevice | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAttackSurfaceReductionRules' }

                if ($matchingConfigPoliciesASRDevice) {
                    foreach ($policy in $matchingConfigPoliciesASRDevice) {
                        if ($processedASRIdsDevice.Add($policy.id)) {
                            $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                            foreach ($assignmentDetail in $assignments) {
                                if (($assignmentDetail.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetail.GroupId) -or
                                    ($assignmentDetail.Reason -eq "All Devices")) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetail.Reason -Force
                                    [void]$asrPoliciesFoundDevice.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetail.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetail.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$asrPoliciesFoundDevice.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }

                $allIntentsForASRDevice = Get-IntuneEntities -EntityType "deviceManagement/intents"
                $matchingIntentsASRDevice = $allIntentsForASRDevice | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAttackSurfaceReductionRules' }

                if ($matchingIntentsASRDevice) {
                    foreach ($policy in $matchingIntentsASRDevice) {
                        if ($processedASRIdsDevice.Add($policy.id)) {
                            $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                            $assignments = $assignmentsResponse.value
                            foreach ($assignment in $assignments) {
                                $assignmentDetails = @{
                                    Reason  = switch ($assignment.target.'@odata.type') {
                                        '#microsoft.graph.allDevicesAssignmentTarget' { "All Devices" }
                                        '#microsoft.graph.groupAssignmentTarget' { "Group Assignment" }
                                        '#microsoft.graph.exclusionGroupAssignmentTarget' { "Group Exclusion" }
                                        default { "Unknown" }
                                    }
                                    GroupId = if ($assignment.target.'@odata.type' -match "groupAssignmentTarget") { $assignment.target.groupId } else { $null }
                                }
                                if (($assignmentDetails.Reason -eq "Group Assignment" -and $groupMemberships.id -contains $assignmentDetails.GroupId) -or
                                    ($assignmentDetails.Reason -eq "All Devices")) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $assignmentDetails.Reason -Force
                                    [void]$asrPoliciesFoundDevice.Add($policy)
                                    break
                                }
                                elseif ($assignmentDetails.Reason -eq "Group Exclusion" -and $groupMemberships.id -contains $assignmentDetails.GroupId) {
                                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "Excluded" -Force
                                    [void]$asrPoliciesFoundDevice.Add($policy)
                                    break
                                }
                            }
                        }
                    }
                }
                $relevantPolicies.AttackSurfaceProfiles = $asrPoliciesFoundDevice

                Write-Host "Fetching Applications..." -ForegroundColor Yellow
                $appUri = "$GraphEndpoint/beta/deviceAppManagement/mobileApps?`$filter=isAssigned eq true"
                $appResponse = Invoke-MgGraphRequest -Uri $appUri -Method Get
                $allApps = $appResponse.value
                while ($appResponse.'@odata.nextLink') {
                    $appResponse = Invoke-MgGraphRequest -Uri $appResponse.'@odata.nextLink' -Method Get
                    $allApps += $appResponse.value
                }
                $totalApps = $allApps.Count

                foreach ($app in $allApps) {
                    if ($app.isFeatured -or $app.isBuiltIn) {
                        continue
                    }

                    $appId = $app.id
                    $assignmentsUri = "$GraphEndpoint/beta/deviceAppManagement/mobileApps('$appId')/assignments"
                    $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get

                    $isExcluded = $false
                    $isIncluded = $false
                    $inclusionReason = ""
                    $exclusionReason = ""

                    foreach ($assignment in $assignmentResponse.value) {
                        if ($assignment.target.'@odata.type' -eq '#microsoft.graph.exclusionGroupAssignmentTarget' -and
                            $groupMemberships.id -contains $assignment.target.groupId) {
                            $isExcluded = $true
                            $groupInfo = Get-GroupInfo -GroupId $assignment.target.groupId
                            $exclusionReason = "Excluded via group: $($groupInfo.DisplayName)"
                            break
                        }
                        elseif ($assignment.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget') {
                            $isIncluded = $true
                            $inclusionReason = "All Devices"
                        }
                        elseif ($assignment.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget' -and
                            $groupMemberships.id -contains $assignment.target.groupId) {
                            $isIncluded = $true
                            $groupInfo = Get-GroupInfo -GroupId $assignment.target.groupId
                            $inclusionReason = "Group Assignment - $($groupInfo.DisplayName)"
                        }
                    }

                    if ($isExcluded) {
                        $appWithReason = $app.PSObject.Copy()
                        $appWithReason | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $exclusionReason -Force
                        switch ($assignment.intent) {
                            "required" { $relevantPolicies.AppsRequired += $appWithReason; break }
                            "available" { $relevantPolicies.AppsAvailable += $appWithReason; break }
                            "uninstall" { $relevantPolicies.AppsUninstall += $appWithReason; break }
                        }
                    }
                    elseif ($isIncluded) {
                        $appWithReason = $app.PSObject.Copy()
                        $appWithReason | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue $inclusionReason -Force
                        switch ($assignment.intent) {
                            "required" { $relevantPolicies.AppsRequired += $appWithReason; break }
                            "available" { $relevantPolicies.AppsAvailable += $appWithReason; break }
                            "uninstall" { $relevantPolicies.AppsUninstall += $appWithReason; break }
                        }
                    }
                }
 
                Write-Host "`nAssignments for Device: $deviceName" -ForegroundColor Green

                function Format-PolicyTable {
                    param (
                        [string]$Title,
                        [object[]]$Policies,
                        [scriptblock]$GetName,
                        [scriptblock]$GetExtra = { param($p) "" }
                    )
                    $tableSeparator = "-" * 120 # Define at the start for use in empty case

                    $headerSeparator = "-" * ($Title.Length + 16)
                    Write-Host "`n$headerSeparator" -ForegroundColor Cyan
                    Write-Host "------- $Title -------" -ForegroundColor Cyan
                    Write-Host "$headerSeparator" -ForegroundColor Cyan
                    
                    if ($Policies.Count -eq 0) {
                        Write-Host "No $Title found for this device." -ForegroundColor Gray
                        Write-Host $tableSeparator -ForegroundColor Gray # Print bottom line for empty table
                        Write-Host ""
                        return
                    }

                    $headerFormat = "{0,-50} {1,-40} {2,-30}" -f "Policy Name", "ID", "Assignment"
                    
                    Write-Host $headerFormat -ForegroundColor Yellow
                    Write-Host $tableSeparator -ForegroundColor Gray # This is the line under the headers
                    
                    foreach ($policy in $Policies) {
                        $name = & $GetName $policy
                        $extra = & $GetExtra $policy
                        
                        if ($name.Length -gt 47) {
                            $name = $name.Substring(0, 44) + "..."
                        }
                        
                        $id = $policy.id
                        if ($id.Length -gt 37) {
                            $id = $id.Substring(0, 34) + "..."
                        }
                        
                        $assignment = if ($policy.AssignmentReason) { $policy.AssignmentReason } else { "No Assignment" }
                        if ($assignment.Length -gt 27) {
                            $assignment = $assignment.Substring(0, 24) + "..."
                        }
                        
                        $rowFormat = "{0,-50} {1,-40} {2,-30}" -f $name, $id, $assignment
                        if ($assignment -eq "Excluded" -or $assignment -like "*Exclusion*") {
                            Write-Host $rowFormat -ForegroundColor Red
                        }
                        else {
                            Write-Host $rowFormat -ForegroundColor White
                        }
                    }
                    
                    Write-Host $tableSeparator -ForegroundColor Gray # This is the closing line of the table
                }

                Format-PolicyTable -Title "Device Configurations" -Policies $relevantPolicies.DeviceConfigs -GetName {
                    param($config)
                    if ([string]::IsNullOrWhiteSpace($config.name)) { $config.displayName } else { $config.name }
                }

                Format-PolicyTable -Title "Settings Catalog Policies" -Policies $relevantPolicies.SettingsCatalog -GetName {
                    param($policy)
                    if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                }

                Format-PolicyTable -Title "Administrative Templates" -Policies $relevantPolicies.AdminTemplates -GetName {
                    param($template)
                    if ([string]::IsNullOrWhiteSpace($template.name)) { $template.displayName } else { $template.name }
                }

                Format-PolicyTable -Title "Compliance Policies" -Policies $relevantPolicies.CompliancePolicies -GetName {
                    param($policy)
                    if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                }

                Format-PolicyTable -Title "App Protection Policies" -Policies $relevantPolicies.AppProtectionPolicies -GetName {
                    param($policy)
                    $policy.displayName
                } -GetExtra {
                    param($policy)
                    @{
                        Label = 'Platform'
                        Value = switch ($policy.'@odata.type') {
                            "#microsoft.graph.androidManagedAppProtection" { "Android" }
                            "#microsoft.graph.iosManagedAppProtection" { "iOS" }
                            "#microsoft.graph.windowsManagedAppProtection" { "Windows" }
                            default { "Unknown" }
                        }
                    }
                }

                Format-PolicyTable -Title "App Configuration Policies" -Policies $relevantPolicies.AppConfigurationPolicies -GetName {
                    param($policy)
                    if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                }

                Format-PolicyTable -Title "Platform Scripts" -Policies $relevantPolicies.PlatformScripts -GetName {
                    param($script)
                    if ([string]::IsNullOrWhiteSpace($script.name)) { $script.displayName } else { $script.name }
                }

                Format-PolicyTable -Title "Proactive Remediation Scripts" -Policies $relevantPolicies.HealthScripts -GetName {
                    param($script)
                    if ([string]::IsNullOrWhiteSpace($script.name)) { $script.displayName } else { $script.name }
                }

                Format-PolicyTable -Title "Required Apps" -Policies $relevantPolicies.AppsRequired -GetName {
                    param($app)
                    $app.displayName
                }

                Format-PolicyTable -Title "Available Apps" -Policies $relevantPolicies.AppsAvailable -GetName {
                    param($app)
                    $app.displayName
                }

                Format-PolicyTable -Title "Uninstall Apps" -Policies $relevantPolicies.AppsUninstall -GetName {
                    param($app)
                    $app.displayName
                }

                Format-PolicyTable -Title "Endpoint Security - Antivirus Profiles" -Policies $relevantPolicies.AntivirusProfiles -GetName { param($profile) if (-not [string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.displayName } elseif (-not [string]::IsNullOrWhiteSpace($profile.name)) { $profile.name } else { "Unnamed Profile" } }
                
                Format-PolicyTable -Title "Endpoint Security - Disk Encryption Profiles" -Policies $relevantPolicies.DiskEncryptionProfiles -GetName { param($profile) if (-not [string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.displayName } elseif (-not [string]::IsNullOrWhiteSpace($profile.name)) { $profile.name } else { "Unnamed Profile" } }
                
                Format-PolicyTable -Title "Endpoint Security - Firewall Profiles" -Policies $relevantPolicies.FirewallProfiles -GetName { param($profile) if (-not [string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.displayName } elseif (-not [string]::IsNullOrWhiteSpace($profile.name)) { $profile.name } else { "Unnamed Profile" } }
                
                Format-PolicyTable -Title "Endpoint Security - EDR Profiles" -Policies $relevantPolicies.EndpointDetectionProfiles -GetName { param($profile) if (-not [string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.displayName } elseif (-not [string]::IsNullOrWhiteSpace($profile.name)) { $profile.name } else { "Unnamed Profile" } }
                
                Format-PolicyTable -Title "Endpoint Security - ASR Profiles" -Policies $relevantPolicies.AttackSurfaceProfiles -GetName { param($profile) if (-not [string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.displayName } elseif (-not [string]::IsNullOrWhiteSpace($profile.name)) { $profile.name } else { "Unnamed Profile" } }

                Add-ExportData -ExportData $exportData -Category "Device" -Items @([PSCustomObject]@{
                        displayName      = $deviceName
                        id               = $deviceInfo.Id
                        AssignmentReason = "N/A"
                    }

                    Add-ExportData -ExportData $exportData -Category "Device Configuration" -Items $relevantPolicies.DeviceConfigs -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Settings Catalog Policy" -Items $relevantPolicies.SettingsCatalog -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Administrative Template" -Items $relevantPolicies.AdminTemplates -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Compliance Policy" -Items $relevantPolicies.CompliancePolicies -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "App Protection Policy" -Items $relevantPolicies.AppProtectionPolicies -AssignmentReason { param($item) $item.AssignmentSummary }
                    Add-ExportData -ExportData $exportData -Category "App Configuration Policy" -Items $relevantPolicies.AppConfigurationPolicies -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Platform Scripts" -Items $relevantPolicies.PlatformScripts -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Proactive Remediation Scripts" -Items $relevantPolicies.HealthScripts -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Antivirus" -Items $relevantPolicies.AntivirusProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Disk Encryption" -Items $relevantPolicies.DiskEncryptionProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Firewall" -Items $relevantPolicies.FirewallProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - EDR" -Items $relevantPolicies.EndpointDetectionProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - ASR" -Items $relevantPolicies.AttackSurfaceProfiles -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Windows 365 Cloud PC Provisioning Policy" -Items $relevantPolicies.CloudPCProvisioningPolicies -AssignmentReason { param($item) $item.AssignmentReason }
                    Add-ExportData -ExportData $exportData -Category "Windows 365 Cloud PC User Setting" -Items $relevantPolicies.CloudPCUserSettings -AssignmentReason { param($item) $item.AssignmentReason }
                )
            }

            Export-ResultsIfRequested -ExportData $exportData -DefaultFileName "IntuneDeviceAssignments.csv" -ForceExport:$ExportToCSV -CustomExportPath $ExportPath
        }
        '4' {
            Write-Host "Fetching all policies and their assignments..." -ForegroundColor Green
            $exportData = [System.Collections.ArrayList]::new()

            $allPolicies = @{
                DeviceConfigs               = @()
                SettingsCatalog             = @()
                AdminTemplates              = @()
                CompliancePolicies          = @()
                AppProtectionPolicies       = @()
                AppConfigurationPolicies    = @()
                PlatformScripts             = @()
                HealthScripts               = @()
                AntivirusProfiles           = @()
                DiskEncryptionProfiles      = @()
                FirewallProfiles            = @()
                EndpointDetectionProfiles   = @()
                AttackSurfaceProfiles       = @()
                DeploymentProfiles          = @()
                ESPProfiles                 = @()
                CloudPCProvisioningPolicies = @()
                CloudPCUserSettings         = @()
            }

            function Process-PolicyAssignments {
                param (
                    [Parameter(Mandatory = $true)]
                    [string]$PolicyType,
                    
                    [Parameter(Mandatory = $false)] # Changed from $true
                    [object[]]$Policies,
                    
                    [Parameter(Mandatory = $true)]
                    [string]$DisplayName
                )

                if ($null -eq $Policies -or $Policies.Count -eq 0) {
                    Write-Host "`n------- $DisplayName -------" -ForegroundColor Cyan
                    Write-Host "No policies found for this category." -ForegroundColor Gray
                    Write-Host ""
                    return
                }
                
                Write-Host "`n------- $DisplayName -------" -ForegroundColor Cyan
                foreach ($policy in $Policies) {
                    $policyName = if (-not [string]::IsNullOrWhiteSpace($policy.displayName)) { $policy.displayName } elseif (-not [string]::IsNullOrWhiteSpace($policy.name)) { $policy.name } else { "Unnamed Profile" }
                    Write-Host "Policy Name: $policyName" -ForegroundColor White
                    Write-Host "Policy ID: $($policy.id)" -ForegroundColor Gray
                    if ($policy.AssignmentSummary) {
                        Write-Host "Assignments: $($policy.AssignmentSummary)" -ForegroundColor Gray
                    }
                    else {
                        Write-Host "No assignments found" -ForegroundColor Yellow
                    }
                    Write-Host ""
                }
            }

            Write-Host "Fetching Device Configurations..." -ForegroundColor Yellow
            $deviceConfigs = Get-IntuneEntities -EntityType "deviceConfigurations"
            foreach ($config in $deviceConfigs) {
                $assignments = Get-IntuneAssignments -EntityType "deviceConfigurations" -EntityId $config.id
                $assignmentSummary = $assignments | ForEach-Object {
                    if ($_.Reason -eq "Group Assignment") {
                        $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                        "$($_.Reason) - $($groupInfo.DisplayName)"
                    }
                    else {
                        $_.Reason
                    }
                }
                $config | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                $allPolicies.DeviceConfigs += $config
            }

            Write-Host "Fetching Settings Catalog Policies..." -ForegroundColor Yellow
            $settingsCatalog = Get-IntuneEntities -EntityType "configurationPolicies"
            foreach ($policy in $settingsCatalog) {
                $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                $assignmentSummary = $assignments | ForEach-Object {
                    if ($_.Reason -eq "Group Assignment") {
                        $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                        "$($_.Reason) - $($groupInfo.DisplayName)"
                    }
                    else {
                        $_.Reason
                    }
                }
                $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                $allPolicies.SettingsCatalog += $policy
            }

            Write-Host "Fetching Administrative Templates..." -ForegroundColor Yellow
            $adminTemplates = Get-IntuneEntities -EntityType "groupPolicyConfigurations"
            foreach ($template in $adminTemplates) {
                $assignments = Get-IntuneAssignments -EntityType "groupPolicyConfigurations" -EntityId $template.id
                $assignmentSummary = $assignments | ForEach-Object {
                    if ($_.Reason -eq "Group Assignment") {
                        $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                        "$($_.Reason) - $($groupInfo.DisplayName)"
                    }
                    else {
                        $_.Reason
                    }
                }
                $template | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                $allPolicies.AdminTemplates += $template
            }

            Write-Host "Fetching Compliance Policies..." -ForegroundColor Yellow
            $compliancePolicies = Get-IntuneEntities -EntityType "deviceCompliancePolicies"
            foreach ($policy in $compliancePolicies) {
                $assignments = Get-IntuneAssignments -EntityType "deviceCompliancePolicies" -EntityId $policy.id
                $assignmentSummary = $assignments | ForEach-Object {
                    if ($_.Reason -eq "Group Assignment") {
                        $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                        "$($_.Reason) - $($groupInfo.DisplayName)"
                    }
                    else {
                        $_.Reason
                    }
                }
                $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                $allPolicies.CompliancePolicies += $policy
            }

            Write-Host "Fetching App Protection Policies..." -ForegroundColor Yellow
            $appProtectionPolicies = Get-IntuneEntities -EntityType "deviceAppManagement/managedAppPolicies"
            foreach ($policy in $appProtectionPolicies) {
                $policyType = $policy.'@odata.type'
                $assignmentsUri = switch ($policyType) {
                    "#microsoft.graph.androidManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/androidManagedAppProtections('$($policy.id)')/assignments" }
                    "#microsoft.graph.iosManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/iosManagedAppProtections('$($policy.id)')/assignments" }
                    "#microsoft.graph.windowsManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/windowsManagedAppProtections('$($policy.id)')/assignments" }
                    default { $null }
                }

                if ($assignmentsUri) {
                    try {
                        $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get
                        $assignments = @()
                        foreach ($assignment in $assignmentResponse.value) {
                            $assignmentReason = $null
                            switch ($assignment.target.'@odata.type') {
                                '#microsoft.graph.allLicensedUsersAssignmentTarget' {
                                    $assignmentReason = "All Users"
                                }
                                '#microsoft.graph.groupAssignmentTarget' {
                                    if (!$GroupId -or $assignment.target.groupId -eq $GroupId) {
                                        $assignmentReason = "Group Assignment"
                                    }
                                }
                                '#microsoft.graph.exclusionGroupAssignmentTarget' {
                                    if (!$GroupId -or $assignment.target.groupId -eq $GroupId) {
                                        $assignmentReason = "Group Exclusion"
                                    }
                                }
                            }

                            if ($assignmentReason) {
                                $assignments += $assignmentReason
                            }
                        }

                        if ($assignments.Count -gt 0) {
                            $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignments -join "; ") -Force
                            $allPolicies.AppProtectionPolicies += $policy
                        }
                    }
                    catch {
                        Write-Host "Error fetching assignments for policy $($policy.displayName): $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }

            Write-Host "Fetching App Configuration Policies..." -ForegroundColor Yellow
            $appConfigPolicies = Get-IntuneEntities -EntityType "deviceAppManagement/mobileAppConfigurations"
            foreach ($policy in $appConfigPolicies) {
                $assignments = Get-IntuneAssignments -EntityType "mobileAppConfigurations" -EntityId $policy.id
                $assignmentSummary = $assignments | ForEach-Object {
                    if ($_.Reason -eq "Group Assignment") {
                        $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                        "$($_.Reason) - $($groupInfo.DisplayName)"
                    }
                    else {
                        $_.Reason
                    }
                }
                $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                $allPolicies.AppConfigurationPolicies += $policy
            }

            Write-Host "Fetching Platform Scripts..." -ForegroundColor Yellow
            $platformScripts = Get-IntuneEntities -EntityType "deviceManagementScripts"
            foreach ($script in $platformScripts) {
                $assignments = Get-IntuneAssignments -EntityType "deviceManagementScripts" -EntityId $script.id
                $assignmentSummary = $assignments | ForEach-Object {
                    if ($_.Reason -eq "Group Assignment") {
                        $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                        "$($_.Reason) - $($groupInfo.DisplayName)"
                    }
                    else {
                        $_.Reason
                    }
                }
                $script | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                $allPolicies.PlatformScripts += $script
            }

            Write-Host "Fetching Proactive Remediation Scripts..." -ForegroundColor Yellow
            $healthScripts = Get-IntuneEntities -EntityType "deviceHealthScripts"
            foreach ($script in $healthScripts) {
                $assignments = Get-IntuneAssignments -EntityType "deviceHealthScripts" -EntityId $script.id
                $assignmentSummary = $assignments | ForEach-Object {
                    if ($_.Reason -eq "Group Assignment") {
                        $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                        "$($_.Reason) - $($groupInfo.DisplayName)"
                    }
                    else {
                        $_.Reason
                    }
                }
                $script | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                $allPolicies.HealthScripts += $script
            }

            Write-Host "Fetching Autopilot Deployment Profiles..." -ForegroundColor Yellow
            $autoProfilesAll = Get-IntuneEntities -EntityType "windowsAutopilotDeploymentProfiles"
            foreach ($profile in $autoProfilesAll) {
                $assignments = Get-IntuneAssignments -EntityType "windowsAutopilotDeploymentProfiles" -EntityId $profile.id
                $assignmentSummary = $assignments | ForEach-Object {
                    if ($_.Reason -eq "Group Assignment") {
                        $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                        "$($_.Reason) - $($groupInfo.DisplayName)"
                    }
                    else { $_.Reason }
                }
                $profile | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                $allPolicies.DeploymentProfiles += $profile
            }

            Write-Host "Fetching Enrollment Status Page Profiles..." -ForegroundColor Yellow
            $enrollmentConfigsAll = Get-IntuneEntities -EntityType "deviceEnrollmentConfigurations"
            $espProfilesAll = $enrollmentConfigsAll | Where-Object { $_.'@odata.type' -match 'EnrollmentCompletionPageConfiguration' }
            foreach ($esp in $espProfilesAll) {
                $assignments = Get-IntuneAssignments -EntityType "deviceEnrollmentConfigurations" -EntityId $esp.id
                $assignmentSummary = $assignments | ForEach-Object {
                    if ($_.Reason -eq "Group Assignment") {
                        $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                        "$($_.Reason) - $($groupInfo.DisplayName)"
                    }
                    else { $_.Reason }
                }
                $esp | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                $allPolicies.ESPProfiles += $esp
            }

            Write-Host "Fetching Windows 365 Cloud PC Provisioning Policies..." -ForegroundColor Yellow
            try {
                $cloudPCProvisioningPoliciesAll = Get-IntuneEntities -EntityType "virtualEndpoint/provisioningPolicies"
                foreach ($policy in $cloudPCProvisioningPoliciesAll) {
                    $assignments = Get-IntuneAssignments -EntityType "virtualEndpoint/provisioningPolicies" -EntityId $policy.id
                    $assignmentSummary = $assignments | ForEach-Object {
                        if ($_.Reason -eq "Group Assignment") {
                            $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                            "$($_.Reason) - $($groupInfo.DisplayName)"
                        }
                        else { $_.Reason }
                    }
                    $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                    $allPolicies.CloudPCProvisioningPolicies += $policy
                }
            }
            catch {
                Write-Warning "Unable to fetch Windows 365 Cloud PC Provisioning Policies: $($_.Exception.Message)"
            }

            Write-Host "Fetching Windows 365 Cloud PC User Settings..." -ForegroundColor Yellow
            try {
                $cloudPCUserSettingsAll = Get-IntuneEntities -EntityType "virtualEndpoint/userSettings"
                foreach ($setting in $cloudPCUserSettingsAll) {
                    $assignments = Get-IntuneAssignments -EntityType "virtualEndpoint/userSettings" -EntityId $setting.id
                    $assignmentSummary = $assignments | ForEach-Object {
                        if ($_.Reason -eq "Group Assignment") {
                            $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                            "$($_.Reason) - $($groupInfo.DisplayName)"
                        }
                        else { $_.Reason }
                    }
                    $setting | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                    $allPolicies.CloudPCUserSettings += $setting
                }
            }
            catch {
                Write-Warning "Unable to fetch Windows 365 Cloud PC User Settings: $($_.Exception.Message)"
            }

            Write-Host "Fetching Antivirus Policies..." -ForegroundColor Yellow
            $antivirusPoliciesFoundAll = [System.Collections.ArrayList]::new()
            $processedAntivirusIdsAll = [System.Collections.Generic.HashSet[string]]::new()

            $configPoliciesForAntivirusAll = Get-IntuneEntities -EntityType "configurationPolicies"
            $matchingConfigPoliciesAntivirusAll = $configPoliciesForAntivirusAll | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAntivirus' }

            if ($matchingConfigPoliciesAntivirusAll) {
                foreach ($policy in $matchingConfigPoliciesAntivirusAll) {
                    if ($processedAntivirusIdsAll.Add($policy.id)) {
                        $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                        $assignmentSummary = $assignments | ForEach-Object {
                            if ($_.Reason -eq "Group Assignment" -or $_.Reason -eq "Group Exclusion") {
                                $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                                "$($_.Reason) - $($groupInfo.DisplayName)"
                            }
                            else { $_.Reason }
                        }
                        $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                        [void]$antivirusPoliciesFoundAll.Add($policy)
                    }
                }
            }

            $allIntentsForAntivirusAll = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $matchingIntentsAntivirusAll = $allIntentsForAntivirusAll | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAntivirus' }
            
            if ($matchingIntentsAntivirusAll) {
                foreach ($policy in $matchingIntentsAntivirusAll) {
                    if ($processedAntivirusIdsAll.Add($policy.id)) {
                        $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        $assignmentSummary = $assignmentsResponse.value | ForEach-Object {
                            $reasonText = switch ($_.target.'@odata.type') {
                                '#microsoft.graph.allLicensedUsersAssignmentTarget' { "All Users" }
                                '#microsoft.graph.allDevicesAssignmentTarget' { "All Devices" }
                                '#microsoft.graph.groupAssignmentTarget' { "Group: " + (Get-GroupInfo -GroupId $_.target.groupId).DisplayName }
                                '#microsoft.graph.exclusionGroupAssignmentTarget' { "Exclude Group: " + (Get-GroupInfo -GroupId $_.target.groupId).DisplayName }
                                default { "Unknown" }
                            }
                            $reasonText
                        }
                        $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                        [void]$antivirusPoliciesFoundAll.Add($policy)
                    }
                }
            }
            $allPolicies.AntivirusProfiles = $antivirusPoliciesFoundAll

            Write-Host "Fetching Disk Encryption Policies..." -ForegroundColor Yellow
            $diskEncryptionPoliciesFoundAll = [System.Collections.ArrayList]::new()
            $processedDiskEncryptionIdsAll = [System.Collections.Generic.HashSet[string]]::new()

            $configPoliciesForDiskEncAll = Get-IntuneEntities -EntityType "configurationPolicies"
            $matchingConfigPoliciesDiskEncAll = $configPoliciesForDiskEncAll | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityDiskEncryption' }

            if ($matchingConfigPoliciesDiskEncAll) {
                foreach ($policy in $matchingConfigPoliciesDiskEncAll) {
                    if ($processedDiskEncryptionIdsAll.Add($policy.id)) {
                        $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                        $assignmentSummary = $assignments | ForEach-Object {
                            if ($_.Reason -eq "Group Assignment" -or $_.Reason -eq "Group Exclusion") {
                                $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                                "$($_.Reason) - $($groupInfo.DisplayName)"
                            }
                            else { $_.Reason }
                        }
                        $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                        [void]$diskEncryptionPoliciesFoundAll.Add($policy)
                    }
                }
            }

            $allIntentsForDiskEncAll = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $matchingIntentsDiskEncAll = $allIntentsForDiskEncAll | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityDiskEncryption' }

            if ($matchingIntentsDiskEncAll) {
                foreach ($policy in $matchingIntentsDiskEncAll) {
                    if ($processedDiskEncryptionIdsAll.Add($policy.id)) {
                        $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        $assignmentSummary = $assignmentsResponse.value | ForEach-Object {
                            $reasonText = switch ($_.target.'@odata.type') {
                                '#microsoft.graph.allLicensedUsersAssignmentTarget' { "All Users" }
                                '#microsoft.graph.allDevicesAssignmentTarget' { "All Devices" }
                                '#microsoft.graph.groupAssignmentTarget' { "Group: " + (Get-GroupInfo -GroupId $_.target.groupId).DisplayName }
                                '#microsoft.graph.exclusionGroupAssignmentTarget' { "Exclude Group: " + (Get-GroupInfo -GroupId $_.target.groupId).DisplayName }
                                default { "Unknown" }
                            }
                            $reasonText
                        }
                        $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                        [void]$diskEncryptionPoliciesFoundAll.Add($policy)
                    }
                }
            }
            $allPolicies.DiskEncryptionProfiles = $diskEncryptionPoliciesFoundAll

            Write-Host "Fetching Firewall Policies..." -ForegroundColor Yellow
            $firewallPoliciesFoundAll = [System.Collections.ArrayList]::new()
            $processedFirewallIdsAll = [System.Collections.Generic.HashSet[string]]::new()

            $configPoliciesForFirewallAll = Get-IntuneEntities -EntityType "configurationPolicies"
            $matchingConfigPoliciesFirewallAll = $configPoliciesForFirewallAll | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityFirewall' }
            
            if ($matchingConfigPoliciesFirewallAll) {
                foreach ($policy in $matchingConfigPoliciesFirewallAll) {
                    if ($processedFirewallIdsAll.Add($policy.id)) {
                        $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                        $assignmentSummary = $assignments | ForEach-Object {
                            if ($_.Reason -eq "Group Assignment" -or $_.Reason -eq "Group Exclusion") {
                                $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                                "$($_.Reason) - $($groupInfo.DisplayName)"
                            }
                            else { $_.Reason }
                        }
                        $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                        [void]$firewallPoliciesFoundAll.Add($policy)
                    }
                }
            }

            $allIntentsForFirewallAll = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $matchingIntentsFirewallAll = $allIntentsForFirewallAll | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityFirewall' }

            if ($matchingIntentsFirewallAll) {
                foreach ($policy in $matchingIntentsFirewallAll) {
                    if ($processedFirewallIdsAll.Add($policy.id)) {
                        $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        $assignmentSummary = $assignmentsResponse.value | ForEach-Object {
                            $reasonText = switch ($_.target.'@odata.type') {
                                '#microsoft.graph.allLicensedUsersAssignmentTarget' { "All Users" }
                                '#microsoft.graph.allDevicesAssignmentTarget' { "All Devices" }
                                '#microsoft.graph.groupAssignmentTarget' { "Group: " + (Get-GroupInfo -GroupId $_.target.groupId).DisplayName }
                                '#microsoft.graph.exclusionGroupAssignmentTarget' { "Exclude Group: " + (Get-GroupInfo -GroupId $_.target.groupId).DisplayName }
                                default { "Unknown" }
                            }
                            $reasonText
                        }
                        $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                        [void]$firewallPoliciesFoundAll.Add($policy)
                    }
                }
            }
            $allPolicies.FirewallProfiles = $firewallPoliciesFoundAll

            Write-Host "Fetching EDR Policies..." -ForegroundColor Yellow
            $edrPoliciesFoundAll = [System.Collections.ArrayList]::new()
            $processedEDRIdsAll = [System.Collections.Generic.HashSet[string]]::new()

            $configPoliciesForEDRAll = Get-IntuneEntities -EntityType "configurationPolicies"
            $matchingConfigPoliciesEDRAll = $configPoliciesForEDRAll | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityEndpointDetectionAndResponse' }

            if ($matchingConfigPoliciesEDRAll) {
                foreach ($policy in $matchingConfigPoliciesEDRAll) {
                    if ($processedEDRIdsAll.Add($policy.id)) {
                        $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                        $assignmentSummary = $assignments | ForEach-Object {
                            if ($_.Reason -eq "Group Assignment" -or $_.Reason -eq "Group Exclusion") {
                                $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                                "$($_.Reason) - $($groupInfo.DisplayName)"
                            }
                            else { $_.Reason }
                        }
                        $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                        [void]$edrPoliciesFoundAll.Add($policy)
                    }
                }
            }

            $allIntentsForEDRAll = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $matchingIntentsEDRAll = $allIntentsForEDRAll | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityEndpointDetectionAndResponse' }

            if ($matchingIntentsEDRAll) {
                foreach ($policy in $matchingIntentsEDRAll) {
                    if ($processedEDRIdsAll.Add($policy.id)) {
                        $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        $assignmentSummary = $assignmentsResponse.value | ForEach-Object {
                            $reasonText = switch ($_.target.'@odata.type') {
                                '#microsoft.graph.allLicensedUsersAssignmentTarget' { "All Users" }
                                '#microsoft.graph.allDevicesAssignmentTarget' { "All Devices" }
                                '#microsoft.graph.groupAssignmentTarget' { "Group: " + (Get-GroupInfo -GroupId $_.target.groupId).DisplayName }
                                '#microsoft.graph.exclusionGroupAssignmentTarget' { "Exclude Group: " + (Get-GroupInfo -GroupId $_.target.groupId).DisplayName }
                                default { "Unknown" }
                            }
                            $reasonText
                        }
                        $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                        [void]$edrPoliciesFoundAll.Add($policy)
                    }
                }
            }
            $allPolicies.EndpointDetectionProfiles = $edrPoliciesFoundAll

            Write-Host "Fetching ASR Policies..." -ForegroundColor Yellow
            $asrPoliciesFoundAll = [System.Collections.ArrayList]::new()
            $processedASRIdsAll = [System.Collections.Generic.HashSet[string]]::new()

            $configPoliciesForASRAll = Get-IntuneEntities -EntityType "configurationPolicies"
            $matchingConfigPoliciesASRAll = $configPoliciesForASRAll | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAttackSurfaceReductionRules' }

            if ($matchingConfigPoliciesASRAll) {
                foreach ($policy in $matchingConfigPoliciesASRAll) {
                    if ($processedASRIdsAll.Add($policy.id)) {
                        $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                        $assignmentSummary = $assignments | ForEach-Object {
                            if ($_.Reason -eq "Group Assignment" -or $_.Reason -eq "Group Exclusion") {
                                $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                                "$($_.Reason) - $($groupInfo.DisplayName)"
                            }
                            else { $_.Reason }
                        }
                        $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                        [void]$asrPoliciesFoundAll.Add($policy)
                    }
                }
            }

            $allIntentsForASRAll = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $matchingIntentsASRAll = $allIntentsForASRAll | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAttackSurfaceReductionRules' }

            if ($matchingIntentsASRAll) {
                foreach ($policy in $matchingIntentsASRAll) {
                    if ($processedASRIdsAll.Add($policy.id)) {
                        $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        $assignmentSummary = $assignmentsResponse.value | ForEach-Object {
                            $reasonText = switch ($_.target.'@odata.type') {
                                '#microsoft.graph.allLicensedUsersAssignmentTarget' { "All Users" }
                                '#microsoft.graph.allDevicesAssignmentTarget' { "All Devices" }
                                '#microsoft.graph.groupAssignmentTarget' { "Group: " + (Get-GroupInfo -GroupId $_.target.groupId).DisplayName }
                                '#microsoft.graph.exclusionGroupAssignmentTarget' { "Exclude Group: " + (Get-GroupInfo -GroupId $_.target.groupId).DisplayName }
                                default { "Unknown" }
                            }
                            $reasonText
                        }
                        $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                        [void]$asrPoliciesFoundAll.Add($policy)
                    }
                }
            }
            $allPolicies.AttackSurfaceProfiles = $asrPoliciesFoundAll

            Process-PolicyAssignments -PolicyType "deviceConfigurations" -Policies $allPolicies.DeviceConfigs -DisplayName "Device Configurations"
            Process-PolicyAssignments -PolicyType "configurationPolicies" -Policies $allPolicies.SettingsCatalog -DisplayName "Settings Catalog Policies"
            Process-PolicyAssignments -PolicyType "groupPolicyConfigurations" -Policies $allPolicies.AdminTemplates -DisplayName "Administrative Templates"
            Process-PolicyAssignments -PolicyType "deviceCompliancePolicies" -Policies $allPolicies.CompliancePolicies -DisplayName "Compliance Policies"
            Process-PolicyAssignments -PolicyType "managedAppPolicies" -Policies $allPolicies.AppProtectionPolicies -DisplayName "App Protection Policies"
            Process-PolicyAssignments -PolicyType "mobileAppConfigurations" -Policies $allPolicies.AppConfigurationPolicies -DisplayName "App Configuration Policies"
            Process-PolicyAssignments -PolicyType "deviceManagementScripts" -Policies $allPolicies.PlatformScripts -DisplayName "Platform Scripts"
            Process-PolicyAssignments -PolicyType "deviceHealthScripts" -Policies $allPolicies.HealthScripts -DisplayName "Proactive Remediation Scripts"
            Process-PolicyAssignments -PolicyType "windowsAutopilotDeploymentProfiles" -Policies $allPolicies.DeploymentProfiles -DisplayName "Autopilot Deployment Profiles"
            Process-PolicyAssignments -PolicyType "deviceEnrollmentConfigurations" -Policies $allPolicies.ESPProfiles -DisplayName "Enrollment Status Page Profiles"
            Process-PolicyAssignments -PolicyType "virtualEndpoint/provisioningPolicies" -Policies $allPolicies.CloudPCProvisioningPolicies -DisplayName "Windows 365 Cloud PC Provisioning Policies"
            Process-PolicyAssignments -PolicyType "virtualEndpoint/userSettings" -Policies $allPolicies.CloudPCUserSettings -DisplayName "Windows 365 Cloud PC User Settings"
            Process-PolicyAssignments -PolicyType "deviceManagementIntents" -Policies $allPolicies.AntivirusProfiles -DisplayName "Endpoint Security - Antivirus Profiles"
            Process-PolicyAssignments -PolicyType "deviceManagementIntents" -Policies $allPolicies.DiskEncryptionProfiles -DisplayName "Endpoint Security - Disk Encryption Profiles"
            Process-PolicyAssignments -PolicyType "deviceManagementIntents" -Policies $allPolicies.FirewallProfiles -DisplayName "Endpoint Security - Firewall Profiles"
            Process-PolicyAssignments -PolicyType "deviceManagementIntents" -Policies $allPolicies.EndpointDetectionProfiles -DisplayName "Endpoint Security - EDR Profiles"
            Process-PolicyAssignments -PolicyType "deviceManagementIntents" -Policies $allPolicies.AttackSurfaceProfiles -DisplayName "Endpoint Security - ASR Profiles"

            Add-ExportData -ExportData $exportData -Category "Device Configuration" -Items $allPolicies.DeviceConfigs -AssignmentReason { param($item) $item.AssignmentSummary }
            Add-ExportData -ExportData $exportData -Category "Settings Catalog Policy" -Items $allPolicies.SettingsCatalog -AssignmentReason { param($item) $item.AssignmentSummary }
            Add-ExportData -ExportData $exportData -Category "Administrative Template" -Items $allPolicies.AdminTemplates -AssignmentReason { param($item) $item.AssignmentSummary }
            Add-ExportData -ExportData $exportData -Category "Compliance Policy" -Items $allPolicies.CompliancePolicies -AssignmentReason { param($item) $item.AssignmentSummary }
            Add-ExportData -ExportData $exportData -Category "App Protection Policy" -Items $allPolicies.AppProtectionPolicies -AssignmentReason { param($item) $item.AssignmentSummary }
            Add-ExportData -ExportData $exportData -Category "App Configuration Policy" -Items $allPolicies.AppConfigurationPolicies -AssignmentReason { param($item) $item.AssignmentSummary }
            Add-ExportData -ExportData $exportData -Category "Platform Scripts" -Items $allPolicies.PlatformScripts -AssignmentReason { param($item) $item.AssignmentSummary }
            Add-ExportData -ExportData $exportData -Category "Proactive Remediation Scripts" -Items $allPolicies.HealthScripts -AssignmentReason { param($item) $item.AssignmentSummary }
            Add-ExportData -ExportData $exportData -Category "Autopilot Deployment Profile" -Items $allPolicies.DeploymentProfiles -AssignmentReason { param($item) $item.AssignmentSummary }
            Add-ExportData -ExportData $exportData -Category "Enrollment Status Page" -Items $allPolicies.ESPProfiles -AssignmentReason { param($item) $item.AssignmentSummary }
            Add-ExportData -ExportData $exportData -Category "Windows 365 Cloud PC Provisioning Policy" -Items $allPolicies.CloudPCProvisioningPolicies -AssignmentReason { param($item) $item.AssignmentSummary }
            Add-ExportData -ExportData $exportData -Category "Windows 365 Cloud PC User Setting" -Items $allPolicies.CloudPCUserSettings -AssignmentReason { param($item) $item.AssignmentSummary }
            Add-ExportData -ExportData $exportData -Category "Endpoint Security - Antivirus" -Items $allPolicies.AntivirusProfiles -AssignmentReason { param($item) $item.AssignmentSummary }
            Add-ExportData -ExportData $exportData -Category "Endpoint Security - Disk Encryption" -Items $allPolicies.DiskEncryptionProfiles -AssignmentReason { param($item) $item.AssignmentSummary }
            Add-ExportData -ExportData $exportData -Category "Endpoint Security - Firewall" -Items $allPolicies.FirewallProfiles -AssignmentReason { param($item) $item.AssignmentSummary }
            Add-ExportData -ExportData $exportData -Category "Endpoint Security - EDR" -Items $allPolicies.EndpointDetectionProfiles -AssignmentReason { param($item) $item.AssignmentSummary }
            Add-ExportData -ExportData $exportData -Category "Endpoint Security - ASR" -Items $allPolicies.AttackSurfaceProfiles -AssignmentReason { param($item) $item.AssignmentSummary }

            Export-ResultsIfRequested -ExportData $exportData -DefaultFileName "IntuneAllPolicies.csv" -ForceExport:$ExportToCSV -CustomExportPath $ExportPath
        }
        '5' {
            Write-Host "Fetching all 'All Users' assignments..." -ForegroundColor Green
            $exportData = [System.Collections.ArrayList]::new()

            $allUsersAssignments = @{
                DeviceConfigs            = @()
                SettingsCatalog          = @()
                AdminTemplates           = @()
                CompliancePolicies       = @()
                AppProtectionPolicies    = @()
                AppConfigurationPolicies = @()
                PlatformScripts          = @()
                HealthScripts            = @()
                RequiredApps             = @()
                AvailableApps            = @()
                UninstallApps            = @()
                DeploymentProfiles       = @()
                ESPProfiles              = @()
            }

            Write-Host "Fetching Device Configurations..." -ForegroundColor Yellow
            $deviceConfigs = Get-IntuneEntities -EntityType "deviceConfigurations"
            foreach ($config in $deviceConfigs) {
                $assignments = Get-IntuneAssignments -EntityType "deviceConfigurations" -EntityId $config.id
                if ($assignments | Where-Object { $_.Reason -eq "All Users" }) {
                    $config | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                    $allUsersAssignments.DeviceConfigs += $config
                }
            }

            Write-Host "Fetching Settings Catalog Policies..." -ForegroundColor Yellow
            $settingsCatalog = Get-IntuneEntities -EntityType "configurationPolicies"
            foreach ($policy in $settingsCatalog) {
                $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                if ($assignments | Where-Object { $_.Reason -eq "All Users" }) {
                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                    $allUsersAssignments.SettingsCatalog += $policy
                }
            }

            Write-Host "Fetching Administrative Templates..." -ForegroundColor Yellow
            $adminTemplates = Get-IntuneEntities -EntityType "groupPolicyConfigurations"
            foreach ($template in $adminTemplates) {
                $assignments = Get-IntuneAssignments -EntityType "groupPolicyConfigurations" -EntityId $template.id
                if ($assignments | Where-Object { $_.Reason -eq "All Users" }) {
                    $template | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                    $allUsersAssignments.AdminTemplates += $template
                }
            }

            Write-Host "Fetching Compliance Policies..." -ForegroundColor Yellow
            $compliancePolicies = Get-IntuneEntities -EntityType "deviceCompliancePolicies"
            foreach ($policy in $compliancePolicies) {
                $assignments = Get-IntuneAssignments -EntityType "deviceCompliancePolicies" -EntityId $policy.id
                if ($assignments | Where-Object { $_.Reason -eq "All Users" }) {
                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                    $allUsersAssignments.CompliancePolicies += $policy
                }
            }

            Write-Host "Fetching App Protection Policies..." -ForegroundColor Yellow
            $appProtectionPolicies = Get-IntuneEntities -EntityType "deviceAppManagement/managedAppPolicies"
            foreach ($policy in $appProtectionPolicies) {
                $policyType = $policy.'@odata.type'
                $assignmentsUri = switch ($policyType) {
                    "#microsoft.graph.androidManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/androidManagedAppProtections('$($policy.id)')/assignments" }
                    "#microsoft.graph.iosManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/iosManagedAppProtections('$($policy.id)')/assignments" }
                    "#microsoft.graph.windowsManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/windowsManagedAppProtections('$($policy.id)')/assignments" }
                    default { $null }
                }

                if ($assignmentsUri) {
                    try {
                        $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get
                        $hasAllUsers = $false
                        foreach ($assignment in $assignmentResponse.value) {
                            if ($assignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
                                $hasAllUsers = $true
                                break
                            }
                        }
                        if ($hasAllUsers) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                            $allUsersAssignments.AppProtectionPolicies += $policy
                        }
                    }
                    catch {
                        Write-Host "Error fetching assignments for policy $($policy.displayName): $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }

            Write-Host "Fetching App Configuration Policies..." -ForegroundColor Yellow
            $appConfigPolicies = Get-IntuneEntities -EntityType "deviceAppManagement/mobileAppConfigurations"
            foreach ($policy in $appConfigPolicies) {
                $assignments = Get-IntuneAssignments -EntityType "mobileAppConfigurations" -EntityId $policy.id
                if ($assignments | Where-Object { $_.Reason -eq "All Users" }) {
                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                    $allUsersAssignments.AppConfigurationPolicies += $policy
                }
            }

            Write-Host "Fetching Applications..." -ForegroundColor Yellow
            $appUri = "$GraphEndpoint/beta/deviceAppManagement/mobileApps?`$filter=isAssigned eq true"
            $appResponse = Invoke-MgGraphRequest -Uri $appUri -Method Get
            $allApps = $appResponse.value
            while ($appResponse.'@odata.nextLink') {
                $appResponse = Invoke-MgGraphRequest -Uri $appResponse.'@odata.nextLink' -Method Get
                $allApps += $appResponse.value
            }
            $totalApps = $allApps.Count

            foreach ($app in $allApps) {
                if ($app.isFeatured -or $app.isBuiltIn) {
                    continue
                }

                $appId = $app.id
                $assignmentsUri = "$GraphEndpoint/beta/deviceAppManagement/mobileApps('$appId')/assignments"
                $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get

                foreach ($assignment in $assignmentResponse.value) {
                    if ($assignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
                        $appWithReason = $app.PSObject.Copy()
                        $appWithReason | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                        switch ($assignment.intent) {
                            "required" { $allUsersAssignments.RequiredApps += $appWithReason; break }
                            "available" { $allUsersAssignments.AvailableApps += $appWithReason; break }
                            "uninstall" { $allUsersAssignments.UninstallApps += $appWithReason; break }
                        }
                        break
                    }
                }
            }   

            Write-Host "Fetching Platform Scripts..." -ForegroundColor Yellow
            $platformScripts = Get-IntuneEntities -EntityType "deviceManagementScripts"
            foreach ($script in $platformScripts) {
                $assignments = Get-IntuneAssignments -EntityType "deviceManagementScripts" -EntityId $script.id
                if ($assignments | Where-Object { $_.Reason -eq "All Users" }) {
                    $script | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                    $allUsersAssignments.PlatformScripts += $script
                }
            }

            Write-Host "Fetching Proactive Remediation Scripts..." -ForegroundColor Yellow
            $healthScripts = Get-IntuneEntities -EntityType "deviceHealthScripts"
            foreach ($script in $healthScripts) {
                $assignments = Get-IntuneAssignments -EntityType "deviceHealthScripts" -EntityId $script.id
                if ($assignments | Where-Object { $_.Reason -eq "All Users" }) {
                    $script | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                    $allUsersAssignments.HealthScripts += $script
                }
            }

            Write-Host "Fetching Antivirus Policies assigned to All Users..." -ForegroundColor Yellow
            $antivirusPoliciesFound_AllUsers = [System.Collections.ArrayList]::new()
            $processedAntivirusIds_AllUsers = [System.Collections.Generic.HashSet[string]]::new()

            $configPoliciesForAntivirus_AllUsers = Get-IntuneEntities -EntityType "configurationPolicies"
            $matchingConfigPoliciesAntivirus_AllUsers = $configPoliciesForAntivirus_AllUsers | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAntivirus' }

            if ($matchingConfigPoliciesAntivirus_AllUsers) {
                foreach ($policy in $matchingConfigPoliciesAntivirus_AllUsers) {
                    if ($processedAntivirusIds_AllUsers.Add($policy.id)) {
                        $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                        if ($assignments | Where-Object { $_.Reason -eq "All Users" }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                            [void]$antivirusPoliciesFound_AllUsers.Add($policy)
                        }
                    }
                }
            }

            $allIntentsForAntivirus_AllUsers = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $matchingIntentsAntivirus_AllUsers = $allIntentsForAntivirus_AllUsers | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAntivirus' }

            if ($matchingIntentsAntivirus_AllUsers) {
                foreach ($policy in $matchingIntentsAntivirus_AllUsers) {
                    if ($processedAntivirusIds_AllUsers.Add($policy.id)) {
                        $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        if ($assignmentsResponse.value | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget' }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                            [void]$antivirusPoliciesFound_AllUsers.Add($policy)
                        }
                    }
                }
            }
            $allUsersAssignments.AntivirusProfiles = $antivirusPoliciesFound_AllUsers

            Write-Host "Fetching Disk Encryption Policies assigned to All Users..." -ForegroundColor Yellow
            $diskEncryptionPoliciesFound_AllUsers = [System.Collections.ArrayList]::new()
            $processedDiskEncryptionIds_AllUsers = [System.Collections.Generic.HashSet[string]]::new()


            $configPoliciesForDiskEnc_AllUsers = Get-IntuneEntities -EntityType "configurationPolicies"
            $matchingConfigPoliciesDiskEnc_AllUsers = $configPoliciesForDiskEnc_AllUsers | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityDiskEncryption' }
            
            if ($matchingConfigPoliciesDiskEnc_AllUsers) {
                foreach ($policy in $matchingConfigPoliciesDiskEnc_AllUsers) {
                    if ($processedDiskEncryptionIds_AllUsers.Add($policy.id)) {
                        $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                        if ($assignments | Where-Object { $_.Reason -eq "All Users" }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                            [void]$diskEncryptionPoliciesFound_AllUsers.Add($policy)
                        }
                    }
                }
            }

            $allIntentsForDiskEnc_AllUsers = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $matchingIntentsDiskEnc_AllUsers = $allIntentsForDiskEnc_AllUsers | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityDiskEncryption' }

            if ($matchingIntentsDiskEnc_AllUsers) {
                foreach ($policy in $matchingIntentsDiskEnc_AllUsers) {
                    if ($processedDiskEncryptionIds_AllUsers.Add($policy.id)) {
                        $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        if ($assignmentsResponse.value | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget' }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                            [void]$diskEncryptionPoliciesFound_AllUsers.Add($policy)
                        }
                    }
                }
            }
            $allUsersAssignments.DiskEncryptionProfiles = $diskEncryptionPoliciesFound_AllUsers

            Write-Host "Fetching Firewall Policies assigned to All Users..." -ForegroundColor Yellow
            $firewallPoliciesFound_AllUsers = [System.Collections.ArrayList]::new()
            $processedFirewallIds_AllUsers = [System.Collections.Generic.HashSet[string]]::new() # Reset for this type

            $configPoliciesForFirewall_AllUsers = Get-IntuneEntities -EntityType "configurationPolicies"
            $matchingConfigPoliciesFirewall_AllUsers = $configPoliciesForFirewall_AllUsers | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityFirewall' }

            if ($matchingConfigPoliciesFirewall_AllUsers) {
                foreach ($policy in $matchingConfigPoliciesFirewall_AllUsers) {
                    if ($processedFirewallIds_AllUsers.Add($policy.id)) {
                        $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                        if ($assignments | Where-Object { $_.Reason -eq "All Users" }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                            [void]$firewallPoliciesFound_AllUsers.Add($policy)
                        }
                    }
                }
            }

            $allIntentsForFirewall_AllUsers = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $matchingIntentsFirewall_AllUsers = $allIntentsForFirewall_AllUsers | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityFirewall' }

            if ($matchingIntentsFirewall_AllUsers) {
                foreach ($policy in $matchingIntentsFirewall_AllUsers) {
                    if ($processedFirewallIds_AllUsers.Add($policy.id)) {
                        $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        if ($assignmentsResponse.value | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget' }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                            [void]$firewallPoliciesFound_AllUsers.Add($policy)
                        }
                    }
                }
            }
            $allUsersAssignments.FirewallProfiles = $firewallPoliciesFound_AllUsers

            Write-Host "Fetching EDR Policies assigned to All Users..." -ForegroundColor Yellow
            $edrPoliciesFound_AllUsers = [System.Collections.ArrayList]::new()
            $processedEDRIds_AllUsers = [System.Collections.Generic.HashSet[string]]::new() # Reset for this type

            $configPoliciesForEDR_AllUsers = Get-IntuneEntities -EntityType "configurationPolicies"
            $matchingConfigPoliciesEDR_AllUsers = $configPoliciesForEDR_AllUsers | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityEndpointDetectionAndResponse' }

            if ($matchingConfigPoliciesEDR_AllUsers) {
                foreach ($policy in $matchingConfigPoliciesEDR_AllUsers) {
                    if ($processedEDRIds_AllUsers.Add($policy.id)) {
                        $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                        if ($assignments | Where-Object { $_.Reason -eq "All Users" }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                            [void]$edrPoliciesFound_AllUsers.Add($policy)
                        }
                    }
                }
            }

            $allIntentsForEDR_AllUsers = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $matchingIntentsEDR_AllUsers = $allIntentsForEDR_AllUsers | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityEndpointDetectionAndResponse' }

            if ($matchingIntentsEDR_AllUsers) {
                foreach ($policy in $matchingIntentsEDR_AllUsers) {
                    if ($processedEDRIds_AllUsers.Add($policy.id)) {
                        $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        if ($assignmentsResponse.value | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget' }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                            [void]$edrPoliciesFound_AllUsers.Add($policy)
                        }
                    }
                }
            }
            $allUsersAssignments.EndpointDetectionProfiles = $edrPoliciesFound_AllUsers

            Write-Host "Fetching ASR Policies assigned to All Users..." -ForegroundColor Yellow
            $asrPoliciesFound_AllUsers = [System.Collections.ArrayList]::new()
            $processedASRIds_AllUsers = [System.Collections.Generic.HashSet[string]]::new() # Reset for this type

            $configPoliciesForASR_AllUsers = Get-IntuneEntities -EntityType "configurationPolicies"
            $matchingConfigPoliciesASR_AllUsers = $configPoliciesForASR_AllUsers | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAttackSurfaceReductionRules' }

            if ($matchingConfigPoliciesASR_AllUsers) {
                foreach ($policy in $matchingConfigPoliciesASR_AllUsers) {
                    if ($processedASRIds_AllUsers.Add($policy.id)) {
                        $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                        if ($assignments | Where-Object { $_.Reason -eq "All Users" }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                            [void]$asrPoliciesFound_AllUsers.Add($policy)
                        }
                    }
                }
            }

            $allIntentsForASR_AllUsers = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $matchingIntentsASR_AllUsers = $allIntentsForASR_AllUsers | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAttackSurfaceReductionRules' }

            if ($matchingIntentsASR_AllUsers) {
                foreach ($policy in $matchingIntentsASR_AllUsers) {
                    if ($processedASRIds_AllUsers.Add($policy.id)) {
                        $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        if ($assignmentsResponse.value | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget' }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                            [void]$asrPoliciesFound_AllUsers.Add($policy)
                        }
                    }
                }
            }
            $allUsersAssignments.AttackSurfaceProfiles = $asrPoliciesFound_AllUsers
            
            Write-Host "Fetching Autopilot Deployment Profiles assigned to All Users..." -ForegroundColor Yellow
            $autoProfilesAU = Get-IntuneEntities -EntityType "windowsAutopilotDeploymentProfiles"
            foreach ($profile in $autoProfilesAU) {
                $assignments = Get-IntuneAssignments -EntityType "windowsAutopilotDeploymentProfiles" -EntityId $profile.id
                if ($assignments | Where-Object { $_.Reason -eq "All Users" }) {
                    $profile | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                    $allUsersAssignments.DeploymentProfiles += $profile
                }
            }

            Write-Host "Fetching Enrollment Status Page Profiles assigned to All Users..." -ForegroundColor Yellow
            $enrollmentConfigsAU = Get-IntuneEntities -EntityType "deviceEnrollmentConfigurations"
            $espProfilesAU = $enrollmentConfigsAU | Where-Object { $_.'@odata.type' -match 'EnrollmentCompletionPageConfiguration' }
            foreach ($esp in $espProfilesAU) {
                $assignments = Get-IntuneAssignments -EntityType "deviceEnrollmentConfigurations" -EntityId $esp.id
                if ($assignments | Where-Object { $_.Reason -eq "All Users" }) {
                    $esp | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Users" -Force
                    $allUsersAssignments.ESPProfiles += $esp
                }
            }

            Write-Host "`nPolicies Assigned to All Users:" -ForegroundColor Green

            Write-Host "`n------- Device Configurations -------" -ForegroundColor Cyan
            if ($allUsersAssignments.DeviceConfigs.Count -eq 0) {
                Write-Host "No Device Configurations assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($config in $allUsersAssignments.DeviceConfigs) {
                    $configName = if ([string]::IsNullOrWhiteSpace($config.name)) { $config.displayName } else { $config.name }
                    Write-Host "Device Configuration Name: $configName, Configuration ID: $($config.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Device Configuration" -Items @($config) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- Settings Catalog Policies -------" -ForegroundColor Cyan
            if ($allUsersAssignments.SettingsCatalog.Count -eq 0) {
                Write-Host "No Settings Catalog Policies assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($policy in $allUsersAssignments.SettingsCatalog) {
                    $policyName = if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                    Write-Host "Settings Catalog Policy Name: $policyName, Policy ID: $($policy.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Settings Catalog Policy" -Items @($policy) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- Administrative Templates -------" -ForegroundColor Cyan
            if ($allUsersAssignments.AdminTemplates.Count -eq 0) {
                Write-Host "No Administrative Templates assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($template in $allUsersAssignments.AdminTemplates) {
                    $templateName = if ([string]::IsNullOrWhiteSpace($template.name)) { $template.displayName } else { $template.name }
                    Write-Host "Administrative Template Name: $templateName, Template ID: $($template.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Administrative Template" -Items @($template) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- Compliance Policies -------" -ForegroundColor Cyan
            if ($allUsersAssignments.CompliancePolicies.Count -eq 0) {
                Write-Host "No Compliance Policies assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($policy in $allUsersAssignments.CompliancePolicies) {
                    $policyName = if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                    Write-Host "Compliance Policy Name: $policyName, Policy ID: $($policy.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Compliance Policy" -Items @($policy) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- App Protection Policies -------" -ForegroundColor Cyan
            if ($allUsersAssignments.AppProtectionPolicies.Count -eq 0) {
                Write-Host "No App Protection Policies assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($policy in $allUsersAssignments.AppProtectionPolicies) {
                    $policyName = $policy.displayName
                    $policyType = switch ($policy.'@odata.type') {
                        "#microsoft.graph.androidManagedAppProtection" { "Android" }
                        "#microsoft.graph.iosManagedAppProtection" { "iOS" }
                        "#microsoft.graph.windowsManagedAppProtection" { "Windows" }
                        default { "Unknown" }
                    }
                    Write-Host "App Protection Policy Name: $policyName, Policy ID: $($policy.id), Type: $policyType" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "App Protection Policy" -Items @($policy) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- App Configuration Policies -------" -ForegroundColor Cyan
            if ($allUsersAssignments.AppConfigurationPolicies.Count -eq 0) {
                Write-Host "No App Configuration Policies assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($policy in $allUsersAssignments.AppConfigurationPolicies) {
                    $policyName = if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                    Write-Host "App Configuration Policy Name: $policyName, Policy ID: $($policy.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "App Configuration Policy" -Items @($policy) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- Platform Scripts -------" -ForegroundColor Cyan
            if ($allUsersAssignments.PlatformScripts.Count -eq 0) {
                Write-Host "No Platform Scripts assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($script in $allUsersAssignments.PlatformScripts) {
                    $scriptName = if ([string]::IsNullOrWhiteSpace($script.name)) { $script.displayName } else { $script.name }
                    Write-Host "Script Name: $scriptName, Script ID: $($script.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Platform Scripts" -Items @($script) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- Proactive Remediation Scripts -------" -ForegroundColor Cyan
            if ($allUsersAssignments.HealthScripts.Count -eq 0) {
                Write-Host "No Proactive Remediation Scripts assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($script in $allUsersAssignments.HealthScripts) {
                    $scriptName = if ([string]::IsNullOrWhiteSpace($script.name)) { $script.displayName } else { $script.name }
                    Write-Host "Script Name: $scriptName, Script ID: $($script.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Proactive Remediation Scripts" -Items @($script) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- Required Apps -------" -ForegroundColor Cyan
            if ($allUsersAssignments.RequiredApps.Count -eq 0) {
                Write-Host "No Required Apps assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($app in $allUsersAssignments.RequiredApps) {
                    $appName = $app.displayName
                    Write-Host "App Name: $appName, App ID: $($app.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Required Apps" -Items @($app) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- Available Apps -------" -ForegroundColor Cyan
            if ($allUsersAssignments.AvailableApps.Count -eq 0) {
                Write-Host "No Available Apps assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($app in $allUsersAssignments.AvailableApps) {
                    $appName = $app.displayName
                    Write-Host "App Name: $appName, App ID: $($app.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Available Apps" -Items @($app) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- Uninstall Apps -------" -ForegroundColor Cyan
            if ($allUsersAssignments.UninstallApps.Count -eq 0) {
                Write-Host "No Uninstall Apps assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($app in $allUsersAssignments.UninstallApps) {
                    $appName = $app.displayName
                    Write-Host "App Name: $appName, App ID: $($app.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Uninstall Apps" -Items @($app) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- Endpoint Security - Antivirus Profiles -------" -ForegroundColor Cyan
            if ($allUsersAssignments.AntivirusProfiles.Count -eq 0) {
                Write-Host "No Antivirus Profiles assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $allUsersAssignments.AntivirusProfiles) {
                    $profileNameForDisplay = if ($profile.displayName) { $profile.displayName } else { $profile.name }
                    Write-Host "Antivirus Profile Name: $profileNameForDisplay, Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Antivirus" -Items @($profile) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- Endpoint Security - Disk Encryption Profiles -------" -ForegroundColor Cyan
            if ($allUsersAssignments.DiskEncryptionProfiles.Count -eq 0) {
                Write-Host "No Disk Encryption Profiles assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $allUsersAssignments.DiskEncryptionProfiles) {
                    $profileNameForDisplay = if ($profile.displayName) { $profile.displayName } else { $profile.name }
                    Write-Host "Disk Encryption Profile Name: $profileNameForDisplay, Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Disk Encryption" -Items @($profile) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- Endpoint Security - Firewall Profiles -------" -ForegroundColor Cyan
            if ($allUsersAssignments.FirewallProfiles.Count -eq 0) {
                Write-Host "No Firewall Profiles assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $allUsersAssignments.FirewallProfiles) {
                    $profileNameForDisplay = if ($profile.displayName) { $profile.displayName } else { $profile.name }
                    Write-Host "Firewall Profile Name: $profileNameForDisplay, Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Firewall" -Items @($profile) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- Endpoint Security - EDR Profiles -------" -ForegroundColor Cyan
            if ($allUsersAssignments.EndpointDetectionProfiles.Count -eq 0) {
                Write-Host "No EDR Profiles assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $allUsersAssignments.EndpointDetectionProfiles) {
                    $profileNameForDisplay = if ($profile.displayName) { $profile.displayName } else { $profile.name }
                    Write-Host "EDR Profile Name: $profileNameForDisplay, Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - EDR" -Items @($profile) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- Endpoint Security - ASR Profiles -------" -ForegroundColor Cyan
            if ($allUsersAssignments.AttackSurfaceProfiles.Count -eq 0) {
                Write-Host "No ASR Profiles assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $allUsersAssignments.AttackSurfaceProfiles) {
                    $profileNameForDisplay = if ($profile.displayName) { $profile.displayName } else { $profile.name }
                    Write-Host "ASR Profile Name: $profileNameForDisplay, Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - ASR" -Items @($profile) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- Autopilot Deployment Profiles -------" -ForegroundColor Cyan
            if ($allUsersAssignments.DeploymentProfiles.Count -eq 0) {
                Write-Host "No Autopilot Deployment Profiles assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $allUsersAssignments.DeploymentProfiles) {
                    $profileName = if ([string]::IsNullOrWhiteSpace($profile.name)) { $profile.displayName } else { $profile.name }
                    Write-Host "Autopilot Deployment Profile Name: $profileName, Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Autopilot Deployment Profile" -Items @($profile) -AssignmentReason "All Users"
                }
            }

            Write-Host "`n------- Enrollment Status Page Profiles -------" -ForegroundColor Cyan
            if ($allUsersAssignments.ESPProfiles.Count -eq 0) {
                Write-Host "No Enrollment Status Page Profiles assigned to All Users" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $allUsersAssignments.ESPProfiles) {
                    $profileName = if ([string]::IsNullOrWhiteSpace($profile.name)) { $profile.displayName } else { $profile.name }
                    Write-Host "Enrollment Status Page Profile Name: $profileName, Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Enrollment Status Page Profile" -Items @($profile) -AssignmentReason "All Users"
                }
            }

            Export-ResultsIfRequested -ExportData $exportData -DefaultFileName "IntuneAllUsersAssignments.csv" -ForceExport:$ExportToCSV -CustomExportPath $ExportPath
        }     
        '6' {
            Write-Host "Fetching all 'All Devices' assignments..." -ForegroundColor Green
            $exportData = [System.Collections.ArrayList]::new()

            $allDevicesAssignments = @{
                DeviceConfigs             = @()
                SettingsCatalog           = @()
                AdminTemplates            = @()
                CompliancePolicies        = @()
                AppProtectionPolicies     = @()
                AppConfigurationPolicies  = @()
                PlatformScripts           = @()
                HealthScripts             = @()
                RequiredApps              = @()
                AvailableApps             = @()
                UninstallApps             = @()
                DeploymentProfiles        = @()
                ESPProfiles               = @()
                AntivirusProfiles         = @()
                DiskEncryptionProfiles    = @()
                FirewallProfiles          = @()
                EndpointDetectionProfiles = @()
                AttackSurfaceProfiles     = @()
            }

            Write-Host "Fetching Device Configurations..." -ForegroundColor Yellow
            $deviceConfigs = Get-IntuneEntities -EntityType "deviceConfigurations"
            foreach ($config in $deviceConfigs) {
                $assignments = Get-IntuneAssignments -EntityType "deviceConfigurations" -EntityId $config.id
                if ($assignments | Where-Object { $_.Reason -eq "All Devices" }) {
                    $config | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                    $allDevicesAssignments.DeviceConfigs += $config
                }
            }

            Write-Host "Fetching Settings Catalog Policies..." -ForegroundColor Yellow
            $settingsCatalog = Get-IntuneEntities -EntityType "configurationPolicies"
            foreach ($policy in $settingsCatalog) {
                $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                if ($assignments | Where-Object { $_.Reason -eq "All Devices" }) {
                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                    $allDevicesAssignments.SettingsCatalog += $policy
                }
            }

            Write-Host "Fetching Administrative Templates..." -ForegroundColor Yellow
            $adminTemplates = Get-IntuneEntities -EntityType "groupPolicyConfigurations"
            foreach ($template in $adminTemplates) {
                $assignments = Get-IntuneAssignments -EntityType "groupPolicyConfigurations" -EntityId $template.id
                if ($assignments | Where-Object { $_.Reason -eq "All Devices" }) {
                    $template | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                    $allDevicesAssignments.AdminTemplates += $template
                }
            }

            Write-Host "Fetching Compliance Policies..." -ForegroundColor Yellow
            $compliancePolicies = Get-IntuneEntities -EntityType "deviceCompliancePolicies"
            foreach ($policy in $compliancePolicies) {
                $assignments = Get-IntuneAssignments -EntityType "deviceCompliancePolicies" -EntityId $policy.id
                if ($assignments | Where-Object { $_.Reason -eq "All Devices" }) {
                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                    $allDevicesAssignments.CompliancePolicies += $policy
                }
            }

            Write-Host "Fetching App Protection Policies..." -ForegroundColor Yellow
            $appProtectionPolicies = Get-IntuneEntities -EntityType "deviceAppManagement/managedAppPolicies"
            foreach ($policy in $appProtectionPolicies) {
                $policyType = $policy.'@odata.type'
                $assignmentsUri = switch ($policyType) {
                    "#microsoft.graph.androidManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/androidManagedAppProtections('$($policy.id)')/assignments" }
                    "#microsoft.graph.iosManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/iosManagedAppProtections('$($policy.id)')/assignments" }
                    "#microsoft.graph.windowsManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/windowsManagedAppProtections('$($policy.id)')/assignments" }
                    default { $null }
                }

                if ($assignmentsUri) {
                    try {
                        $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get
                        $hasAllDevices = $false
                        foreach ($assignment in $assignmentResponse.value) {
                            if ($assignment.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget') {
                                $hasAllDevices = $true
                                break
                            }
                        }
                        if ($hasAllDevices) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                            $allDevicesAssignments.AppProtectionPolicies += $policy
                        }
                    }
                    catch {
                        Write-Host "Error fetching assignments for policy $($policy.displayName): $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }

            Write-Host "Fetching App Configuration Policies..." -ForegroundColor Yellow
            $appConfigPolicies = Get-IntuneEntities -EntityType "deviceAppManagement/mobileAppConfigurations"
            foreach ($policy in $appConfigPolicies) {
                $assignments = Get-IntuneAssignments -EntityType "mobileAppConfigurations" -EntityId $policy.id
                if ($assignments | Where-Object { $_.Reason -eq "All Devices" }) {
                    $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                    $allDevicesAssignments.AppConfigurationPolicies += $policy
                }
            }

            Write-Host "Fetching Applications..." -ForegroundColor Yellow
            $appUri = "$GraphEndpoint/beta/deviceAppManagement/mobileApps?`$filter=isAssigned eq true"
            $appResponse = Invoke-MgGraphRequest -Uri $appUri -Method Get
            $allApps = $appResponse.value
            while ($appResponse.'@odata.nextLink') {
                $appResponse = Invoke-MgGraphRequest -Uri $appResponse.'@odata.nextLink' -Method Get
                $allApps += $appResponse.value
            }
            $totalApps = $allApps.Count

            foreach ($app in $allApps) {
                if ($app.isFeatured -or $app.isBuiltIn) {
                    continue
                }

                $appId = $app.id
                $assignmentsUri = "$GraphEndpoint/beta/deviceAppManagement/mobileApps('$appId')/assignments"
                $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get

                foreach ($assignment in $assignmentResponse.value) {
                    if ($assignment.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget') {
                        $appWithReason = $app.PSObject.Copy()
                        $appWithReason | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                        switch ($assignment.intent) {
                            "required" { $allDevicesAssignments.RequiredApps += $appWithReason; break }
                            "available" { $allDevicesAssignments.AvailableApps += $appWithReason; break }
                            "uninstall" { $allDevicesAssignments.UninstallApps += $appWithReason; break }
                        }
                        break
                    }
                }
            }

            Write-Host "Fetching Platform Scripts..." -ForegroundColor Yellow
            $platformScripts = Get-IntuneEntities -EntityType "deviceManagementScripts"
            foreach ($script in $platformScripts) {
                $assignments = Get-IntuneAssignments -EntityType "deviceManagementScripts" -EntityId $script.id
                if ($assignments | Where-Object { $_.Reason -eq "All Devices" }) {
                    $script | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                    $allDevicesAssignments.PlatformScripts += $script
                }
            }

            Write-Host "Fetching Proactive Remediation Scripts..." -ForegroundColor Yellow
            $healthScripts = Get-IntuneEntities -EntityType "deviceHealthScripts"
            foreach ($script in $healthScripts) {
                $assignments = Get-IntuneAssignments -EntityType "deviceHealthScripts" -EntityId $script.id
                if ($assignments | Where-Object { $_.Reason -eq "All Devices" }) {
                    $script | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                    $allDevicesAssignments.HealthScripts += $script
                }
            }

            Write-Host "Fetching Autopilot Deployment Profiles assigned to All Devices..." -ForegroundColor Yellow
            $autoProfilesAD = Get-IntuneEntities -EntityType "windowsAutopilotDeploymentProfiles"
            foreach ($profile in $autoProfilesAD) {
                $assignments = Get-IntuneAssignments -EntityType "windowsAutopilotDeploymentProfiles" -EntityId $profile.id
                if ($assignments | Where-Object { $_.Reason -eq "All Devices" }) {
                    $profile | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                    $allDevicesAssignments.DeploymentProfiles += $profile
                }
            }

            Write-Host "Fetching Enrollment Status Page Profiles assigned to All Devices..." -ForegroundColor Yellow
            $enrollmentConfigsAD = Get-IntuneEntities -EntityType "deviceEnrollmentConfigurations"
            $espProfilesAD = $enrollmentConfigsAD | Where-Object { $_.'@odata.type' -match 'EnrollmentCompletionPageConfiguration' }
            foreach ($esp in $espProfilesAD) {
                $assignments = Get-IntuneAssignments -EntityType "deviceEnrollmentConfigurations" -EntityId $esp.id
                if ($assignments | Where-Object { $_.Reason -eq "All Devices" }) {
                    $esp | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                    $allDevicesAssignments.ESPProfiles += $esp
                }
            }

            Write-Host "Fetching Antivirus Policies assigned to All Devices..." -ForegroundColor Yellow
            $antivirusPoliciesFound_AllDevices = [System.Collections.ArrayList]::new()
            $processedAntivirusIds_AllDevices = [System.Collections.Generic.HashSet[string]]::new()

            $configPoliciesForAntivirus_AllDevices = Get-IntuneEntities -EntityType "configurationPolicies"
            $matchingConfigPoliciesAntivirus_AllDevices = $configPoliciesForAntivirus_AllDevices | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAntivirus' }
            if ($matchingConfigPoliciesAntivirus_AllDevices) {
                foreach ($policy in $matchingConfigPoliciesAntivirus_AllDevices) {
                    if ($processedAntivirusIds_AllDevices.Add($policy.id)) {
                        $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                        if ($assignments | Where-Object { $_.Reason -eq "All Devices" }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                            [void]$antivirusPoliciesFound_AllDevices.Add($policy)
                        }
                    }
                }
            }

            $allIntentsForAntivirus_AllDevices = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $matchingIntentsAntivirus_AllDevices = $allIntentsForAntivirus_AllDevices | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAntivirus' }
            if ($matchingIntentsAntivirus_AllDevices) {
                foreach ($policy in $matchingIntentsAntivirus_AllDevices) {
                    if ($processedAntivirusIds_AllDevices.Add($policy.id)) {
                        $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        if ($assignmentsResponse.value | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget' }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                            [void]$antivirusPoliciesFound_AllDevices.Add($policy)
                        }
                    }
                }
            }
            $allDevicesAssignments.AntivirusProfiles = $antivirusPoliciesFound_AllDevices

            Write-Host "Fetching Disk Encryption Policies assigned to All Devices..." -ForegroundColor Yellow
            $diskEncryptionPoliciesFound_AllDevices = [System.Collections.ArrayList]::new()
            $processedDiskEncryptionIds_AllDevices = [System.Collections.Generic.HashSet[string]]::new()

            $configPoliciesForDiskEnc_AllDevices = Get-IntuneEntities -EntityType "configurationPolicies"
            $matchingConfigPoliciesDiskEnc_AllDevices = $configPoliciesForDiskEnc_AllDevices | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityDiskEncryption' }
            if ($matchingConfigPoliciesDiskEnc_AllDevices) {
                foreach ($policy in $matchingConfigPoliciesDiskEnc_AllDevices) {
                    if ($processedDiskEncryptionIds_AllDevices.Add($policy.id)) {
                        $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                        if ($assignments | Where-Object { $_.Reason -eq "All Devices" }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                            [void]$diskEncryptionPoliciesFound_AllDevices.Add($policy)
                        }
                    }
                }
            }

            $allIntentsForDiskEnc_AllDevices = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $matchingIntentsDiskEnc_AllDevices = $allIntentsForDiskEnc_AllDevices | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityDiskEncryption' }
            if ($matchingIntentsDiskEnc_AllDevices) {
                foreach ($policy in $matchingIntentsDiskEnc_AllDevices) {
                    if ($processedDiskEncryptionIds_AllDevices.Add($policy.id)) {
                        $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        if ($assignmentsResponse.value | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget' }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                            [void]$diskEncryptionPoliciesFound_AllDevices.Add($policy)
                        }
                    }
                }
            }
            $allDevicesAssignments.DiskEncryptionProfiles = $diskEncryptionPoliciesFound_AllDevices

            Write-Host "Fetching Firewall Policies assigned to All Devices..." -ForegroundColor Yellow
            $firewallPoliciesFound_AllDevices = [System.Collections.ArrayList]::new()
            $processedFirewallIds_AllDevices = [System.Collections.Generic.HashSet[string]]::new()

            $configPoliciesForFirewall_AllDevices = Get-IntuneEntities -EntityType "configurationPolicies"
            $matchingConfigPoliciesFirewall_AllDevices = $configPoliciesForFirewall_AllDevices | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityFirewall' }
            if ($matchingConfigPoliciesFirewall_AllDevices) {
                foreach ($policy in $matchingConfigPoliciesFirewall_AllDevices) {
                    if ($processedFirewallIds_AllDevices.Add($policy.id)) {
                        $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                        if ($assignments | Where-Object { $_.Reason -eq "All Devices" }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                            [void]$firewallPoliciesFound_AllDevices.Add($policy)
                        }
                    }
                }
            }

            $allIntentsForFirewall_AllDevices = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $matchingIntentsFirewall_AllDevices = $allIntentsForFirewall_AllDevices | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityFirewall' }
            if ($matchingIntentsFirewall_AllDevices) {
                foreach ($policy in $matchingIntentsFirewall_AllDevices) {
                    if ($processedFirewallIds_AllDevices.Add($policy.id)) {
                        $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        if ($assignmentsResponse.value | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget' }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                            [void]$firewallPoliciesFound_AllDevices.Add($policy)
                        }
                    }
                }
            }
            $allDevicesAssignments.FirewallProfiles = $firewallPoliciesFound_AllDevices

            Write-Host "Fetching EDR Policies assigned to All Devices..." -ForegroundColor Yellow
            $edrPoliciesFound_AllDevices = [System.Collections.ArrayList]::new()
            $processedEDRIds_AllDevices = [System.Collections.Generic.HashSet[string]]::new()

            $configPoliciesForEDR_AllDevices = Get-IntuneEntities -EntityType "configurationPolicies"
            $matchingConfigPoliciesEDR_AllDevices = $configPoliciesForEDR_AllDevices | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityEndpointDetectionAndResponse' }
            if ($matchingConfigPoliciesEDR_AllDevices) {
                foreach ($policy in $matchingConfigPoliciesEDR_AllDevices) {
                    if ($processedEDRIds_AllDevices.Add($policy.id)) {
                        $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                        if ($assignments | Where-Object { $_.Reason -eq "All Devices" }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                            [void]$edrPoliciesFound_AllDevices.Add($policy)
                        }
                    }
                }
            }

            $allIntentsForEDR_AllDevices = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $matchingIntentsEDR_AllDevices = $allIntentsForEDR_AllDevices | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityEndpointDetectionAndResponse' }
            if ($matchingIntentsEDR_AllDevices) {
                foreach ($policy in $matchingIntentsEDR_AllDevices) {
                    if ($processedEDRIds_AllDevices.Add($policy.id)) {
                        $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        if ($assignmentsResponse.value | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget' }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                            [void]$edrPoliciesFound_AllDevices.Add($policy)
                        }
                    }
                }
            }
            $allDevicesAssignments.EndpointDetectionProfiles = $edrPoliciesFound_AllDevices

            Write-Host "Fetching ASR Policies assigned to All Devices..." -ForegroundColor Yellow
            $asrPoliciesFound_AllDevices = [System.Collections.ArrayList]::new()
            $processedASRIds_AllDevices = [System.Collections.Generic.HashSet[string]]::new()

            $configPoliciesForASR_AllDevices = Get-IntuneEntities -EntityType "configurationPolicies"
            $matchingConfigPoliciesASR_AllDevices = $configPoliciesForASR_AllDevices | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAttackSurfaceReductionRules' }
            if ($matchingConfigPoliciesASR_AllDevices) {
                foreach ($policy in $matchingConfigPoliciesASR_AllDevices) {
                    if ($processedASRIds_AllDevices.Add($policy.id)) {
                        $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                        if ($assignments | Where-Object { $_.Reason -eq "All Devices" }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                            [void]$asrPoliciesFound_AllDevices.Add($policy)
                        }
                    }
                }
            }

            $allIntentsForASR_AllDevices = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $matchingIntentsASR_AllDevices = $allIntentsForASR_AllDevices | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAttackSurfaceReductionRules' }
            if ($matchingIntentsASR_AllDevices) {
                foreach ($policy in $matchingIntentsASR_AllDevices) {
                    if ($processedASRIds_AllDevices.Add($policy.id)) {
                        $assignmentsResponse = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        if ($assignmentsResponse.value | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget' }) {
                            $policy | Add-Member -NotePropertyName 'AssignmentReason' -NotePropertyValue "All Devices" -Force
                            [void]$asrPoliciesFound_AllDevices.Add($policy)
                        }
                    }
                }
            }
            $allDevicesAssignments.AttackSurfaceProfiles = $asrPoliciesFound_AllDevices

            Write-Host "`nPolicies Assigned to All Devices:" -ForegroundColor Green

            Write-Host "`n------- Device Configurations -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.DeviceConfigs.Count -eq 0) {
                Write-Host "No Device Configurations assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($config in $allDevicesAssignments.DeviceConfigs) {
                    $configName = if ([string]::IsNullOrWhiteSpace($config.name)) { $config.displayName } else { $config.name }
                    Write-Host "Device Configuration Name: $configName, Configuration ID: $($config.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Device Configuration" -Items @($config) -AssignmentReason "All Devices"
                }
            }

            Write-Host "`n------- Settings Catalog Policies -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.SettingsCatalog.Count -eq 0) {
                Write-Host "No Settings Catalog Policies assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($policy in $allDevicesAssignments.SettingsCatalog) {
                    $policyName = if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                    Write-Host "Settings Catalog Policy Name: $policyName, Policy ID: $($policy.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Settings Catalog Policy" -Items @($policy) -AssignmentReason "All Devices"
                }
            }

            Write-Host "`n------- Administrative Templates -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.AdminTemplates.Count -eq 0) {
                Write-Host "No Administrative Templates assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($template in $allDevicesAssignments.AdminTemplates) {
                    $templateName = if ([string]::IsNullOrWhiteSpace($template.name)) { $template.displayName } else { $template.name }
                    Write-Host "Administrative Template Name: $templateName, Template ID: $($template.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Administrative Template" -Items @($template) -AssignmentReason "All Devices"
                }
            }

            Write-Host "`n------- Compliance Policies -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.CompliancePolicies.Count -eq 0) {
                Write-Host "No Compliance Policies assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($policy in $allDevicesAssignments.CompliancePolicies) {
                    $policyName = if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                    Write-Host "Compliance Policy Name: $policyName, Policy ID: $($policy.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Compliance Policy" -Items @($policy) -AssignmentReason "All Devices"
                }
            }

            Write-Host "`n------- App Protection Policies -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.AppProtectionPolicies.Count -eq 0) {
                Write-Host "No App Protection Policies assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($policy in $allDevicesAssignments.AppProtectionPolicies) {
                    $policyName = $policy.displayName
                    $policyType = switch ($policy.'@odata.type') {
                        "#microsoft.graph.androidManagedAppProtection" { "Android" }
                        "#microsoft.graph.iosManagedAppProtection" { "iOS" }
                        "#microsoft.graph.windowsManagedAppProtection" { "Windows" }
                        default { "Unknown" }
                    }
                    Write-Host "App Protection Policy Name: $policyName, Policy ID: $($policy.id), Type: $policyType" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "App Protection Policy" -Items @($policy) -AssignmentReason "All Devices"
                }
            }

            Write-Host "`n------- App Configuration Policies -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.AppConfigurationPolicies.Count -eq 0) {
                Write-Host "No App Configuration Policies assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($policy in $allDevicesAssignments.AppConfigurationPolicies) {
                    $policyName = if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                    Write-Host "App Configuration Policy Name: $policyName, Policy ID: $($policy.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "App Configuration Policy" -Items @($policy) -AssignmentReason "All Devices"
                }
            }

            Write-Host "`n------- Platform Scripts -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.PlatformScripts.Count -eq 0) {
                Write-Host "No Platform Scripts assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($script in $allDevicesAssignments.PlatformScripts) {
                    $scriptName = if ([string]::IsNullOrWhiteSpace($script.name)) { $script.displayName } else { $script.name }
                    Write-Host "Script Name: $scriptName, Script ID: $($script.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Platform Scripts" -Items @($script) -AssignmentReason "All Devices"
                }
            }

            Write-Host "`n------- Proactive Remediation Scripts -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.HealthScripts.Count -eq 0) {
                Write-Host "No Proactive Remediation Scripts assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($script in $allDevicesAssignments.HealthScripts) {
                    $scriptName = if ([string]::IsNullOrWhiteSpace($script.name)) { $script.displayName } else { $script.name }
                    Write-Host "Script Name: $scriptName, Script ID: $($script.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Proactive Remediation Scripts" -Items @($script) -AssignmentReason "All Devices"
                }
            }

            Write-Host "`n------- Required Apps -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.RequiredApps.Count -eq 0) {
                Write-Host "No Required Apps assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($app in $allDevicesAssignments.RequiredApps) {
                    $appName = $app.displayName
                    Write-Host "App Name: $appName, App ID: $($app.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Required Apps" -Items @($app) -AssignmentReason "All Devices"
                }
            }

            Write-Host "`n------- Available Apps -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.AvailableApps.Count -eq 0) {
                Write-Host "No Available Apps assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($app in $allDevicesAssignments.AvailableApps) {
                    $appName = $app.displayName
                    Write-Host "App Name: $appName, App ID: $($app.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Available Apps" -Items @($app) -AssignmentReason "All Devices"
                }
            }

            Write-Host "`n------- Uninstall Apps -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.UninstallApps.Count -eq 0) {
                Write-Host "No Uninstall Apps assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($app in $allDevicesAssignments.UninstallApps) {
                    $appName = $app.displayName
                    Write-Host "App Name: $appName, App ID: $($app.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Uninstall Apps" -Items @($app) -AssignmentReason "All Devices"
                }
            }

            Write-Host "`n------- Endpoint Security - Antivirus Profiles -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.AntivirusProfiles.Count -eq 0) {
                Write-Host "No Antivirus Profiles assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $allDevicesAssignments.AntivirusProfiles) {
                    Write-Host "Antivirus Profile Name: $($profile.displayName), Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Antivirus" -Items @($profile) -AssignmentReason "All Devices"
                }
            }

            Write-Host "`n------- Endpoint Security - Disk Encryption Profiles -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.DiskEncryptionProfiles.Count -eq 0) {
                Write-Host "No Disk Encryption Profiles assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $allDevicesAssignments.DiskEncryptionProfiles) {
                    Write-Host "Disk Encryption Profile Name: $($profile.displayName), Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Disk Encryption" -Items @($profile) -AssignmentReason "All Devices"
                }
            }

            Write-Host "`n------- Endpoint Security - Firewall Profiles -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.FirewallProfiles.Count -eq 0) {
                Write-Host "No Firewall Profiles assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $allDevicesAssignments.FirewallProfiles) {
                    Write-Host "Firewall Profile Name: $($profile.displayName), Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Firewall" -Items @($profile) -AssignmentReason "All Devices"
                }
            }

            Write-Host "`n------- Endpoint Security - EDR Profiles -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.EndpointDetectionProfiles.Count -eq 0) {
                Write-Host "No EDR Profiles assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $allDevicesAssignments.EndpointDetectionProfiles) {
                    $profileNameForDisplay = if (-not [string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.displayName } elseif (-not [string]::IsNullOrWhiteSpace($profile.name)) { $profile.name } else { "Unnamed EDR Profile" }
                    Write-Host "EDR Profile Name: $profileNameForDisplay, Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - EDR" -Items @($profile) -AssignmentReason "All Devices"
                }
            }

            Write-Host "`n------- Endpoint Security - ASR Profiles -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.AttackSurfaceProfiles.Count -eq 0) {
                Write-Host "No ASR Profiles assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $allDevicesAssignments.AttackSurfaceProfiles) {
                    Write-Host "ASR Profile Name: $($profile.displayName), Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - ASR" -Items @($profile) -AssignmentReason "All Devices"
                }
            }

            Write-Host "`n------- Autopilot Deployment Profiles -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.DeploymentProfiles.Count -eq 0) {
                Write-Host "No Autopilot Deployment Profiles assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $allDevicesAssignments.DeploymentProfiles) {
                    $profileName = if ([string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.name } else { $profile.displayName }
                    Write-Host "Deployment Profile Name: $profileName, Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Autopilot Deployment Profile" -Items @($profile) -AssignmentReason "All Devices"
                }
            }
            
            Write-Host "`n------- Enrollment Status Page Profiles -------" -ForegroundColor Cyan
            if ($allDevicesAssignments.ESPProfiles.Count -eq 0) {
                Write-Host "No Enrollment Status Page Profiles assigned to All Devices" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $allDevicesAssignments.ESPProfiles) {
                    $profileName = if ([string]::IsNullOrWhiteSpace($profile.displayName)) { $profile.name } else { $profile.displayName }
                    Write-Host "Enrollment Status Page Name: $profileName, Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Enrollment Status Page" -Items @($profile) -AssignmentReason "All Devices"
                }
            }

            Export-ResultsIfRequested -ExportData $exportData -DefaultFileName "IntuneAllDevicesAssignments.csv" -ForceExport:$ExportToCSV -CustomExportPath $ExportPath
        }            
         '7' {
                Write-Host "Generating HTML Report..." -ForegroundColor Green

                try {
                    $scriptPath = Join-Path $PSScriptRoot "html.ps1"

                    if (-Not (Test-Path $scriptPath)) {
                        throw "The script file $scriptPath was not found."
                    }

                    . $scriptPath  

                    $filePath = Join-Path (Get-Location) "IntuneReport.html"
                    Export-HTMLReport -FilePath $filePath

                    $openReport = Read-Host "Would you like to open the report now? (y/n)"
                    if ($openReport -eq 'y') {
                        Start-Process $filePath
                    }
                }
                catch {
                    Write-Host "Error: Failed to generate the HTML report. $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        '8' {
            Write-Host "Fetching policies without assignments..." -ForegroundColor Green
            $exportData = [System.Collections.ArrayList]::new()

            $unassignedPolicies = @{
                DeviceConfigs            = @()
                SettingsCatalog          = @()
                AdminTemplates           = @()
                CompliancePolicies       = @()
                AppProtectionPolicies    = @()
                AppConfigurationPolicies = @()
                PlatformScripts          = @()
                HealthScripts            = @()
            }

            Write-Host "Fetching Device Configurations..." -ForegroundColor Yellow
            $deviceConfigs = Get-IntuneEntities -EntityType "deviceConfigurations"
            foreach ($config in $deviceConfigs) {
                $assignments = Get-IntuneAssignments -EntityType "deviceConfigurations" -EntityId $config.id
                if ($assignments.Count -eq 0) {
                    $unassignedPolicies.DeviceConfigs += $config
                }
            }

            Write-Host "Fetching Settings Catalog Policies..." -ForegroundColor Yellow
            $settingsCatalog = Get-IntuneEntities -EntityType "configurationPolicies"
            foreach ($policy in $settingsCatalog) {
                $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                if ($assignments.Count -eq 0) {
                    $unassignedPolicies.SettingsCatalog += $policy
                }
            }

            Write-Host "Fetching Administrative Templates..." -ForegroundColor Yellow
            $adminTemplates = Get-IntuneEntities -EntityType "groupPolicyConfigurations"
            foreach ($template in $adminTemplates) {
                $assignments = Get-IntuneAssignments -EntityType "groupPolicyConfigurations" -EntityId $template.id
                if ($assignments.Count -eq 0) {
                    $unassignedPolicies.AdminTemplates += $template
                }
            }

            Write-Host "Fetching Compliance Policies..." -ForegroundColor Yellow
            $compliancePolicies = Get-IntuneEntities -EntityType "deviceCompliancePolicies"
            foreach ($policy in $compliancePolicies) {
                $assignments = Get-IntuneAssignments -EntityType "deviceCompliancePolicies" -EntityId $policy.id
                if ($assignments.Count -eq 0) {
                    $unassignedPolicies.CompliancePolicies += $policy
                }
            }

            Write-Host "Fetching App Protection Policies..." -ForegroundColor Yellow
            $appProtectionPolicies = Get-IntuneEntities -EntityType "deviceAppManagement/managedAppPolicies"
            foreach ($policy in $appProtectionPolicies) {
                $policyType = $policy.'@odata.type'
                $assignmentsUri = switch ($policyType) {
                    "#microsoft.graph.androidManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/androidManagedAppProtections('$($policy.id)')/assignments" }
                    "#microsoft.graph.iosManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/iosManagedAppProtections('$($policy.id)')/assignments" }
                    "#microsoft.graph.windowsManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/windowsManagedAppProtections('$($policy.id)')/assignments" }
                    default { $null }
                }

                if ($assignmentsUri) {
                    try {
                        $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get
                        if ($assignmentResponse.value.Count -eq 0) {
                            $unassignedPolicies.AppProtectionPolicies += $policy
                        }
                    }
                    catch {
                        Write-Host "Error fetching assignments for policy $($policy.displayName): $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }

            Write-Host "Fetching App Configuration Policies..." -ForegroundColor Yellow
            $appConfigPolicies = Get-IntuneEntities -EntityType "deviceAppManagement/mobileAppConfigurations"
            foreach ($policy in $appConfigPolicies) {
                $assignments = Get-IntuneAssignments -EntityType "mobileAppConfigurations" -EntityId $policy.id
                if ($assignments.Count -eq 0) {
                    $unassignedPolicies.AppConfigurationPolicies += $policy
                }
            }

            Write-Host "Fetching Platform Scripts..." -ForegroundColor Yellow
            $platformScripts = Get-IntuneEntities -EntityType "deviceManagementScripts"
            foreach ($script in $platformScripts) {
                $assignments = Get-IntuneAssignments -EntityType "deviceManagementScripts" -EntityId $script.id
                if ($assignments.Count -eq 0) {
                    $unassignedPolicies.PlatformScripts += $script
                }
            }

            Write-Host "Fetching Proactive Remediation Scripts..." -ForegroundColor Yellow
            $healthScripts = Get-IntuneEntities -EntityType "deviceHealthScripts"
            foreach ($script in $healthScripts) {
                $assignments = Get-IntuneAssignments -EntityType "deviceHealthScripts" -EntityId $script.id
                if ($assignments.Count -eq 0) {
                    $unassignedPolicies.HealthScripts += $script
                }
            }

            Write-Host "Fetching Antivirus Policies..." -ForegroundColor Yellow
            $allIntentsForAntivirusUnassigned = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $antivirusPolicies = $allIntentsForAntivirusUnassigned | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAntivirus' }
            if ($antivirusPolicies) {
                foreach ($policy in $antivirusPolicies) {
                    $assignments = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                    if ($assignments.value.Count -eq 0) {
                        $unassignedPolicies.AntivirusProfiles += $policy
                    }
                }
            }

            Write-Host "Fetching Disk Encryption Policies..." -ForegroundColor Yellow
            $allIntentsForDiskEncUnassigned = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $diskEncryptionPolicies = $allIntentsForDiskEncUnassigned | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityDiskEncryption' }
            if ($diskEncryptionPolicies) {
                foreach ($policy in $diskEncryptionPolicies) {
                    $assignments = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                    if ($assignments.value.Count -eq 0) {
                        $unassignedPolicies.DiskEncryptionProfiles += $policy
                    }
                }
            }

            Write-Host "Fetching Firewall Policies..." -ForegroundColor Yellow
            $allIntentsForFirewallUnassigned = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $firewallPolicies = $allIntentsForFirewallUnassigned | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityFirewall' }
            if ($firewallPolicies) {
                foreach ($policy in $firewallPolicies) {
                    $assignments = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                    if ($assignments.value.Count -eq 0) {
                        $unassignedPolicies.FirewallProfiles += $policy
                    }
                }
            }

            Write-Host "Fetching EDR Policies..." -ForegroundColor Yellow
            $allIntentsForEDRUnassigned = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $edrPolicies = $allIntentsForEDRUnassigned | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityEndpointDetectionAndResponse' }
            if ($edrPolicies) {
                foreach ($policy in $edrPolicies) {
                    $assignments = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                    if ($assignments.value.Count -eq 0) {
                        $unassignedPolicies.EndpointDetectionProfiles += $policy
                    }
                }
            }

            Write-Host "Fetching ASR Policies..." -ForegroundColor Yellow
            $allIntentsForASRUnassigned = Get-IntuneEntities -EntityType "deviceManagement/intents"
            $asrPolicies = $allIntentsForASRUnassigned | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAttackSurfaceReductionRules' }
            if ($asrPolicies) {
                foreach ($policy in $asrPolicies) {
                    $assignments = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                    if ($assignments.value.Count -eq 0) {
                        $unassignedPolicies.AttackSurfaceProfiles += $policy
                    }
                }
            }
            
            Write-Host "`nPolicies Without Assignments:" -ForegroundColor Green

            Write-Host "`n------- Device Configurations -------" -ForegroundColor Cyan
            if ($unassignedPolicies.DeviceConfigs.Count -eq 0) {
                Write-Host "No unassigned Device Configurations found" -ForegroundColor Gray
            }
            else {
                foreach ($config in $unassignedPolicies.DeviceConfigs) {
                    $configName = if ([string]::IsNullOrWhiteSpace($config.name)) { $config.displayName } else { $config.name }
                    Write-Host "Device Configuration Name: $configName, Configuration ID: $($config.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Device Configuration" -Items @($config) -AssignmentReason "No Assignment"
                }
            }

            Write-Host "`n------- Settings Catalog Policies -------" -ForegroundColor Cyan
            if ($unassignedPolicies.SettingsCatalog.Count -eq 0) {
                Write-Host "No unassigned Settings Catalog Policies found" -ForegroundColor Gray
            }
            else {
                foreach ($policy in $unassignedPolicies.SettingsCatalog) {
                    $policyName = if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                    Write-Host "Settings Catalog Policy Name: $policyName, Policy ID: $($policy.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Settings Catalog Policy" -Items @($policy) -AssignmentReason "No Assignment"
                }
            }

            Write-Host "`n------- Administrative Templates -------" -ForegroundColor Cyan
            if ($unassignedPolicies.AdminTemplates.Count -eq 0) {
                Write-Host "No unassigned Administrative Templates found" -ForegroundColor Gray
            }
            else {
                foreach ($template in $unassignedPolicies.AdminTemplates) {
                    $templateName = if ([string]::IsNullOrWhiteSpace($template.name)) { $template.displayName } else { $template.name }
                    Write-Host "Administrative Template Name: $templateName, Template ID: $($template.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Administrative Template" -Items @($template) -AssignmentReason "No Assignment"
                }
            }

            Write-Host "`n------- Compliance Policies -------" -ForegroundColor Cyan
            if ($unassignedPolicies.CompliancePolicies.Count -eq 0) {
                Write-Host "No unassigned Compliance Policies found" -ForegroundColor Gray
            }
            else {
                foreach ($policy in $unassignedPolicies.CompliancePolicies) {
                    $policyName = if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                    Write-Host "Compliance Policy Name: $policyName, Policy ID: $($policy.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Compliance Policy" -Items @($policy) -AssignmentReason "No Assignment"
                }
            }

            Write-Host "`n------- App Protection Policies -------" -ForegroundColor Cyan
            if ($unassignedPolicies.AppProtectionPolicies.Count -eq 0) {
                Write-Host "No unassigned App Protection Policies found" -ForegroundColor Gray
            }
            else {
                foreach ($policy in $unassignedPolicies.AppProtectionPolicies) {
                    $policyName = $policy.displayName
                    $policyType = switch ($policy.'@odata.type') {
                        "#microsoft.graph.androidManagedAppProtection" { "Android" }
                        "#microsoft.graph.iosManagedAppProtection" { "iOS" }
                        "#microsoft.graph.windowsManagedAppProtection" { "Windows" }
                        default { "Unknown" }
                    }
                    Write-Host "App Protection Policy Name: $policyName, Policy ID: $($policy.id), Type: $policyType" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "App Protection Policy" -Items @($policy) -AssignmentReason "No Assignment"
                }
            }

            Write-Host "`n------- App Configuration Policies -------" -ForegroundColor Cyan
            if ($unassignedPolicies.AppConfigurationPolicies.Count -eq 0) {
                Write-Host "No unassigned App Configuration Policies found" -ForegroundColor Gray
            }
            else {
                foreach ($policy in $unassignedPolicies.AppConfigurationPolicies) {
                    $policyName = if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                    Write-Host "App Configuration Policy Name: $policyName, Policy ID: $($policy.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "App Configuration Policy" -Items @($policy) -AssignmentReason "No Assignment"
                }
            }

            Write-Host "`n------- Platform Scripts -------" -ForegroundColor Cyan
            if ($unassignedPolicies.PlatformScripts.Count -eq 0) {
                Write-Host "No unassigned Platform Scripts found" -ForegroundColor Gray
            }
            else {
                foreach ($script in $unassignedPolicies.PlatformScripts) {
                    $scriptName = if ([string]::IsNullOrWhiteSpace($script.name)) { $script.displayName } else { $script.name }
                    Write-Host "Script Name: $scriptName, Script ID: $($script.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Platform Scripts" -Items @($script) -AssignmentReason "No Assignment"
                }
            }

            Write-Host "`n------- Proactive Remediation Scripts -------" -ForegroundColor Cyan
            if ($unassignedPolicies.HealthScripts.Count -eq 0) {
                Write-Host "No unassigned Proactive Remediation Scripts found" -ForegroundColor Gray
            }
            else {
                foreach ($script in $unassignedPolicies.HealthScripts) {
                    $scriptName = if ([string]::IsNullOrWhiteSpace($script.name)) { $script.displayName } else { $script.name }
                    Write-Host "Script Name: $scriptName, Script ID: $($script.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Proactive Remediation Scripts" -Items @($script) -AssignmentReason "No Assignment"
                }
            }

            Write-Host "`n------- Endpoint Security - Antivirus Profiles -------" -ForegroundColor Cyan
            if ($unassignedPolicies.AntivirusProfiles.Count -eq 0) {
                Write-Host "No unassigned Antivirus Profiles found" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $unassignedPolicies.AntivirusProfiles) {
                    Write-Host "Antivirus Profile Name: $($profile.displayName), Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Antivirus" -Items @($profile) -AssignmentReason "No Assignment"
                }
            }

            Write-Host "`n------- Endpoint Security - Disk Encryption Profiles -------" -ForegroundColor Cyan
            if ($unassignedPolicies.DiskEncryptionProfiles.Count -eq 0) {
                Write-Host "No unassigned Disk Encryption Profiles found" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $unassignedPolicies.DiskEncryptionProfiles) {
                    Write-Host "Disk Encryption Profile Name: $($profile.displayName), Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Disk Encryption" -Items @($profile) -AssignmentReason "No Assignment"
                }
            }

            Write-Host "`n------- Endpoint Security - Firewall Profiles -------" -ForegroundColor Cyan
            if ($unassignedPolicies.FirewallProfiles.Count -eq 0) {
                Write-Host "No unassigned Firewall Profiles found" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $unassignedPolicies.FirewallProfiles) {
                    Write-Host "Firewall Profile Name: $($profile.displayName), Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Firewall" -Items @($profile) -AssignmentReason "No Assignment"
                }
            }

            Write-Host "`n------- Endpoint Security - EDR Profiles -------" -ForegroundColor Cyan
            if ($unassignedPolicies.EndpointDetectionProfiles.Count -eq 0) {
                Write-Host "No unassigned EDR Profiles found" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $unassignedPolicies.EndpointDetectionProfiles) {
                    Write-Host "EDR Profile Name: $($profile.displayName), Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - EDR" -Items @($profile) -AssignmentReason "No Assignment"
                }
            }

            Write-Host "`n------- Endpoint Security - ASR Profiles -------" -ForegroundColor Cyan
            if ($unassignedPolicies.AttackSurfaceProfiles.Count -eq 0) {
                Write-Host "No unassigned ASR Profiles found" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $unassignedPolicies.AttackSurfaceProfiles) {
                    Write-Host "ASR Profile Name: $($profile.displayName), Profile ID: $($profile.id)" -ForegroundColor White
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - ASR" -Items @($profile) -AssignmentReason "No Assignment"
                }
            }

            Export-ResultsIfRequested -ExportData $exportData -DefaultFileName "IntuneUnassignedPolicies.csv" -ForceExport:$ExportToCSV -CustomExportPath $ExportPath
        }
       
        '9' {
            Write-Host "Checking for policies assigned to empty groups..." -ForegroundColor Green
            $exportData = [System.Collections.ArrayList]::new()

            function Test-EmptyGroup {
                param (
                    [Parameter(Mandatory = $true)]
                    [string]$GroupId
                )

                try {
                    $membersUri = "$GraphEndpoint/v1.0/groups/$GroupId/members?`$select=id"
                    $response = Invoke-MgGraphRequest -Uri $membersUri -Method Get
                    return $response.value.Count -eq 0
                }
                catch {
                    Write-Host "Error checking members for group $GroupId : $($_.Exception.Message)" -ForegroundColor Red
                    return $false
                }
            }

            $emptyGroupAssignments = @{
                DeviceConfigs            = @()
                SettingsCatalog          = @()
                AdminTemplates           = @()
                CompliancePolicies       = @()
                AppProtectionPolicies    = @()
                AppConfigurationPolicies = @()
                PlatformScripts          = @()
                HealthScripts            = @()
            }

            Write-Host "Fetching Device Configurations..." -ForegroundColor Yellow
            $deviceConfigs = Get-IntuneEntities -EntityType "deviceConfigurations"
            foreach ($config in $deviceConfigs) {
                $assignments = Get-IntuneAssignments -EntityType "deviceConfigurations" -EntityId $config.id
                foreach ($assignment in $assignments) {
                    if ($assignment.Reason -eq "Group Assignment" -and $assignment.GroupId) {
                        $groupInfo = Get-GroupInfo -GroupId $assignment.GroupId
                        if ($groupInfo.Success -and (Test-EmptyGroup -GroupId $assignment.GroupId)) {
                            $config | Add-Member -NotePropertyName 'EmptyGroupInfo' -NotePropertyValue "Assigned to empty group: $($groupInfo.DisplayName)" -Force
                            $emptyGroupAssignments.DeviceConfigs += $config
                            break
                        }
                    }
                }
            }

            Write-Host "Fetching Settings Catalog Policies..." -ForegroundColor Yellow
            $settingsCatalog = Get-IntuneEntities -EntityType "configurationPolicies"
            foreach ($policy in $settingsCatalog) {
                $assignments = Get-IntuneAssignments -EntityType "configurationPolicies" -EntityId $policy.id
                foreach ($assignment in $assignments) {
                    if ($assignment.Reason -eq "Group Assignment" -and $assignment.GroupId) {
                        $groupInfo = Get-GroupInfo -GroupId $assignment.GroupId
                        if ($groupInfo.Success -and (Test-EmptyGroup -GroupId $assignment.GroupId)) {
                            $policy | Add-Member -NotePropertyName 'EmptyGroupInfo' -NotePropertyValue "Assigned to empty group: $($groupInfo.DisplayName)" -Force
                            $emptyGroupAssignments.SettingsCatalog += $policy
                            break
                        }
                    }
                }
            }

            Write-Host "Fetching Administrative Templates..." -ForegroundColor Yellow
            $adminTemplates = Get-IntuneEntities -EntityType "groupPolicyConfigurations"
            foreach ($template in $adminTemplates) {
                $assignments = Get-IntuneAssignments -EntityType "groupPolicyConfigurations" -EntityId $template.id
                foreach ($assignment in $assignments) {
                    if ($assignment.Reason -eq "Group Assignment" -and $assignment.GroupId) {
                        $groupInfo = Get-GroupInfo -GroupId $assignment.GroupId
                        if ($groupInfo.Success -and (Test-EmptyGroup -GroupId $assignment.GroupId)) {
                            $template | Add-Member -NotePropertyName 'EmptyGroupInfo' -NotePropertyValue "Assigned to empty group: $($groupInfo.DisplayName)" -Force
                            $emptyGroupAssignments.AdminTemplates += $template
                            break
                        }
                    }
                }
            }

            Write-Host "Fetching Compliance Policies..." -ForegroundColor Yellow
            $compliancePolicies = Get-IntuneEntities -EntityType "deviceCompliancePolicies"
            foreach ($policy in $compliancePolicies) {
                $assignments = Get-IntuneAssignments -EntityType "deviceCompliancePolicies" -EntityId $policy.id
                foreach ($assignment in $assignments) {
                    if ($assignment.Reason -eq "Group Assignment" -and $assignment.GroupId) {
                        $groupInfo = Get-GroupInfo -GroupId $assignment.GroupId
                        if ($groupInfo.Success -and (Test-EmptyGroup -GroupId $assignment.GroupId)) {
                            $policy | Add-Member -NotePropertyName 'EmptyGroupInfo' -NotePropertyValue "Assigned to empty group: $($groupInfo.DisplayName)" -Force
                            $emptyGroupAssignments.CompliancePolicies += $policy
                            break
                        }
                    }
                }
            }

            Write-Host "Fetching App Protection Policies..." -ForegroundColor Yellow
            $appProtectionPolicies = Get-IntuneEntities -EntityType "deviceAppManagement/managedAppPolicies"
            foreach ($policy in $appProtectionPolicies) {
                $policyType = $policy.'@odata.type'
                $assignmentsUri = switch ($policyType) {
                    "#microsoft.graph.androidManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/androidManagedAppProtections('$($policy.id)')/assignments" }
                    "#microsoft.graph.iosManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/iosManagedAppProtections('$($policy.id)')/assignments" }
                    "#microsoft.graph.windowsManagedAppProtection" { "$GraphEndpoint/beta/deviceAppManagement/windowsManagedAppProtections('$($policy.id)')/assignments" }
                    default { $null }
                }

                if ($assignmentsUri) {
                    try {
                        $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get
                        $assignments = @()
                        foreach ($assignment in $assignmentResponse.value) {
                            $assignmentReason = $null
                            switch ($assignment.target.'@odata.type') {
                                '#microsoft.graph.allLicensedUsersAssignmentTarget' { 
                                    $assignmentReason = "All Users"
                                }
                                '#microsoft.graph.groupAssignmentTarget' {
                                    if (!$GroupId -or $assignment.target.groupId -eq $GroupId) {
                                        $assignmentReason = "Group Assignment"
                                    }
                                }
                            }

                            if ($assignmentReason) {
                                $assignments += @{
                                    Reason  = $assignmentReason
                                    GroupId = $assignment.target.groupId
                                }
                            }
                        }

                        if ($assignments.Count -gt 0) {
                            $assignmentSummary = $assignments | ForEach-Object {
                                if ($_.Reason -eq "Group Assignment") {
                                    $groupInfo = Get-GroupInfo -GroupId $_.GroupId
                                    "$($_.Reason) - $($groupInfo.DisplayName)"
                                }
                                else {
                                    $_.Reason
                                }
                            }
                            $policy | Add-Member -NotePropertyName 'AssignmentSummary' -NotePropertyValue ($assignmentSummary -join "; ") -Force
                            $emptyGroupAssignments.AppProtectionPolicies += $policy
                        }
                    }
                    catch {
                        Write-Host "Error fetching assignments for policy $($policy.displayName): $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }

            Write-Host "Fetching App Configuration Policies..." -ForegroundColor Yellow
            $appConfigPolicies = Get-IntuneEntities -EntityType "deviceAppManagement/mobileAppConfigurations"
            foreach ($policy in $appConfigPolicies) {
                $assignments = Get-IntuneAssignments -EntityType "mobileAppConfigurations" -EntityId $policy.id
                foreach ($assignment in $assignments) {
                    if ($assignment.Reason -eq "Group Assignment" -and $assignment.GroupId) {
                        $groupInfo = Get-GroupInfo -GroupId $assignment.GroupId
                        if ($groupInfo.Success -and (Test-EmptyGroup -GroupId $assignment.GroupId)) {
                            $policy | Add-Member -NotePropertyName 'EmptyGroupInfo' -NotePropertyValue "Assigned to empty group: $($groupInfo.DisplayName)" -Force
                            $emptyGroupAssignments.AppConfigurationPolicies += $policy
                            break
                        }
                    }
                }
            }

            Write-Host "Fetching Platform Scripts..." -ForegroundColor Yellow
            $platformScripts = Get-IntuneEntities -EntityType "deviceManagementScripts"
            foreach ($script in $platformScripts) {
                $assignments = Get-IntuneAssignments -EntityType "deviceManagementScripts" -EntityId $script.id
                foreach ($assignment in $assignments) {
                    if ($assignment.Reason -eq "Group Assignment" -and $assignment.GroupId) {
                        $groupInfo = Get-GroupInfo -GroupId $assignment.GroupId
                        if ($groupInfo.Success -and (Test-EmptyGroup -GroupId $assignment.GroupId)) {
                            $script | Add-Member -NotePropertyName 'EmptyGroupInfo' -NotePropertyValue "Assigned to empty group: $($groupInfo.DisplayName)" -Force
                            $emptyGroupAssignments.PlatformScripts += $script
                            break
                        }
                    }
                }
            }

            Write-Host "Fetching Proactive Remediation Scripts..." -ForegroundColor Yellow
            $healthScripts = Get-IntuneEntities -EntityType "deviceHealthScripts"
            foreach ($script in $healthScripts) {
                $assignments = Get-IntuneAssignments -EntityType "deviceHealthScripts" -EntityId $script.id
                foreach ($assignment in $assignments) {
                    if ($assignment.Reason -eq "Group Assignment" -and $assignment.GroupId) {
                        $groupInfo = Get-GroupInfo -GroupId $assignment.GroupId
                        if ($groupInfo.Success -and (Test-EmptyGroup -GroupId $assignment.GroupId)) {
                            $script | Add-Member -NotePropertyName 'EmptyGroupInfo' -NotePropertyValue "Assigned to empty group: $($groupInfo.DisplayName)" -Force
                            $emptyGroupAssignments.HealthScripts += $script
                            break
                        }
                    }
                }
            }

            Write-Host "`nPolicies Assigned to Empty Groups:" -ForegroundColor Green

            Write-Host "`n------- Device Configurations -------" -ForegroundColor Cyan
            if ($emptyGroupAssignments.DeviceConfigs.Count -eq 0) {
                Write-Host "No Device Configurations assigned to empty groups" -ForegroundColor Gray
            }
            else {
                foreach ($config in $emptyGroupAssignments.DeviceConfigs) {
                    $configName = if ([string]::IsNullOrWhiteSpace($config.name)) { $config.displayName } else { $config.name }
                    Write-Host "Device Configuration Name: $configName" -ForegroundColor White
                    Write-Host "Configuration ID: $($config.id)" -ForegroundColor Gray
                    Write-Host "$($config.EmptyGroupInfo)" -ForegroundColor Yellow
                    Write-Host ""
                    Add-ExportData -ExportData $exportData -Category "Device Configuration" -Items @($config) -AssignmentReason $config.EmptyGroupInfo
                }
            }

            Write-Host "`n------- Settings Catalog Policies -------" -ForegroundColor Cyan
            if ($emptyGroupAssignments.SettingsCatalog.Count -eq 0) {
                Write-Host "No Settings Catalog Policies assigned to empty groups" -ForegroundColor Gray
            }
            else {
                foreach ($policy in $emptyGroupAssignments.SettingsCatalog) {
                    $policyName = if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                    Write-Host "Settings Catalog Policy Name: $policyName" -ForegroundColor White
                    Write-Host "Policy ID: $($policy.id)" -ForegroundColor Gray
                    Write-Host "$($policy.EmptyGroupInfo)" -ForegroundColor Yellow
                    Write-Host ""
                    Add-ExportData -ExportData $exportData -Category "Settings Catalog Policy" -Items @($policy) -AssignmentReason $policy.EmptyGroupInfo
                }
            }

            Write-Host "`n------- Administrative Templates -------" -ForegroundColor Cyan
            if ($emptyGroupAssignments.AdminTemplates.Count -eq 0) {
                Write-Host "No Administrative Templates assigned to empty groups" -ForegroundColor Gray
            }
            else {
                foreach ($template in $emptyGroupAssignments.AdminTemplates) {
                    $templateName = if ([string]::IsNullOrWhiteSpace($template.name)) { $template.displayName } else { $template.name }
                    Write-Host "Administrative Template Name: $templateName" -ForegroundColor White
                    Write-Host "Template ID: $($template.id)" -ForegroundColor Gray
                    Write-Host "$($template.EmptyGroupInfo)" -ForegroundColor Yellow
                    Write-Host ""
                    Add-ExportData -ExportData $exportData -Category "Administrative Template" -Items @($template) -AssignmentReason $template.EmptyGroupInfo
                }
            }

            Write-Host "`n------- Compliance Policies -------" -ForegroundColor Cyan
            if ($emptyGroupAssignments.CompliancePolicies.Count -eq 0) {
                Write-Host "No Compliance Policies assigned to empty groups" -ForegroundColor Gray
            }
            else {
                foreach ($policy in $emptyGroupAssignments.CompliancePolicies) {
                    $policyName = if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                    Write-Host "Compliance Policy Name: $policyName" -ForegroundColor White
                    Write-Host "Policy ID: $($policy.id)" -ForegroundColor Gray
                    Write-Host "$($policy.EmptyGroupInfo)" -ForegroundColor Yellow
                    Write-Host ""
                    Add-ExportData -ExportData $exportData -Category "Compliance Policy" -Items @($policy) -AssignmentReason $policy.EmptyGroupInfo
                }
            }

            Write-Host "`n------- App Protection Policies -------" -ForegroundColor Cyan
            if ($emptyGroupAssignments.AppProtectionPolicies.Count -eq 0) {
                Write-Host "No App Protection Policies assigned to empty groups" -ForegroundColor Gray
            }
            else {
                foreach ($policy in $emptyGroupAssignments.AppProtectionPolicies) {
                    $policyName = $policy.displayName
                    $policyType = switch ($policy.'@odata.type') {
                        "#microsoft.graph.androidManagedAppProtection" { "Android" }
                        "#microsoft.graph.iosManagedAppProtection" { "iOS" }
                        "#microsoft.graph.windowsManagedAppProtection" { "Windows" }
                        default { "Unknown" }
                    }
                    Write-Host "App Protection Policy Name: $policyName" -ForegroundColor White
                    Write-Host "Policy ID: $($policy.id), Type: $policyType" -ForegroundColor Gray
                    Write-Host "$($policy.EmptyGroupInfo)" -ForegroundColor Yellow
                    Write-Host ""
                    Add-ExportData -ExportData $exportData -Category "App Protection Policy" -Items @($policy) -AssignmentReason $policy.EmptyGroupInfo
                }
            }

            Write-Host "`n------- App Configuration Policies -------" -ForegroundColor Cyan
            if ($emptyGroupAssignments.AppConfigurationPolicies.Count -eq 0) {
                Write-Host "No App Configuration Policies assigned to empty groups" -ForegroundColor Gray
            }
            else {
                foreach ($policy in $emptyGroupAssignments.AppConfigurationPolicies) {
                    $policyName = if ([string]::IsNullOrWhiteSpace($policy.name)) { $policy.displayName } else { $policy.name }
                    Write-Host "App Configuration Policy Name: $policyName" -ForegroundColor White
                    Write-Host "Policy ID: $($policy.id)" -ForegroundColor Gray
                    Write-Host "$($policy.EmptyGroupInfo)" -ForegroundColor Yellow
                    Write-Host ""
                    Add-ExportData -ExportData $exportData -Category "App Configuration Policy" -Items @($policy) -AssignmentReason $policy.EmptyGroupInfo
                }
            }

            Write-Host "`n------- Platform Scripts -------" -ForegroundColor Cyan
            if ($emptyGroupAssignments.PlatformScripts.Count -eq 0) {
                Write-Host "No Platform Scripts assigned to empty groups" -ForegroundColor Gray
            }
            else {
                foreach ($script in $emptyGroupAssignments.PlatformScripts) {
                    $scriptName = if ([string]::IsNullOrWhiteSpace($script.name)) { $script.displayName } else { $script.name }
                    Write-Host "Script Name: $scriptName" -ForegroundColor White
                    Write-Host "Script ID: $($script.id)" -ForegroundColor Gray
                    Write-Host "$($script.EmptyGroupInfo)" -ForegroundColor Yellow
                    Write-Host ""
                    Add-ExportData -ExportData $exportData -Category "Platform Scripts" -Items @($script) -AssignmentReason $script.EmptyGroupInfo
                }
            }

            Write-Host "`n------- Proactive Remediation Scripts -------" -ForegroundColor Cyan
            if ($emptyGroupAssignments.HealthScripts.Count -eq 0) {
                Write-Host "No Proactive Remediation Scripts assigned to empty groups" -ForegroundColor Gray
            }
            else {
                foreach ($script in $emptyGroupAssignments.HealthScripts) {
                    $scriptName = if ([string]::IsNullOrWhiteSpace($script.name)) { $script.displayName } else { $script.name }
                    Write-Host "Script Name: $scriptName" -ForegroundColor White
                    Write-Host "Script ID: $($script.id)" -ForegroundColor Gray
                    Write-Host "$($script.EmptyGroupInfo)" -ForegroundColor Yellow
                    Write-Host ""
                    Add-ExportData -ExportData $exportData -Category "Proactive Remediation Scripts" -Items @($script) -AssignmentReason $script.EmptyGroupInfo
                }
            }

            Write-Host "`n------- Endpoint Security - Antivirus Profiles -------" -ForegroundColor Cyan
            if ($emptyGroupAssignments.AntivirusProfiles.Count -eq 0) {
                Write-Host "No Antivirus Profiles assigned to empty groups" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $emptyGroupAssignments.AntivirusProfiles) {
                    Write-Host "Antivirus Profile Name: $($profile.displayName)" -ForegroundColor White
                    Write-Host "Profile ID: $($profile.id)" -ForegroundColor Gray
                    Write-Host "$($profile.EmptyGroupInfo)" -ForegroundColor Yellow
                    Write-Host ""
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Antivirus" -Items @($profile) -AssignmentReason $profile.EmptyGroupInfo
                }
            }

            Write-Host "`n------- Endpoint Security - Disk Encryption Profiles -------" -ForegroundColor Cyan
            if ($emptyGroupAssignments.DiskEncryptionProfiles.Count -eq 0) {
                Write-Host "No Disk Encryption Profiles assigned to empty groups" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $emptyGroupAssignments.DiskEncryptionProfiles) {
                    Write-Host "Disk Encryption Profile Name: $($profile.displayName)" -ForegroundColor White
                    Write-Host "Profile ID: $($profile.id)" -ForegroundColor Gray
                    Write-Host "$($profile.EmptyGroupInfo)" -ForegroundColor Yellow
                    Write-Host ""
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Disk Encryption" -Items @($profile) -AssignmentReason $profile.EmptyGroupInfo
                }
            }

            Write-Host "`n------- Endpoint Security - Firewall Profiles -------" -ForegroundColor Cyan
            if ($emptyGroupAssignments.FirewallProfiles.Count -eq 0) {
                Write-Host "No Firewall Profiles assigned to empty groups" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $emptyGroupAssignments.FirewallProfiles) {
                    Write-Host "Firewall Profile Name: $($profile.displayName)" -ForegroundColor White
                    Write-Host "Profile ID: $($profile.id)" -ForegroundColor Gray
                    Write-Host "$($profile.EmptyGroupInfo)" -ForegroundColor Yellow
                    Write-Host ""
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - Firewall" -Items @($profile) -AssignmentReason $profile.EmptyGroupInfo
                }
            }

            Write-Host "`n------- Endpoint Security - EDR Profiles -------" -ForegroundColor Cyan
            if ($emptyGroupAssignments.EndpointDetectionProfiles.Count -eq 0) {
                Write-Host "No EDR Profiles assigned to empty groups" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $emptyGroupAssignments.EndpointDetectionProfiles) {
                    Write-Host "EDR Profile Name: $($profile.displayName)" -ForegroundColor White
                    Write-Host "Profile ID: $($profile.id)" -ForegroundColor Gray
                    Write-Host "$($profile.EmptyGroupInfo)" -ForegroundColor Yellow
                    Write-Host ""
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - EDR" -Items @($profile) -AssignmentReason $profile.EmptyGroupInfo
                }
            }

            Write-Host "`n------- Endpoint Security - ASR Profiles -------" -ForegroundColor Cyan
            if ($emptyGroupAssignments.AttackSurfaceProfiles.Count -eq 0) {
                Write-Host "No ASR Profiles assigned to empty groups" -ForegroundColor Gray
            }
            else {
                foreach ($profile in $emptyGroupAssignments.AttackSurfaceProfiles) {
                    Write-Host "ASR Profile Name: $($profile.displayName)" -ForegroundColor White
                    Write-Host "Profile ID: $($profile.id)" -ForegroundColor Gray
                    Write-Host "$($profile.EmptyGroupInfo)" -ForegroundColor Yellow
                    Write-Host ""
                    Add-ExportData -ExportData $exportData -Category "Endpoint Security - ASR" -Items @($profile) -AssignmentReason $profile.EmptyGroupInfo
                }
            }

            Export-ResultsIfRequested -ExportData $exportData -DefaultFileName "IntuneEmptyGroupAssignments.csv" -ForceExport:$ExportToCSV -CustomExportPath $ExportPath
        }
        '10' {
            Write-Host "Compare Group Assignments chosen" -ForegroundColor Green

            if ($parameterMode -and $CompareGroupNames) {
                $groupInput = $CompareGroupNames
            }
            else {
                Write-Host "Please enter Group names or Object IDs to compare, separated by commas (,): " -ForegroundColor Cyan
                Write-Host "Example: 'Marketing Team, 12345678-1234-1234-1234-123456789012'" -ForegroundColor Gray
                $groupInput = Read-Host
            }
            
            $groupInputs = $groupInput -split ',' | ForEach-Object { $_.Trim() }

            if ($groupInputs.Count -lt 2) {
                Write-Host "Please provide at least two groups to compare." -ForegroundColor Red
                if ($parameterMode) { exit 1 } else { continue }
            }

            $groupAssignments = @{}

            $resolvedGroups = @{}
            foreach ($input in $groupInputs) {
                Write-Host "`nProcessing input: $input" -ForegroundColor Yellow

                $groupId = $null
                $groupName = $null

                if ($input -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
                    try {
                        $groupUri = "$GraphEndpoint/v1.0/groups/$input"
                        $groupResponse = Invoke-MgGraphRequest -Uri $groupUri -Method Get
                        $groupId = $groupResponse.id
                        $groupName = $groupResponse.displayName
                        $resolvedGroups[$groupId] = $groupName
                
                        $groupAssignments[$groupName] = @{
                            DeviceConfigs      = [System.Collections.ArrayList]::new()
                            SettingsCatalog    = [System.Collections.ArrayList]::new()
                            AdminTemplates     = [System.Collections.ArrayList]::new()
                            CompliancePolicies = [System.Collections.ArrayList]::new()
                            RequiredApps       = [System.Collections.ArrayList]::new()
                            AvailableApps      = [System.Collections.ArrayList]::new()
                            AppsUninstall      = [System.Collections.ArrayList]::new()
                            PlatformScripts    = [System.Collections.ArrayList]::new()
                            HealthScripts      = [System.Collections.ArrayList]::new()
                        }
                
                        Write-Host "Found group by ID: $groupName" -ForegroundColor Green
                    }
                    catch {
                        Write-Host "No group found with ID: $input" -ForegroundColor Red
                        continue
                    }
                }
                else {
                    $groupUri = "$GraphEndpoint/v1.0/groups?`$filter=displayName eq '$input'"
                    $groupResponse = Invoke-MgGraphRequest -Uri $groupUri -Method Get

                    if ($groupResponse.value.Count -eq 0) {
                        Write-Host "No group found with name: $input" -ForegroundColor Red
                        continue
                    }
                    elseif ($groupResponse.value.Count -gt 1) {
                        Write-Host "Multiple groups found with name: $input. Please use the Object ID instead:" -ForegroundColor Red
                        foreach ($group in $groupResponse.value) {
                            Write-Host "  - $($group.displayName) (ID: $($group.id))" -ForegroundColor Yellow
                        }
                        continue
                    }

                    $groupId = $groupResponse.value[0].id
                    $groupName = $groupResponse.value[0].displayName
                    $resolvedGroups[$groupId] = $groupName
            
                    $groupAssignments[$groupName] = @{
                        DeviceConfigs      = [System.Collections.ArrayList]::new()
                        SettingsCatalog    = [System.Collections.ArrayList]::new()
                        AdminTemplates     = [System.Collections.ArrayList]::new()
                        CompliancePolicies = [System.Collections.ArrayList]::new()
                        RequiredApps       = [System.Collections.ArrayList]::new()
                        AvailableApps      = [System.Collections.ArrayList]::new()
                        AppsUninstall      = [System.Collections.ArrayList]::new()
                        PlatformScripts    = [System.Collections.ArrayList]::new()
                        HealthScripts      = [System.Collections.ArrayList]::new()
                    }
            
                    Write-Host "Found group by name: $groupName (ID: $groupId)" -ForegroundColor Green
                }

                $deviceConfigsUri = "$GraphEndpoint/beta/deviceManagement/deviceConfigurations"
                $deviceConfigsResponse = Invoke-MgGraphRequest -Uri $deviceConfigsUri -Method Get
                $allDeviceConfigs = $deviceConfigsResponse.value
                while ($deviceConfigsResponse.'@odata.nextLink') {
                    $deviceConfigsResponse = Invoke-MgGraphRequest -Uri $deviceConfigsResponse.'@odata.nextLink' -Method Get
                    $allDeviceConfigs += $deviceConfigsResponse.value
                }
                $totalDeviceConfigs = $allDeviceConfigs.Count
                $currentDeviceConfig = 0
                foreach ($config in $allDeviceConfigs) {
                    $currentDeviceConfig++
                    Write-Host "`rFetching Device Configuration $currentDeviceConfig of $totalDeviceConfigs" -NoNewline
                    $configId = $config.id
                    $assignmentsUri = "$GraphEndpoint/beta/deviceManagement/deviceConfigurations('$configId')/assignments"
                    $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get
            
                    if ($assignmentResponse.value | Where-Object { $_.target.groupId -eq $groupId }) {
                        [void]$groupAssignments[$groupName].DeviceConfigs.Add($config.displayName)
                    }
                }
                Write-Host "`rFetching Device Configuration $totalDeviceConfigs of $totalDeviceConfigs" -NoNewline
                Start-Sleep -Milliseconds 100
                Write-Host ""  # Move to the next line after the loop

                $settingsCatalogUri = "$GraphEndpoint/beta/deviceManagement/configurationPolicies"
                $settingsCatalogResponse = Invoke-MgGraphRequest -Uri $settingsCatalogUri -Method Get

                foreach ($policy in $settingsCatalogResponse.value) {
                    $policyId = $policy.id
                    $assignmentsUri = "$GraphEndpoint/beta/deviceManagement/configurationPolicies('$policyId')/assignments"
                    $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get

                    if ($assignmentResponse.value | Where-Object { $_.target.groupId -eq $groupId }) {
                        [void]$groupAssignments[$groupName].SettingsCatalog.Add($policy.name)
                    }
                }

                $adminTemplatesUri = "$GraphEndpoint/beta/deviceManagement/groupPolicyConfigurations"
                $adminTemplatesResponse = Invoke-MgGraphRequest -Uri $adminTemplatesUri -Method Get

                foreach ($template in $adminTemplatesResponse.value) {
                    $templateId = $template.id
                    $assignmentsUri = "$GraphEndpoint/beta/deviceManagement/groupPolicyConfigurations('$templateId')/assignments"
                    $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get

                    if ($assignmentResponse.value | Where-Object { $_.target.groupId -eq $groupId }) {
                        [void]$groupAssignments[$groupName].AdminTemplates.Add($template.displayName)
                    }
                }

                $complianceUri = "$GraphEndpoint/beta/deviceManagement/deviceCompliancePolicies"
                $complianceResponse = Invoke-MgGraphRequest -Uri $complianceUri -Method Get

                foreach ($policy in $complianceResponse.value) {
                    $policyId = $policy.id
                    $assignmentsUri = "$GraphEndpoint/beta/deviceManagement/deviceCompliancePolicies('$policyId')/assignments"
                    $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get

                    if ($assignmentResponse.value | Where-Object { $_.target.groupId -eq $groupId }) {
                        [void]$groupAssignments[$groupName].CompliancePolicies.Add($policy.displayName)
                    }
                }

                $appUri = "$GraphEndpoint/beta/deviceAppManagement/mobileApps?`$filter=isAssigned eq true"
                $appResponse = Invoke-MgGraphRequest -Uri $appUri -Method Get

                foreach ($app in $appResponse.value) {
                    if ($app.isFeatured -or $app.isBuiltIn) {
                        continue
                    }

                    $appId = $app.id
                    $assignmentsUri = "$GraphEndpoint/beta/deviceAppManagement/mobileApps('$appId')/assignments"
                    $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get

                    foreach ($assignment in $assignmentResponse.value) {
                        if ($assignment.target.groupId -eq $groupId) {
                            switch ($assignment.intent) {
                                "required" { [void]$groupAssignments[$groupName].RequiredApps.Add($app.displayName) }
                                "available" { [void]$groupAssignments[$groupName].AvailableApps.Add($app.displayName) }
                                "uninstall" { [void]$groupAssignments[$groupName].UninstallApps.Add($app.displayName) }
                            }
                        }
                    }
                }

                $scriptsUri = "$GraphEndpoint/beta/deviceManagement/deviceManagementScripts"
                $scriptsResponse = Invoke-MgGraphRequest -Uri $scriptsUri -Method Get
                foreach ($script in $scriptsResponse.value) {
                    $scriptId = $script.id
                    $assignmentsUri = "$GraphEndpoint/beta/deviceManagement/deviceManagementScripts('$scriptId')/assignments"
                    $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get

                    if ($assignmentResponse.value | Where-Object { $_.target.groupId -eq $groupId }) {
                        $scriptInfo = "$($script.displayName) (PowerShell)"
                        [void]$groupAssignments[$groupName].PlatformScripts.Add($scriptInfo)
                    }
                }

                $shellScriptsUri = "$GraphEndpoint/beta/deviceManagement/deviceShellScripts"
                $shellScriptsResponse = Invoke-MgGraphRequest -Uri $shellScriptsUri -Method Get
                foreach ($script in $shellScriptsResponse.value) {
                    $scriptId = $script.id
                    $assignmentsUri = "$GraphEndpoint/beta/deviceManagement/deviceShellScripts('$scriptId')/groupAssignments"
                    $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get

                    if ($assignmentResponse.value | Where-Object { $_.targetGroupId -eq $groupId }) {
                        $scriptInfo = "$($script.displayName) (Shell)"
                        [void]$groupAssignments[$groupName].PlatformScripts.Add($scriptInfo)
                    }
                }

                $healthScriptsUri = "$GraphEndpoint/beta/deviceManagement/deviceHealthScripts"
                $healthScriptsResponse = Invoke-MgGraphRequest -Uri $healthScriptsUri -Method Get
                foreach ($script in $healthScriptsResponse.value) {
                    $scriptId = $script.id
                    $assignmentsUri = "$GraphEndpoint/beta/deviceManagement/deviceHealthScripts('$scriptId')/assignments"
                    $assignmentResponse = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get

                    if ($assignmentResponse.value | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget' -and $_.target.groupId -eq $groupId }) {
                        [void]$groupAssignments[$groupName].HealthScripts.Add($script.displayName)
                    }
                }

                $allIntentsForAntivirusCompare = Get-IntuneEntities -EntityType "deviceManagement/intents"
                $antivirusPolicies = $allIntentsForAntivirusCompare | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAntivirus' }
                if ($antivirusPolicies) {
                    foreach ($policy in $antivirusPolicies) {
                        $assignments = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        if ($assignments.value | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget' -and $_.target.groupId -eq $groupId }) {
                            [void]$groupAssignments[$groupName].AntivirusProfiles.Add($policy.displayName)
                        }
                    }
                }

                $allIntentsForDiskEncCompare = Get-IntuneEntities -EntityType "deviceManagement/intents"
                $diskEncryptionPolicies = $allIntentsForDiskEncCompare | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityDiskEncryption' }
                if ($diskEncryptionPolicies) {
                    foreach ($policy in $diskEncryptionPolicies) {
                        $assignments = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        if ($assignments.value | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget' -and $_.target.groupId -eq $groupId }) {
                            [void]$groupAssignments[$groupName].DiskEncryptionProfiles.Add($policy.displayName)
                        }
                    }
                }

                $allIntentsForFirewallCompare = Get-IntuneEntities -EntityType "deviceManagement/intents"
                $firewallPolicies = $allIntentsForFirewallCompare | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityFirewall' }
                if ($firewallPolicies) {
                    foreach ($policy in $firewallPolicies) {
                        $assignments = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        if ($assignments.value | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget' -and $_.target.groupId -eq $groupId }) {
                            [void]$groupAssignments[$groupName].FirewallProfiles.Add($policy.displayName)
                        }
                    }
                }

                $allIntentsForEDRCompare = Get-IntuneEntities -EntityType "deviceManagement/intents"
                $edrPolicies = $allIntentsForEDRCompare | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityEndpointDetectionAndResponse' }
                if ($edrPolicies) {
                    foreach ($policy in $edrPolicies) {
                        $assignments = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        if ($assignments.value | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget' -and $_.target.groupId -eq $groupId }) {
                            [void]$groupAssignments[$groupName].EndpointDetectionProfiles.Add($policy.displayName)
                        }
                    }
                }

                $allIntentsForASRCompare = Get-IntuneEntities -EntityType "deviceManagement/intents"
                $asrPolicies = $allIntentsForASRCompare | Where-Object { $_.templateReference -and $_.templateReference.templateFamily -eq 'endpointSecurityAttackSurfaceReductionRules' }
                if ($asrPolicies) {
                    foreach ($policy in $asrPolicies) {
                        $assignments = Invoke-MgGraphRequest -Uri "$GraphEndpoint/beta/deviceManagement/intents/$($policy.id)/assignments" -Method Get
                        if ($assignments.value | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget' -and $_.target.groupId -eq $groupId }) {
                            [void]$groupAssignments[$groupName].AttackSurfaceProfiles.Add($policy.displayName)
                        }
                    }
                }
            }

            Write-Host "`nComparison Results:" -ForegroundColor Cyan
            Write-Host "Comparing assignments between groups:" -ForegroundColor White
            foreach ($groupName in $groupAssignments.Keys) {
                Write-Host "  • $groupName" -ForegroundColor White
            }
            Write-Host ""

            $categories = @{
                "Settings Catalog"                    = "SettingsCatalog"
                "Administrative Templates"            = "AdminTemplates"
                "Compliance Policies"                 = "CompliancePolicies"
                "Available Apps"                      = "AvailableApps"
                "Required Apps"                       = "RequiredApps"
                "Platform Scripts"                    = "PlatformScripts"
                "Device Configurations"               = "DeviceConfigs"
                "Uninstall Apps"                      = "UninstallApps"
                "Proactive Remediation Scripts"       = "HealthScripts"
                "Endpoint Security - Antivirus"       = "AntivirusProfiles"
                "Endpoint Security - Disk Encryption" = "DiskEncryptionProfiles"
                "Endpoint Security - Firewall"        = "FirewallProfiles"
                "Endpoint Security - EDR"             = "EndpointDetectionProfiles"
                "Endpoint Security - ASR"             = "AttackSurfaceProfiles"
            }

            $uniquePolicies = [System.Collections.ArrayList]@()
            foreach ($groupName in $groupAssignments.Keys) {
                foreach ($categoryKey in $categories.Values) {
                    foreach ($policy in $groupAssignments[$groupName][$categoryKey]) {
                        if ($uniquePolicies -notcontains $policy) {
                            $null = $uniquePolicies.Add($policy)
                        }
                    }
                }
            }

            Write-Host "Found $($uniquePolicies.Count) unique policies/apps/scripts across all groups`n" -ForegroundColor Yellow

            foreach ($category in $categories.Keys) {
                $categoryKey = $categories[$category]

                Write-Host "=== $category ===" -ForegroundColor Cyan
                $foundAssignments = $false

                foreach ($policy in $uniquePolicies) {
                    $assignedGroups = @()
                    foreach ($groupName in $groupAssignments.Keys) {
                        if ($groupAssignments[$groupName][$categoryKey] -contains $policy) {
                            $assignedGroups += $groupName
                        }
                    }

                    if ($assignedGroups.Count -gt 0) {
                        $foundAssignments = $true
                        Write-Host "📋 Policy: " -NoNewline -ForegroundColor White
                        Write-Host "$policy" -ForegroundColor Yellow

                        if ($assignedGroups.Count -gt 1) {
                            Write-Host "  🔗 Shared Assignment!" -ForegroundColor Magenta
                        }

                        Write-Host "  ✅ Assigned to: " -NoNewline -ForegroundColor Green
                        Write-Host "$($assignedGroups -join ', ')" -ForegroundColor White

                        $notAssignedGroups = $groupAssignments.Keys | Where-Object { $assignedGroups -notcontains $_ }
                        if ($notAssignedGroups) {
                            Write-Host "  ❌ Not assigned to: " -NoNewline -ForegroundColor Red
                            Write-Host "$($notAssignedGroups -join ', ')" -ForegroundColor White
                        }
                        Write-Host ""
                    }
                }

                if (-not $foundAssignments) {
                    Write-Host "No assignments found in this category" -ForegroundColor Gray
                    Write-Host ""
                }
            }

            Write-Host "=== Summary ===" -ForegroundColor Cyan
            foreach ($groupName in $groupAssignments.Keys) {
                $totalAssignments = 0
                foreach ($categoryKey in $categories.Values) {
                    $totalAssignments += $groupAssignments[$groupName][$categoryKey].Count
                }
                Write-Host "$groupName has $totalAssignments total assignments" -ForegroundColor Yellow
            }
            Write-Host ""

            $comparisonResults = [System.Collections.ArrayList]@()
            foreach ($category in $categories.Keys) {
                $categoryKey = $categories[$category]
                foreach ($policy in $uniquePolicies) {
                    $assignedGroups = @()
                    foreach ($groupName in $groupAssignments.Keys) {
                        if ($groupAssignments[$groupName][$categoryKey] -contains $policy) {
                            $assignedGroups += $groupName
                        }
                    }

                    if ($assignedGroups.Count -gt 0) {
                        [void]$comparisonResults.Add([PSCustomObject]@{
                                Category           = $category
                                PolicyName         = $policy
                                AssignedTo         = $assignedGroups -join '; '
                                NotAssignedTo      = ($groupAssignments.Keys | Where-Object { $assignedGroups -notcontains $_ }) -join '; '
                                IsSharedAssignment = ($assignedGroups.Count -gt 1)
                            })
                    }
                }
            }
            
            if ($ExportToCSV -or -not $parameterMode) {
                $exportPath = if ($ExportPath) {
                    $ExportPath
                }
                elseif (-not $parameterMode) {
                    $export = Read-Host "Would you like to export the comparison results to CSV? (y/n)"
                    if ($export -eq 'y') {
                        Show-SaveFileDialog -DefaultFileName "IntuneGroupAssignmentComparison.csv"
                    }
                    else {
                        $null
                    }
                }
                else {
                    $null
                }
                
                if ($exportPath) {
                    $comparisonResults | Export-Csv -Path $exportPath -NoTypeInformation
                    Write-Host "Results exported to $exportPath" -ForegroundColor Green
                }
            }
        }

        '11' {
            Write-Host "Fetching all failed assignments..." -ForegroundColor Green
            $exportData = [System.Collections.ArrayList]::new()
            
            $failedAssignments = Get-AssignmentFailures
            
            if ($failedAssignments.Count -eq 0) {
                Write-Host "`nNo assignment failures found!" -ForegroundColor Green
            }
            else {
                Write-Host "`nFound $($failedAssignments.Count) assignment failures:" -ForegroundColor Yellow
                
                $groupedFailures = $failedAssignments | Group-Object -Property Type
                
                foreach ($group in $groupedFailures) {
                    Write-Host "`n=== $($group.Name) Failures ($($group.Count)) ===" -ForegroundColor Cyan
                    
                    foreach ($failure in $group.Group) {
                        Write-Host "`nPolicy: $($failure.PolicyName)" -ForegroundColor White
                        Write-Host "Device: $($failure.Target -replace 'Device: ', '')" -ForegroundColor Gray
                        Write-Host "Reason: $($failure.ErrorCode)" -ForegroundColor White
                        if ($failure.LastAttempt -and $failure.LastAttempt -ne "01/01/0001 00:00:00") {
                            Write-Host "Last Attempt: $($failure.LastAttempt)" -ForegroundColor Gray
                        }
                        
                        $null = $exportData.Add([PSCustomObject]@{
                                Type             = $failure.Type
                                PolicyName       = $failure.PolicyName
                                Target           = $failure.Target
                                ErrorCode        = $failure.ErrorCode
                                ErrorDescription = $failure.ErrorDescription
                                LastAttempt      = $failure.LastAttempt
                            })
                    }
                }
                
                Export-ResultsIfRequested -ExportData $exportData -ExportPath $ExportPath -ForceExport:$false
            }
        }

        '0' {
            Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Yellow
            Disconnect-MgGraph | Out-Null
            Write-Host ""
            exit
        }
        default {
            Write-Host "Invalid choice, please select 1-11, or 0." -ForegroundColor Red
        }
    }

    if ($selection -ne '0') {
        if ($parameterMode) {
            break
        }
        else {
            Write-Host "Press any key to return to the main menu..." -ForegroundColor Cyan
            $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
    }
} while ($selection -ne '0')
