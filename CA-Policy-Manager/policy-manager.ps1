#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.SignIns

<#
.SYNOPSIS
Conditional Access Policy Manager - Export and Import CA policies between tenants

.DESCRIPTION
This script provides a menu-driven interface to:
- Export Conditional Access policies from a source tenant
- Import Conditional Access policies to a destination tenant
- Clean and prepare policies for cross-tenant migration

.AUTHOR
IT Administrator

.VERSION
1.1
#>

# Function to show the main menu
function Show-Menu {
    Clear-Host
    Write-Host "====================================" -ForegroundColor Cyan
    Write-Host "  Conditional Access Policy Manager" -ForegroundColor Cyan
    Write-Host "====================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "1. Export CA Policies from Tenant" -ForegroundColor Green
    Write-Host "2. Import CA Policies to Tenant" -ForegroundColor Yellow
    Write-Host "3. Check Connection Status" -ForegroundColor Blue
    Write-Host "4. Disconnect from Microsoft Graph" -ForegroundColor Red
    Write-Host "5. Exit" -ForegroundColor Gray
    Write-Host ""
}

# Function to ensure we're connected with proper permissions
function Connect-ToGraph {
    param(
        [string]$Operation
    )
    
    $context = Get-MgContext
    if (-not $context) {
        Write-Host "Not connected to Microsoft Graph. Connecting..." -ForegroundColor Yellow
        Connect-MgGraph -Scopes "Policy.ReadWrite.ConditionalAccess", "Policy.Read.All", "Directory.Read.All"
    } else {
        Write-Host "Connected to tenant: $($context.TenantId)" -ForegroundColor Green
        Write-Host "Account: $($context.Account)" -ForegroundColor Green
    }
}

# Function to convert PSCustomObject to Hashtable recursively
function ConvertTo-Hashtable {
    param([Parameter(ValueFromPipeline)] $InputObject)
    
    if ($null -eq $InputObject) { return $null }
    
    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        $collection = @(
            foreach ($object in $InputObject) {
                ConvertTo-Hashtable $object
            }
        )
        return $collection
    }
    elseif ($InputObject -is [PSCustomObject]) {
        $hash = @{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $hash[$property.Name] = ConvertTo-Hashtable $property.Value
        }
        return $hash
    }
    else {
        return $InputObject
    }
}

# Function to remove null and empty values recursively
function Remove-EmptyProperties {
    param($InputObject)
    
    if ($null -eq $InputObject) { 
        return $null 
    }
    
    if ($InputObject -is [System.Collections.IDictionary]) {
        $result = @{}
        foreach ($key in $InputObject.Keys) {
            $value = Remove-EmptyProperties $InputObject[$key]
            if ($null -ne $value -and $value -ne @() -and $value -ne "") {
                $result[$key] = $value
            }
        }
        if ($result.Count -gt 0) { return $result } else { return $null }
    }
    elseif ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        $result = @()
        foreach ($item in $InputObject) {
            $cleanItem = Remove-EmptyProperties $item
            if ($null -ne $cleanItem) {
                $result += $cleanItem
            }
        }
        if ($result.Count -gt 0) { return $result } else { return $null }
    }
    else {
        return $InputObject
    }
}

# Function to clean policy for export
function Clean-PolicyForExport {
    param($Policy)
    
    # Create a minimal, valid policy structure
    $cleanedPolicy = @{
        displayName = $Policy.DisplayName
        state = "disabled"
        conditions = @{
            applications = @{
                includeApplications = if ($Policy.Conditions.Applications.IncludeApplications) { 
                    $Policy.Conditions.Applications.IncludeApplications 
                } else { @("All") }
                excludeApplications = if ($Policy.Conditions.Applications.ExcludeApplications) { 
                    $Policy.Conditions.Applications.ExcludeApplications 
                } else { @() }
            }
            users = @{
                includeUsers = @("None")
                excludeUsers = @()
                includeGroups = @()
                excludeGroups = @()
                includeRoles = if ($Policy.Conditions.Users.IncludeRoles) { 
                    $Policy.Conditions.Users.IncludeRoles 
                } else { @() }
                excludeRoles = if ($Policy.Conditions.Users.ExcludeRoles) { 
                    $Policy.Conditions.Users.ExcludeRoles 
                } else { @() }
            }
            clientAppTypes = if ($Policy.Conditions.ClientAppTypes) { 
                $Policy.Conditions.ClientAppTypes 
            } else { @("all") }
        }
        grantControls = @{
            operator = if ($Policy.GrantControls.Operator) { 
                $Policy.GrantControls.Operator 
            } else { "OR" }
            builtInControls = if ($Policy.GrantControls.BuiltInControls) { 
                $Policy.GrantControls.BuiltInControls 
            } else { @() }
        }
    }
    
    # Add optional conditions only if they exist
    if ($Policy.Conditions.Platforms -and ($Policy.Conditions.Platforms.IncludePlatforms -or $Policy.Conditions.Platforms.ExcludePlatforms)) {
        $cleanedPolicy.conditions.platforms = @{}
        if ($Policy.Conditions.Platforms.IncludePlatforms) {
            $cleanedPolicy.conditions.platforms.includePlatforms = $Policy.Conditions.Platforms.IncludePlatforms
        }
        if ($Policy.Conditions.Platforms.ExcludePlatforms) {
            $cleanedPolicy.conditions.platforms.excludePlatforms = $Policy.Conditions.Platforms.ExcludePlatforms
        }
    }
    
    if ($Policy.Conditions.Locations) {
        $hasValidLocations = $false
        $locationCondition = @{}
        
        if ($Policy.Conditions.Locations.IncludeLocations) {
            $standardIncludes = $Policy.Conditions.Locations.IncludeLocations | Where-Object { $_ -match '^(All|AllTrusted|None)$' }
            if ($standardIncludes) {
                $locationCondition.includeLocations = $standardIncludes
                $hasValidLocations = $true
            }
        }
        
        if ($Policy.Conditions.Locations.ExcludeLocations) {
            $standardExcludes = $Policy.Conditions.Locations.ExcludeLocations | Where-Object { $_ -match '^(All|AllTrusted|None)$' }
            if ($standardExcludes) {
                $locationCondition.excludeLocations = $standardExcludes
                $hasValidLocations = $true
            }
        }
        
        if ($hasValidLocations) {
            $cleanedPolicy.conditions.locations = $locationCondition
        }
    }
    
    if ($Policy.Conditions.SignInRiskLevels -and $Policy.Conditions.SignInRiskLevels.Count -gt 0) {
        $cleanedPolicy.conditions.signInRiskLevels = $Policy.Conditions.SignInRiskLevels
    }
    
    if ($Policy.Conditions.UserRiskLevels -and $Policy.Conditions.UserRiskLevels.Count -gt 0) {
        $cleanedPolicy.conditions.userRiskLevels = $Policy.Conditions.UserRiskLevels
    }
    
    # Add session controls if they exist
    if ($Policy.SessionControls -and $Policy.SessionControls.SignInFrequency -and $Policy.SessionControls.SignInFrequency.IsEnabled) {
        $cleanedPolicy.sessionControls = @{
            signInFrequency = $Policy.SessionControls.SignInFrequency
        }
    }
    
    # Add description if it exists
    if ($Policy.Description) {
        $cleanedPolicy.description = $Policy.Description
    }
    
    # Remove any null or empty properties
    $result = Remove-EmptyProperties $cleanedPolicy
    
    return $result
}

# Function to export CA policies
function Export-CAPolicies {
    Write-Host "Exporting Conditional Access Policies..." -ForegroundColor Cyan
    
    try {
        Connect-ToGraph -Operation "Export"
        
        Write-Host "Retrieving all Conditional Access policies..." -ForegroundColor Yellow
        $policies = Get-MgIdentityConditionalAccessPolicy -All
        
        if ($policies.Count -eq 0) {
            Write-Host "No Conditional Access policies found in this tenant." -ForegroundColor Yellow
            return
        }
        
        Write-Host "Found $($policies.Count) policies to export" -ForegroundColor Green
        
        # Clean policies for export
        $cleanedPolicies = @()
        foreach ($policy in $policies) {
            Write-Host "Processing: $($policy.DisplayName)" -ForegroundColor Gray
            $cleanedPolicy = Clean-PolicyForExport -Policy $policy
            $cleanedPolicies += $cleanedPolicy
        }
        
        # Get save path
        $exportPath = $null
        
        try {
            Add-Type -AssemblyName System.Windows.Forms
            $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
            $saveDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            $saveDialog.Title = "Save CA Policies Export"
            $saveDialog.FileName = "CA-Policies-Export-$(Get-Date -Format 'yyyyMMdd-HHmm').json"
            
            Write-Host "Opening save dialog..." -ForegroundColor Yellow
            
            if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $exportPath = $saveDialog.FileName
            }
        }
        catch {
            Write-Host "File dialog not available, using manual input method." -ForegroundColor Yellow
        }
        
        # Fallback to manual path input
        if (-not $exportPath) {
            Write-Host ""
            Write-Host "Please enter the full path where you want to save the export:" -ForegroundColor Yellow
            $defaultName = "CA-Policies-Export-$(Get-Date -Format 'yyyyMMdd-HHmm').json"
            Write-Host "Example: C:\Users\Username\Desktop\$defaultName" -ForegroundColor Gray
            $exportPath = Read-Host "Save path"
            
            if ($exportPath -and -not $exportPath.EndsWith('.json')) {
                $exportPath += '.json'
            }
        }
        
        if ($exportPath) {
            $jsonContent = $cleanedPolicies | ConvertTo-Json -Depth 10
            $jsonContent | Out-File -FilePath $exportPath -Encoding UTF8
            
            Write-Host ""
            Write-Host "Export completed successfully!" -ForegroundColor Green
            Write-Host "File saved to: $exportPath" -ForegroundColor Green
            Write-Host "Exported $($cleanedPolicies.Count) policies" -ForegroundColor Green
            Write-Host ""
            Write-Host "IMPORTANT NOTES:" -ForegroundColor Yellow
            Write-Host "- All policies exported as 'disabled' for safety" -ForegroundColor Yellow
            Write-Host "- User/group assignments cleared - will need reconfiguration" -ForegroundColor Yellow
            Write-Host "- Named locations with GUIDs removed - will need manual setup" -ForegroundColor Yellow
        }
        else {
            Write-Host "Export cancelled by user." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error during export: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Function to import CA policies
function Import-CAPolicies {
    Write-Host "Importing Conditional Access Policies..." -ForegroundColor Cyan
    
    try {
        Connect-ToGraph -Operation "Import"
        
        # Get import file path
        $importPath = $null
        
        try {
            Add-Type -AssemblyName System.Windows.Forms
            $openDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            $openDialog.Title = "Select CA Policies JSON File"
            
            Write-Host "Opening file selection dialog..." -ForegroundColor Yellow
            
            if ($openDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $importPath = $openDialog.FileName
            }
        }
        catch {
            Write-Host "File dialog not available, using manual input method." -ForegroundColor Yellow
        }
        
        # Fallback to manual path input
        if (-not $importPath) {
            Write-Host ""
            Write-Host "Please enter the full path to your JSON file:" -ForegroundColor Yellow
            Write-Host "Example: C:\Users\Username\Desktop\policies.json" -ForegroundColor Gray
            $importPath = Read-Host "File path"
            
            if (-not (Test-Path $importPath)) {
                Write-Host "File not found: $importPath" -ForegroundColor Red
                return
            }
        }
        
        if ($importPath) {
            Write-Host "Selected file: $importPath" -ForegroundColor Green
            
            # Load and validate JSON
            $jsonData = Get-Content $importPath | ConvertFrom-Json
            $policies = ConvertTo-Hashtable $jsonData
            
            Write-Host "Found $($policies.Count) policies to import" -ForegroundColor Green
            
            # Confirm import
            $confirmation = Read-Host "Do you want to proceed with importing $($policies.Count) policies? (y/N)"
            if ($confirmation -ne 'y' -and $confirmation -ne 'Y') {
                Write-Host "Import cancelled by user." -ForegroundColor Yellow
                return
            }
            
            # Import each policy
            $successCount = 0
            $failCount = 0
            
            foreach ($policy in $policies) {
                Write-Host "Creating policy: $($policy.displayName)" -ForegroundColor Gray
                
                try {
                    $result = New-MgIdentityConditionalAccessPolicy -BodyParameter $policy
                    Write-Host "SUCCESS: $($policy.displayName) - ID: $($result.Id)" -ForegroundColor Green
                    $successCount++
                }
                catch {
                    Write-Host "FAILED: $($policy.displayName)" -ForegroundColor Red
                    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Yellow
                    $failCount++
                }
                
                Start-Sleep -Seconds 1
            }
            
            Write-Host ""
            Write-Host "Import Summary:" -ForegroundColor Cyan
            Write-Host "Successfully created: $successCount policies" -ForegroundColor Green
            Write-Host "Failed to create: $failCount policies" -ForegroundColor Red
            
            if ($successCount -gt 0) {
                Write-Host ""
                Write-Host "NEXT STEPS:" -ForegroundColor Yellow
                Write-Host "1. Review imported policies in Azure AD portal" -ForegroundColor Yellow
                Write-Host "2. Configure user/group assignments" -ForegroundColor Yellow
                Write-Host "3. Set up named locations if needed" -ForegroundColor Yellow
                Write-Host "4. Test policies in 'Report-only' mode first" -ForegroundColor Yellow
                Write-Host "5. Enable policies gradually" -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "Import cancelled by user." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error during import: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Function to check connection status
function Check-Connection {
    $context = Get-MgContext
    if ($context) {
        Write-Host "Connected to Microsoft Graph" -ForegroundColor Green
        Write-Host "Tenant ID: $($context.TenantId)" -ForegroundColor Gray
        Write-Host "Account: $($context.Account)" -ForegroundColor Gray
        Write-Host "Scopes: $($context.Scopes -join ', ')" -ForegroundColor Gray
        
        try {
            $policyCount = (Get-MgIdentityConditionalAccessPolicy -Top 1 | Measure-Object).Count
            Write-Host "Permissions verified - can access CA policies" -ForegroundColor Green
        }
        catch {
            Write-Host "Permission issue - cannot access CA policies: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Not connected to Microsoft Graph" -ForegroundColor Red
    }
}

# Function to disconnect from Graph
function Disconnect-FromGraph {
    $context = Get-MgContext
    if ($context) {
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Yellow
        Disconnect-MgGraph
        Write-Host "Disconnected successfully." -ForegroundColor Green
    }
    else {
        Write-Host "Not currently connected to Microsoft Graph." -ForegroundColor Gray
    }
}

# Main script execution
do {
    Show-Menu
    $choice = Read-Host "Please select an option (1-5)"
    
    switch ($choice) {
        '1' {
            Export-CAPolicies
            Read-Host "Press Enter to continue..."
        }
        '2' {
            Import-CAPolicies
            Read-Host "Press Enter to continue..."
        }
        '3' {
            Check-Connection
            Read-Host "Press Enter to continue..."
        }
        '4' {
            Disconnect-FromGraph
            Read-Host "Press Enter to continue..."
        }
        '5' {
            Disconnect-FromGraph
            Write-Host "Exiting..." -ForegroundColor Gray
        }
        default {
            Write-Host "Invalid option. Please select 1-5." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
} while ($choice -ne '5')

Write-Host "Script completed." -ForegroundColor Green