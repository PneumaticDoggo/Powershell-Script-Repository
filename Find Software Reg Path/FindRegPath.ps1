$displayName = Read-Host -Prompt "Enter the DisplayName to search for"

$outputFile = "D:\RegistryPaths.csv"

$outputDir = [System.IO.Path]::GetDirectoryName($outputFile)
if (-not (Test-Path -Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory | Out-Null
}

function Search-Registry {
    param (
        [string]$displayName
    )

    $matches = @()

    function Search-Key {
        param (
            [string]$key
        )

        try {
            $subKeys = Get-ChildItem -Path $key -ErrorAction Stop
            foreach ($subKey in $subKeys) {
                Search-Key -key $subKey.PSPath
            }
        } catch {
        }

        $properties = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
        if ($properties) {
            foreach ($property in $properties.PSObject.Properties) {
                if ($property.Name -eq "DisplayName" -and $property.Value -like "*$displayName*") {
                    $matches += [PSCustomObject]@{
                        DisplayName = $displayName
                        RegistryPath = $key
                        PropertyName = $property.Name
                        PropertyValue = $property.Value
                    }
                }
            }
        }
    }

    $rootKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    foreach ($rootKey in $rootKeys) {
        Search-Key -key $rootKey
    }

    return $matches
}

$results = Search-Registry -displayName $displayName

if ($results) {
    $results | Export-Csv -Path $outputFile -NoTypeInformation
    Write-Host "Registry paths have been written to $outputFile"
} else {
    Write-Host "No matching registry paths found."
}
