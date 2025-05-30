# filepath: /FastCruise-PowerShell/FastCruise-PowerShell/src/Start-FastCruise.ps1

# This is the main entry point for the FastCruise application.
# It initializes the script and calls the necessary functions to perform the FastCruise operations.

# Import necessary modules and scripts
Import-Module ./Get-InstalledSoftware.ps1
Import-Module ./Get-WorkstationInfo.ps1
Import-Module ./Get-MacAddress.ps1
Import-Module ./Get-ComputerLocation.ps1
Import-Module ./Get-LastComputerStatus.ps1
Import-Module ./Get-FacilityIssues.ps1
Import-Module ./Start-ApplicationTest.ps1
Import-Module ./Show-VbForm.ps1
Import-Module ./Show-AsciiMenu.ps1

# Define the main function
function Start-FastCruise {
    param (
        [string]$FastCruiseReportPath = 'S:\FastCruise',
        [string]$FastCruiseFile = 'FastCruise.csv',
        [switch]$ManualInput = $false
    )

    # Check if the S: drive is mapped
    if (-not (Test-Path -Path 'S:\')) {
        Clear-Host
        Write-Warning -Message 'Yo. Mapping your S: Drive'
        Write-Host 'Net Use S: \\localhost\Folder-1' -ForegroundColor Cyan
        net.exe Use S: \\localhost\Folder-1 
    }

    # Initialize variables
    $jsonFilePath = "S:\ComputerLocation.json"
    $LocalCruiseFile = 'C:\temp\FastCruise\FastCruiseFile.csv'

    # Create local file if it doesn't exist
    if (-not (Test-Path -Path $LocalCruiseFile)) {
        Write-Verbose -Message 'Creating Local File.'
        $null = New-Item -Path $LocalCruiseFile -ItemType File -Force
    }

    # Call functions to gather information and perform operations
    $ComputerStat = @{}
    $ComputerStat['ComputerName'] = Get-WorkstationInfo -Info 'Name'
    $ComputerStat['SerialNumber'] = Get-WorkstationInfo -Info 'serialnumber'
    $ComputerStat['MacAddress'] = Get-MacAddress
    $ComputerStat['InstalledSoftware'] = Get-InstalledSoftware -SoftwareName @('Axway', 'Mozilla Firefox', 'McAfee Agent', 'Java')

    # Get computer location
    Get-ComputerLocation -jsonFilePath $jsonFilePath

    # Handle manual input if required
    if ($ManualInput) {
        $ComputerStat['ManualInput'] = $true
    }

    # Export the results to CSV
    $ComputerStat | Export-Csv -Path $FastCruiseReportPath\$FastCruiseFile -NoTypeInformation -Append -Force
}

# Call the main function
Start-FastCruise