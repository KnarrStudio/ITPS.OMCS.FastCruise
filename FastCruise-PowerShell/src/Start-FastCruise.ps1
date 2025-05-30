# filepath: /FastCruise-PowerShell/FastCruise-PowerShell/src/Start-FastCruise.ps1

# This is the main entry point for the FastCruise application.
# It initializes the script and calls the necessary functions to perform the FastCruise operations.

# Import necessary modules and scripts
Import-Module ./Get-InstalledSoftware.ps1 -Verbose
Import-Module ./Get-WorkstationInfo.ps1 -Verbose
Import-Module ./Get-MacAddress.ps1 -Verbose
Import-Module ./Get-ComputerLocation.ps1 -Verbose
Import-Module ./Get-LastComputerStatus.ps1 -Verbose
Import-Module ./Get-FacilityIssues.ps1 -Verbose
Import-Module ./Start-ApplicationTest.ps1 -Verbose
Import-Module ./Show-VbForm.ps1 -Verbose
Import-Module ./Show-AsciiMenu.ps1 -Verbose

# Define the main function
function Start-FastCruise {
    param (
        [string]$FastCruiseReportPath = 'S:\',
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
    $ComputerStat = [PSCustomObject]@{
    ComputerName      = Get-WorkstationInfo -Info 'Name'
    SerialNumber      = Get-WorkstationInfo -Info 'serialnumber'
    MacAddress        = Get-MacAddress
    InstalledSoftware = Get-InstalledSoftware -SoftwareName @('Axway', 'Mozilla Firefox', 'McAfee Agent', 'Java')
    ManualInput       = [bool]$ManualInput
    Location          = Get-ComputerLocation -jsonFilePath $jsonFilePath
}

    $ExportPath = Join-Path $FastCruiseReportPath $FastCruiseFile
    $ComputerStat | Export-Csv -Path $ExportPath -NoTypeInformation -Append -Force
}

# Call the main function
Start-FastCruise