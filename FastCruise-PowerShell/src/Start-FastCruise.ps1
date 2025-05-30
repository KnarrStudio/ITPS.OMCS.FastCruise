# filepath: /FastCruise-PowerShell/FastCruise-PowerShell/src/Start-FastCruise.ps1

# This is the main entry point for the FastCruise application.
# It initializes the script and calls the necessary functions to perform the FastCruise operations.

# Import necessary modules and scripts
Import-Module ./Get-InstalledSoftware.psm1 -Verbose
Import-Module ./Get-WorkstationInfo.psm1 -Verbose
Import-Module ./Get-MacAddress.psm1 -Verbose
Import-Module ./Get-ComputerLocation.psm1 -Verbose
Import-Module ./Get-LastComputerStatus.psm1 -Verbose
Import-Module ./Get-FacilityIssues.psm1 -Verbose
Import-Module ./Start-ApplicationTest.psm1 -Verbose
Import-Module ./Show-VbForm.psm1 -Verbose
Import-Module ./Show-AsciiMenu.psm1 -Verbose

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
    $jsonFilePath = "C:\Users\erika\OneDrive\Documents\GitHub\ITPS.OMCS.FastCruise\Configfiles\computerlocation.json"
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
    InstalledSoftware = Get-InstalledSoftware -SoftwareName 'InstallRoot', 'Mozilla Firefox', 'KeePass', 'NetWorx'
    ManualInput       = [bool]$ManualInput
    Location          = Get-ComputerLocation -jsonFilePath $jsonFilePath
}

    $ExportPath = Join-Path $FastCruiseReportPath $FastCruiseFile
    $ComputerStat #| Export-Csv -Path $ExportPath -NoTypeInformation -Append -Force
}

# Call the main function
Start-FastCruise