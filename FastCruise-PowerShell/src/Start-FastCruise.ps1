# filepath: /FastCruise-PowerShell/FastCruise-PowerShell/src/Start-FastCruise.ps1

# This is the main entry point for the FastCruise application.
# It initializes the script and calls the necessary functions to perform the FastCruise operations.

# Import necessary modules and scripts
Import-Module -Name ./Get-InstalledSoftware.psm1 -Verbose
Import-Module -Name ./Get-WorkstationInfo.psm1 -Verbose
Import-Module -Name ./Get-MacAddress.psm1 -Verbose
Import-Module -Name ./Get-ComputerLocation.psm1 -Verbose
Import-Module -Name ./Get-LastComputerStatus.psm1 -Verbose
Import-Module -Name ./Get-FacilityIssues.psm1 -Verbose
Import-Module -Name ./Start-ApplicationTest.psm1 -Verbose
Import-Module -Name ./Show-VbForm.psm1 -Verbose
Import-Module -Name ./Show-AsciiMenu.psm1 -Verbose

param (
    [string]$FastCruiseReportPath = 'S:\',
    [string]$FastCruiseFile = 'FastCruise.csv',
    [switch]$ManualInput = $false,
    [string]$JsonFilePath = "C:\Users\erika\OneDrive\Documents\GitHub\ITPS.OMCS.FastCruise\Configfiles\computerlocation.json",
    [string]$LocalCruiseFile = "$env:HOMEDRIVE\temp\FastCruise\FastCruiseFile.csv"
)

# Check if the S: drive is mapped
if (-not (Get-PSDrive -Name S -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name S -PSProvider FileSystem -Root '\\localhost\Folder-1' -Persist
}

# Export file path
$ExportPath = Join-Path -Path $FastCruiseReportPath -ChildPath $FastCruiseFile

# Create Export file if it doesn't exist
if (-not (Test-Path -Path $ExportPath)) {
    Write-Verbose -Message 'Creating Export file.'
    $null = New-Item -Path $ExportPath -ItemType File -Force
}


# Create local file if it doesn't exist
if (-not (Test-Path -Path $LocalCruiseFile)) {
    Write-Verbose -Message 'Creating Local File.'
    $null = New-Item -Path $LocalCruiseFile -ItemType File -Force
}
 
    # Get installed software details as objects
    $SoftwareList = Get-InstalledSoftware -SoftwareName 'InstallRoot', 'Mozilla Firefox', 'KeePass', 'NetWorx'

    # Flatten software details for CSV columns
    $Software_Installdate     = ($SoftwareList | ForEach-Object { $_.Installdate })     -join '; '
    $Software_DisplayVersion  = ($SoftwareList | ForEach-Object { $_.DisplayVersion })  -join '; '
    $Software_DisplayName     = ($SoftwareList | ForEach-Object { $_.DisplayName })     -join '; '

    # Get location details
    $Location = Get-ComputerLocation -jsonFilePath $jsonFilePath

    $ComputerStat = [PSCustomObject]@{
        ComputerName           = Get-WorkstationInfo -Info 'Name'
        SerialNumber           = Get-WorkstationInfo -Info 'serialnumber'
        MacAddress             = Get-MacAddress
        Software_Installdate   = $Software_Installdate
        Software_DisplayVersion= $Software_DisplayVersion
        Software_DisplayName   = $Software_DisplayName
        ManualInput            = [bool]$ManualInput
        Location_Department    = $Location.Department
        Location_Building      = $Location.Building
        Location_Room          = $Location.Room
        Location_Desk          = $Location.Desk
    }

    
    $ComputerStat #| Export-Csv -Path $ExportPath -NoTypeInformation -Append -Force
