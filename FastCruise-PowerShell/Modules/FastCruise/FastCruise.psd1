@{
    RootModule           = 'FastCruise.psm1'
    ModuleVersion        = '2.0.0'
    GUID                 = 'a1c9d3e4-2b8f-4a7e-9c5a-1d6f3e9b7a44'
    Author               = 'KnarrStudio'
    CompanyName          = 'KnarrStudio'
    Copyright            = '(c) KnarrStudio. GPL-3.0 license.'
    Description          = 'A quick operational test ("Fast Cruise") of workstations. Captures the date, time, and user who performed the check, records hardware/software facts automatically, and prompts for physical location, phone number, notes, and facility issues.'
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    RequiredModules      = @()
    FunctionsToExport    = @('Start-FastCruise', 'Export-ComputerDescription', 'ConvertTo-LocationJson', 'Show-FastCruiseMenu')
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()
    PrivateData          = @{
        PSData = @{
            Tags       = @('FastCruise', 'ITPS.OMCS', 'OperationalTest')
            LicenseUri = 'https://github.com/KnarrStudio/ITPS.OMCS.FastCruise/blob/Development/LICENSE'
            ProjectUri = 'https://github.com/KnarrStudio/ITPS.OMCS.FastCruise'
        }
    }
}
