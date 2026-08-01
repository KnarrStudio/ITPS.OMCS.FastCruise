@{
    RootModule        = 'FCLogging.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'b3f6a6d0-6e3b-4c2a-9a2e-6e8f5f1a2b10'
    Author            = 'KnarrStudio'
    CompanyName       = 'KnarrStudio'
    Copyright         = '(c) KnarrStudio. GPL-3.0 license.'
    Description       = 'Common, reusable logging module shared by every ITPS.OMCS.FastCruise script and module. Provides one uniform Write-FCLog function so all scripts write console and file logs the same way.'
    PowerShellVersion = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    FunctionsToExport = @('Initialize-FCLog', 'Write-FCLog', 'Get-FCLogPath')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags = @('Logging', 'FastCruise', 'ITPS.OMCS')
        }
    }
}
