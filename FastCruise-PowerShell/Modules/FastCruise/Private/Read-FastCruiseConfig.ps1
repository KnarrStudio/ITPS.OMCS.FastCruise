function Read-FastCruiseConfig
{
    <#
        .SYNOPSIS
        Loads the FastCruise.config.psd1 settings file.

        .DESCRIPTION
        Reads and validates the editable FastCruise configuration file. This
        is the only place site-specific settings (paths, software checks,
        desk labels, application tests, logging options) should ever need to
        change - editing this file should never require editing a .ps1 or
        .psm1 file.

        .PARAMETER ConfigPath
        Path to the FastCruise.config.psd1 file.

        .EXAMPLE
        Read-FastCruiseConfig -ConfigPath 'C:\FastCruise-PowerShell\Config\FastCruise.config.psd1'

        .OUTPUTS
        System.Collections.Hashtable
    #>
    [CmdletBinding()]
    [OutputType([Hashtable])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateScript({
            if (-not (Test-Path -Path $_))
            {
                throw ('Config file not found: {0}' -f $_)
            }
            $true
        })]
        [String]$ConfigPath
    )
    Write-FCLog -Message ('Loading configuration from {0}' -f $ConfigPath) -Level Debug
    $Config = Import-PowerShellDataFile -Path $ConfigPath

    $RequiredKeys = @('ReportRoot', 'ReportFileName', 'LocalReportPath', 'LocationConfigFile', 'SoftwareChecks', 'DeskLabels')
    foreach ($Key in $RequiredKeys)
    {
        if (-not $Config.ContainsKey($Key))
        {
            throw ('FastCruise config file is missing required key: {0}' -f $Key)
        }
    }
    return $Config
}
