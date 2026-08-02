function Export-ComputerDescription
{
    <#
        .SYNOPSIS
        Builds an Active Directory description list from a Fast Cruise report.

        .DESCRIPTION
        Reads a Fast Cruise CSV report and produces a second CSV containing
        each ComputerName and a formatted ComputerDescription (department,
        building, room, desk, and the last four characters of the MAC
        address), suitable for updating computer descriptions in Active
        Directory.

        .PARAMETER InputReportFile
        A Fast Cruise CSV report with ComputerName, Department, Building,
        Room, Desk, and MacAddress columns.

        .PARAMETER OutputListFile
        Path to write the resulting ComputerName/ComputerDescription CSV to.

        .PARAMETER DescriptionPrefix
        Text prepended to every generated description. Defaults to
        'KnarrStudio'.

        .EXAMPLE
        Export-ComputerDescription -InputReportFile 'C:\temp\Reports\FastCruise_2026-July.csv' -OutputListFile 'C:\temp\Reports\ComputerDescriptions.csv'

        .NOTES
        Intended to be run after Start-FastCruise to build a location
        reference list for every system that has completed a Fast Cruise.

        .INPUTS
        .CSV

        .OUTPUTS
        .CSV
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateScript({
            if ($_ -match '\.csv$')
            {
                $true
            }
            else
            {
                throw 'Input file needs to be a .csv file.'
            }
        })]
        [String]$InputReportFile,
        [Parameter(Mandatory, Position = 1)]
        [String]$OutputListFile,
        [Parameter(Position = 2)]
        [String]$DescriptionPrefix = 'KnarrStudio'
    )
    Write-FCLog -Message ('Building computer description list from {0}' -f $InputReportFile) -Level Info

    if (Test-Path -Path $OutputListFile)
    {
        # This list is a full rebuild each run, not an append.
        Remove-Item -Path $OutputListFile -Force
    }

    $FastCruiseData = Import-FastCruiseRecord -Path $InputReportFile | Sort-Object -Property Department, Building

    foreach ($Record in $FastCruiseData)
    {
        $MacSuffix = ConvertTo-MacAddressSuffix -MacAddress $Record.MacAddress
        $Description = [PSCustomObject]@{
            ComputerName        = $Record.ComputerName
            ComputerDescription = '{0}-{1}-{2}-{3}{4} [{5}]' -f $DescriptionPrefix, $Record.Department, $Record.Building, $Record.Room, $Record.Desk, $MacSuffix
        }
        Export-FastCruiseRecord -Path $OutputListFile -Record $Description
    }
    Write-FCLog -Message ('Wrote description list to {0}' -f $OutputListFile) -Level Info
}
