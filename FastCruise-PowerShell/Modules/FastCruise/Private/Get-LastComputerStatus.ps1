function Get-LastComputerStatus
{
    <#
        .SYNOPSIS
        Returns the most recent Fast Cruise record for this workstation.

        .DESCRIPTION
        Reads a FastCruise CSV report and returns the last row whose
        ComputerName matches the local computer name, if any.

        .PARAMETER LastCruiseStatus
        Path to the FastCruise CSV report to search.

        .EXAMPLE
        Get-LastComputerStatus -LastCruiseStatus 'C:\temp\FastCruise\FastCruiseFile.csv'

        .OUTPUTS
        PSCustomObject, or $null if no prior record exists for this computer.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [String]$LastCruiseStatus
    )
    Write-FCLog -Message ('Looking up last recorded status for {0}' -f $env:COMPUTERNAME) -Level Verbose
    $Import = Import-FastCruiseRecord -Path $LastCruiseStatus
    $LatestStatus = $Import |
        Where-Object -FilterScript { $PSItem.ComputerName -eq $env:COMPUTERNAME } |
        Select-Object -Last 1

    if (-not $LatestStatus)
    {
        Write-FCLog -Message 'No existing record found for this workstation.' -Level Info
    }
    return $LatestStatus
}
