function Import-FastCruiseRecord
{
    <#
        .SYNOPSIS
        Reads a FastCruise CSV report in a uniform format.

        .DESCRIPTION
        Thin wrapper around Import-Csv that pins the encoding to UTF-8 so
        every read goes through the same code path as Export-FastCruiseRecord,
        regardless of PowerShell edition.

        .PARAMETER Path
        Full path of the CSV file to read.

        .EXAMPLE
        Import-FastCruiseRecord -Path $ReportPath

        .OUTPUTS
        PSCustomObject[]
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [String]$Path
    )
    if (-not (Test-Path -Path $Path))
    {
        Write-FCLog -Message ('Import-FastCruiseRecord: {0} does not exist.' -f $Path) -Level Warning
        return @()
    }
    return Import-Csv -Path $Path -Encoding UTF8
}
