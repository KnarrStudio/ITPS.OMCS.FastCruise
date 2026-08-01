function Export-FastCruiseRecord
{
    <#
        .SYNOPSIS
        Appends one record to a FastCruise CSV report in a uniform format.

        .DESCRIPTION
        Every FastCruise CSV report (the shared network report, the local
        per-workstation copy, the exported description list) is written
        through this one function so column quoting, line endings, and
        encoding are identical no matter which script or which PowerShell
        edition produced the file. Writes the header line automatically the
        first time a record is written to a given path.

        .PARAMETER Path
        Full path of the CSV file to append to.

        .PARAMETER Record
        The record to write, as an ordered hashtable or PSCustomObject.

        .EXAMPLE
        Export-FastCruiseRecord -Path $ReportPath -Record $ComputerStat
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [String]$Path,
        [Parameter(Mandatory, Position = 1)]
        [Object]$Record
    )
    $RecordObject = [PSCustomObject]$Record
    $NeedsHeader  = (-not (Test-Path -Path $Path)) -or ((Get-Item -Path $Path).Length -eq 0)
    $Directory    = Split-Path -Path $Path -Parent
    if ($Directory -and -not (Test-Path -Path $Directory))
    {
        $null = New-Item -Path $Directory -ItemType Directory -Force
    }

    $Csv = $RecordObject | ConvertTo-Csv -NoTypeInformation
    if (-not $NeedsHeader)
    {
        # Header already exists in the file, only append the data row.
        $Csv = $Csv | Select-Object -Skip 1
    }
    foreach ($Line in $Csv)
    {
        Add-FCContent -Path $Path -Line $Line
    }
    Write-FCLog -Message ('Wrote 1 record to {0}' -f $Path) -Level Debug
}
