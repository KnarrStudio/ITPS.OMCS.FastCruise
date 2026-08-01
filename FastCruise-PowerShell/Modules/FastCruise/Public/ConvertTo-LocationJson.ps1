function ConvertTo-LocationJson
{
    <#
        .SYNOPSIS
        Converts a Department/Building/Room hashtable file into ComputerLocation.json.

        .DESCRIPTION
        Hand-writing the Department > Building > Room hierarchy directly as
        JSON is easy to get wrong (matching braces, commas, quoting). This
        function lets you describe the same hierarchy as a plain PowerShell
        data file (.psd1) instead, which is far easier to edit, and converts
        it to the JSON file Get-ComputerLocation reads.

        The source hashtable is no longer hardcoded inside this function - see
        Config\ComputerLocation.hashtable.psd1 for an editable example.

        .PARAMETER SourceHashtableFile
        Path to a .psd1 file containing the Department/Building/Room
        hashtable.

        .PARAMETER DestinationJsonFile
        Path to write the resulting JSON file to.

        .EXAMPLE
        ConvertTo-LocationJson -SourceHashtableFile 'Config\ComputerLocation.hashtable.psd1' -DestinationJsonFile 'Config\ComputerLocation.json'

        .INPUTS
        .PSD1

        .OUTPUTS
        .JSON
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [String]$SourceHashtableFile,
        [Parameter(Mandatory, Position = 1)]
        [String]$DestinationJsonFile
    )
    Write-FCLog -Message ('Converting {0} to {1}' -f $SourceHashtableFile, $DestinationJsonFile) -Level Info

    $LocationHash = Import-PowerShellDataFile -Path $SourceHashtableFile
    $Json = $LocationHash | ConvertTo-Json -Depth 5
    Set-FCContent -Path $DestinationJsonFile -Content $Json

    Write-FCLog -Message ('Wrote {0}' -f $DestinationJsonFile) -Level Info
}
