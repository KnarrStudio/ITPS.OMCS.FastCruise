function Get-MacAddress
{
    <#
        .SYNOPSIS
        Returns the MAC address of the active physical network adapter.

        .DESCRIPTION
        Reads the MAC address from the first physical adapter reporting an
        'Up' status via Get-NetAdapter.

        .PARAMETER LastFour
        Returns only the last four hex characters, colon-separated (matches
        the format used elsewhere for Active Directory descriptions).

        .EXAMPLE
        Get-MacAddress

        .EXAMPLE
        Get-MacAddress -LastFour

        .OUTPUTS
        System.String
    #>
    [CmdletBinding()]
    [OutputType([String])]
    param(
        [Parameter(Position = 0)]
        [Switch]$LastFour
    )
    $MacAddress = (Get-NetAdapter -Physical | Where-Object -Property Status -EQ -Value 'Up').MacAddress
    if (-not $MacAddress)
    {
        Write-FCLog -Message 'No active physical network adapter found.' -Level Warning
        return ''
    }
    if ($LastFour)
    {
        return ConvertTo-MacAddressSuffix -MacAddress $MacAddress
    }
    return $MacAddress.ToLower()
}
