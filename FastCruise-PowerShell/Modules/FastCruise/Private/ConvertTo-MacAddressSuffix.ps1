function ConvertTo-MacAddressSuffix
{
    <#
        .SYNOPSIS
        Returns the last four hex characters of a MAC address, colon-separated.

        .DESCRIPTION
        Small shared helper so the "last four of the MAC" formatting logic
        exists in exactly one place instead of being duplicated between the
        live MAC lookup (Get-MacAddress) and report post-processing
        (Export-ComputerDescription).

        .PARAMETER MacAddress
        A MAC address string in dash-separated form, e.g. '00-11-22-33-44-55'.

        .EXAMPLE
        ConvertTo-MacAddressSuffix -MacAddress '00-11-22-33-44-55'
        Returns '44:55'.

        .OUTPUTS
        System.String
    #>
    [CmdletBinding()]
    [OutputType([String])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [String]$MacAddress
    )
    return ($MacAddress -replace '^(.+-)(..)-(..)$', '$2:$3').ToLower()
}
