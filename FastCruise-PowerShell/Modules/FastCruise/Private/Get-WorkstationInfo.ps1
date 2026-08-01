function Get-WorkstationInfo
{
    <#
        .SYNOPSIS
        Returns a single piece of information about the local workstation.

        .DESCRIPTION
        Reads Win32_ComputerSystem (or Win32_SystemEnclosure for the serial
        number) via Get-CimInstance. CIM is used instead of the older
        Get-WmiObject so this works on both Windows PowerShell 5.1 and
        PowerShell 7+.

        .PARAMETER Info
        Which property to return: Manufacturer, Model, Name, PrimaryOwnerName,
        Domain, serialnumber, PartOfDomain, or Workgroup.

        .EXAMPLE
        Get-WorkstationInfo -Info Manufacturer

        .EXAMPLE
        Get-WorkstationInfo -Info serialnumber
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateSet('Manufacturer', 'Model', 'Name', 'PrimaryOwnerName', 'Domain', 'serialnumber', 'PartOfDomain', 'Workgroup')]
        [String]$Info
    )
    Write-FCLog -Message ('Reading workstation info: {0}' -f $Info) -Level Debug
    if ($Info -eq 'serialnumber')
    {
        return (Get-CimInstance -ClassName Win32_SystemEnclosure).SerialNumber
    }
    return (Get-CimInstance -ClassName Win32_ComputerSystem).$Info
}
