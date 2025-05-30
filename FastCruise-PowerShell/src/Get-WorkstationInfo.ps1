function Get-WorkstationInfo    
{
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet('Manufacturer', 'Model', 'Name', 'PrimaryOwnerName', 'Domain', 'serialnumber', 'PartOfDomain', 'Workgroup')] 
        [String]$Info
    )
    if ($Info -eq 'serialnumber') {
        (Get-CimInstance -ClassName win32_SystemEnclosure).serialnumber
    } else {
        (Get-CimInstance -ClassName Win32_ComputerSystem).$Info
    }
}