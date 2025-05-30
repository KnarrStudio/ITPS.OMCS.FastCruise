function Get-WorkstationInfo    
{
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet('Manufacturer', 'Model', 'Name', 'PrimaryOwnerName', 'Domain', 'serialnumber', 'PartOfDomain', 'Workgroup')] 
        [String]$Info
    )
    if ($Info -eq 'serialnumber') {
        (Get-WmiObject -Class win32_SystemEnclosure).serialnumber
    } else {
        (Get-WmiObject -Class:Win32_ComputerSystem).$Info
    }
}