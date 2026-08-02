function Get-ComputerLocation
{
    <#
        .SYNOPSIS
        Asks the user to pick this workstation's Department, Building, Room, and Desk.

        .DESCRIPTION
        If LocationConfigFile exists, its Department > Building > Room
        hierarchy is used to drive cascading drop-down pickers. If the file is
        missing, the user is prompted for free-text Department/Building/Room
        values instead. The desk is always chosen from DeskLabels.

        JSON returned by ConvertFrom-Json is already a navigable object
        (PSCustomObject with nested PSCustomObjects), so this function reads
        it directly - no separate hashtable-conversion step is needed.

        .PARAMETER LocationConfigFile
        Path to the Department/Building/Room JSON reference file.

        .PARAMETER DeskLabels
        The list of desk labels to offer in the desk picker.

        .EXAMPLE
        Get-ComputerLocation -LocationConfigFile 'S:\ComputerLocation.json' -DeskLabels @('A','B','C')

        .OUTPUTS
        PSCustomObject with Department, Building, Room, and Desk properties.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [AllowNull()]
        [String]$LocationConfigFile,
        [Parameter(Mandatory, Position = 1)]
        [String[]]$DeskLabels
    )
    if ($LocationConfigFile -and (Test-Path -Path $LocationConfigFile -ErrorAction SilentlyContinue))
    {
        Write-FCLog -Message ('Using location reference file: {0}' -f $LocationConfigFile) -Level Verbose
        $Location = Get-Content -Path $LocationConfigFile -Raw | ConvertFrom-Json

        $DeptNames = $Location.Department.PSObject.Properties.Name
        $Department = Show-SelectionDialog -Message 'Select Department:' -Options $DeptNames -Title 'Department'

        $BuildingNames = $Location.Department.$Department.Building.PSObject.Properties.Name
        $Building = Show-SelectionDialog -Message 'Select Building:' -Options $BuildingNames -Title 'Building'

        $RoomList = $Location.Department.$Department.Building.$Building.Room
        $Room = Show-SelectionDialog -Message 'Select Room:' -Options $RoomList -Title 'Room'
    }
    else
    {
        Write-FCLog -Message ('Unable to find or use location reference file: {0}' -f $LocationConfigFile) -Level Warning
        $Department = Show-InputDialog -Message 'Department: Produce, Bakery, Dairy' -Title 'Department' -DefaultValue 'Other'
        $Building   = Show-InputDialog -Message 'Building: Office-4, Bay-34' -Title 'Building' -DefaultValue 'Office'
        $Room       = Show-InputDialog -Message 'Room Number:' -Title 'Room' -DefaultValue '1'
    }
    $Desk = Show-SelectionDialog -Message 'Select Desk:' -Options $DeskLabels -Title 'Desk'

    [PSCustomObject]@{
        Department = $Department
        Building   = $Building
        Room       = $Room
        Desk       = $Desk
    }
}
