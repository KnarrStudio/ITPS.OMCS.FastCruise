function Get-FacilityIssues
{
    <#
        .SYNOPSIS
        Collects facility issues (broken lights, A/C, etc.) for the current room.

        .DESCRIPTION
        Appends any issues the user enters to a weekly rolling report file, so
        one file per Monday-through-Sunday week accumulates every issue
        reported in that room. Prompts repeatedly until the user enters
        nothing, "Exit", or the default placeholder text.

        .PARAMETER RoomStatusFile
        Base path for the facility-issue report. The current week is inserted
        into the file name (e.g. Facility_Issue_Report-WeekOf_July-27.txt).

        .PARAMETER LatestStatus
        An object with Building and Room properties, used in the report header.

        .EXAMPLE
        Get-FacilityIssues -RoomStatusFile 'S:\FC-Facility_Issue\Facility_Issue_Report.txt' -LatestStatus $ComputerStat
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [String]$RoomStatusFile,
        [Parameter(Mandatory, Position = 1)]
        [Object]$LatestStatus
    )
    $DblLine    = '=' * 20
    $DateFormat = 'MMMM-dd'

    # Facility issues are grouped by the Monday that starts the current week.
    $DaysSinceMonday = @{
        Monday = 0; Tuesday = 1; Wednesday = 2; Thursday = 3
        Friday = 4; Saturday = 5; Sunday = 6
    }
    $CurrentDay  = Get-Date -UFormat '%A'
    $CurrentWeek = (Get-Date).AddDays(-1 * $DaysSinceMonday[$CurrentDay]).ToString($DateFormat)

    $FacilityIssueReport = $RoomStatusFile.Replace('.', ('-WeekOf_{0}.' -f $CurrentWeek))
    if (-not (Test-Path -Path $FacilityIssueReport))
    {
        Write-FCLog -Message ('Creating facility issue report for the week of {0}' -f $CurrentWeek) -Level Info
        Add-FCContent -Path $FacilityIssueReport -Line 'Fast Cruise Room Discrepancy Report'
    }

    $Header = @"
$DblLine
$(Get-Date) - Building: $($LatestStatus.Building)  Room: $($LatestStatus.Room)

"@
    Add-FCContent -Path $FacilityIssueReport -Line $Header

    $DefaultInput = 'None Found'
    $DoNotWrite   = @('', ' ', 'Exit')
    do
    {
        $RoomIssue = Show-InputDialog -Message 'Enter any issues with the room. These will be sent to the facilities department. Type "EXIT" to exit' -Title 'Facility Issue Report' -DefaultValue $DefaultInput
        if ($RoomIssue -notin $DoNotWrite)
        {
            Add-FCContent -Path $FacilityIssueReport -Line ('- {0}' -f $RoomIssue)
            if ($RoomIssue -eq $DefaultInput)
            {
                $DoNotWrite += $RoomIssue
            }
        }
    }
    while ($RoomIssue -notin $DoNotWrite)
}
