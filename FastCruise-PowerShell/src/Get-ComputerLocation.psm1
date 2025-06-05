function Get-ComputerLocation     
{
    param
    (
        [Parameter(Mandatory = $false, Position = 0)]
        [String]$jsonFilePath = "$PSScriptRoot\..\config\ComputerLocation.json"
    )

    if (Test-Path -Path $jsonFilePath -ErrorAction SilentlyContinue)
    {
        Write-Verbose -Message 'Using JSON File'
        $location = Get-Content -Path $jsonFilePath -Raw | ConvertFrom-Json

        # Department selection
        $deptNames = $location.Department.PSObject.Properties.Name
        $LclDept = Show-VbDropdownForm -Message 'Select Department:' -InputFile ([string[]]$deptNames) -TitleBar 'Department'

        # Building selection
        $buildNames = $location.Department.$LclDept.Building.PSObject.Properties.Name
        $LclBuild = Show-VbDropdownForm -Message 'Select Building:' -InputFile ([string[]]$buildNames) -TitleBar 'Building'

        # Room selection
        $roomList = $location.Department.$LclDept.Building.$LclBuild.Room
        $LclRm = Show-VbDropdownForm -Message 'Select Room:' -InputFile ([string[]]$roomList) -TitleBar 'Room'

        # Desk selection
        if (-not $Desk) { $Desk = @('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q') }
        $LclDesk = Show-VbDropdownForm -Message 'Select Desk:' -InputFile ([string[]]$Desk) -TitleBar 'Desk'
    }
    else
    {
        Write-Error "Unable to find or use JSON File: $jsonFilePath"
        # [TAGGED: ManualInputFallback] -- Placeholder for future manual input logic
        return
    }

    # Return the collected data as a PSCustomObject
    [PSCustomObject]@{
        Department = $LclDept
        Building   = $LclBuild
        Room       = $LclRm
        Desk       = $LclDesk
    }
}

Export-ModuleMember -Function Get-ComputerLocation