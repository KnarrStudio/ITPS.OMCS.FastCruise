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

        # Direct property access from JSON object
        $deptNames = $location.Department.PSObject.Properties.Name
        $LclDept = $deptNames | Out-GridView -Title 'Department' -OutputMode Single

        $buildNames = $location.Department.$LclDept.Building.PSObject.Properties.Name
        $LclBuild = $buildNames | Out-GridView -Title 'Building' -OutputMode Single

        $roomList = $location.Department.$LclDept.Building.$LclBuild.Room
        $LclRm = $roomList | Out-GridView -Title 'Room' -OutputMode Single

        # If $Desk is not defined in this scope, define a default list
        if (-not $Desk) { $Desk = @('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q') }
        $LclDesk = $Desk | Out-GridView -Title 'Desk' -OutputMode Single
    }
    else
    {
        Write-Verbose -Message 'Unable to find or use JSON File'
        $LclDept = Show-VbForm -InputBox -Message 'Department: Produce, Bakery, Dairy' -TitleBar 'Department' -DefaultValue 'Other'
        $LclBuild = Show-VbForm -InputBox -Message 'Building: Office-4, Bay-34' -TitleBar 'Building' -DefaultValue 'Office'
        $LclRm = Show-VbForm -InputBox -Message 'Room Number:' -TitleBar 'Room' -DefaultValue 1
        if (-not $Desk) { $Desk = @('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q') }
        $LclDesk = $Desk | Out-GridView -Title 'Desk' -OutputMode Single
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