function Show-FastCruiseMenu
{
    <#
        .SYNOPSIS
        Shows the interactive follow-up menu after a Fast Cruise check.

        .DESCRIPTION
        Presents a small selection dialog letting the person repeat the Fast
        Cruise check, add a computer manually, record a facility issue,
        restart the computer, or quit. Loops until 'Quit' is chosen or the
        dialog is dismissed.

        .PARAMETER ConfigPath
        Path to FastCruise.config.psd1.

        .EXAMPLE
        Show-FastCruiseMenu -ConfigPath 'C:\FastCruise-PowerShell\Config\FastCruise.config.psd1'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [String]$ConfigPath
    )
    $MenuOptions = @(
        'Repeat Fast Cruise'
        'Manually Add Computer'
        'Record a Facility Issue'
        'Restart Computer'
        'Quit'
    )

    do
    {
        $Choice = Show-SelectionDialog -Message 'What would you like to do next?' -Options $MenuOptions -Title 'Fast Cruise'
        switch ($Choice)
        {
            'Repeat Fast Cruise'
            {
                Start-FastCruise -ConfigPath $ConfigPath
            }
            'Manually Add Computer'
            {
                Start-FastCruise -ConfigPath $ConfigPath -ManualInput
            }
            'Record a Facility Issue'
            {
                $Config = Read-FastCruiseConfig -ConfigPath $ConfigPath
                $Building = Show-InputDialog -Message 'Building for this facility issue:'
                $Room     = Show-InputDialog -Message 'Room for this facility issue:'
                $Location = [PSCustomObject]@{ Building = $Building; Room = $Room }
                Get-FacilityIssues -RoomStatusFile $Config.FacilityIssueFile -LatestStatus $Location
            }
            'Restart Computer'
            {
                Restart-Computer
            }
        }
    }
    while ($Choice -ne 'Quit' -and $Choice -ne '')
}
