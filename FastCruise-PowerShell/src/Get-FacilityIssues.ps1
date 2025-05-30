function Get-FacilityIssues {
    param (
        [Parameter(Mandatory = $true)]
        [String]$RoomStatusFile,
        [Parameter(Mandatory = $true)]
        [Object]$LatestStatus
    )
    
    $DblLine = ('=' * 20)
    $dateFormat = 'MMMM-dd'
    $CurrentDay = Get-Date -UFormat %A
    Switch ($CurrentDay) {
        Monday {
            $CurrentWeek = (Get-Date).AddDays(0).ToString($dateFormat)
        }
        Tuesday {
            $CurrentWeek = (Get-Date).AddDays(-1).ToString($dateFormat)
        }
        Wednesday {
            $CurrentWeek  = (Get-Date).AddDays(-2).ToString($dateFormat)
        }
        Thursday {
            $CurrentWeek = (Get-Date).AddDays(-3).ToString($dateFormat)
        }
        Friday {
            $CurrentWeek = (Get-Date).AddDays(-4).ToString($dateFormat)
        }
        Saturday {
            $CurrentWeek = (Get-Date).AddDays(-5).ToString($dateFormat)
        }
        Sunday {
            $CurrentWeek = (Get-Date).AddDays(-6).ToString($dateFormat)
        }
    }
    
    $FacilityIssueReport = [String]$($RoomStatusFile.Replace('.',('-WeekOf_{0}.' -f $CurrentWeek)))
    if (-not (Test-Path -Path $FacilityIssueReport)) {
        $null = New-Item -Path $FacilityIssueReport -ItemType File -Force
        'Fast Cruise Room Discrepancy Report ' | Out-File -FilePath $FacilityIssueReport -Append
    }
    
    $FacilityIssuesHeader = (@'
{2}
{3} - Building: {0}  Room: {1} 

'@ -f $LatestStatus.Building, $LatestStatus.Room, $DblLine, (Get-Date))
    $FacilityIssuesHeader | Out-File -FilePath $FacilityIssueReport -Append

    $DefaultInput = 'None Found'
    $DoNotWrite = @(' ', 'Exit')
    do {
        if ($RoomIssue -ne $DefaultInput) {
            $RoomIssue = Show-VbForm -InputBox -Message 'Enter any issues with the room. These will be sent to the facilities department. Type "EXIT" to exit' -TitleBar 'Facility Issue Report' -DefaultValue $DefaultInput
            
            if ($RoomIssue -notin $DoNotWrite) {
                ('- {0}' -f $RoomIssue) | Out-File -FilePath $FacilityIssueReport -Append
                if ($RoomIssue -match $DefaultInput) {
                    $DoNotWrite += $RoomIssue
                }
            }
        } else {
            $RoomIssue = 'Exit'
        }
    } while ($RoomIssue -notin $DoNotWrite)
}