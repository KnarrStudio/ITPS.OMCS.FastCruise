function Start-FastCruise
{
    <#
        .SYNOPSIS
        Runs one Fast Cruise operational check of this workstation.

        .DESCRIPTION
        Captures automated workstation facts (serial number, MAC address,
        installed software versions, WSUS status, domain/workgroup) and walks
        the person running the check through a short set of prompts (physical
        location, phone number, notes, optional application tests, optional
        facility issues). The combined record is appended to the shared
        network report and to a local per-workstation copy.

        All site-specific settings (paths, software checks, desk labels,
        application tests) come from the FastCruise.config.psd1 file rather
        than from values hardcoded in this function - see -ConfigPath.

        .PARAMETER ConfigPath
        Path to FastCruise.config.psd1. Defaults to the Config folder shipped
        alongside this module.

        .PARAMETER ReportPath
        Overrides the shared report directory from the config file.

        .PARAMETER ReportFileName
        Overrides the shared report file name from the config file.

        .PARAMETER ManualInput
        Skips automatic hardware/software detection and prompts for a
        computer name instead. Used when a workstation cannot run the
        detection commands itself (e.g. recording a kiosk or loaner device).

        .EXAMPLE
        Start-FastCruise
        Runs a normal Fast Cruise using every default from the config file.

        .EXAMPLE
        Start-FastCruise -ManualInput
        Records a Fast Cruise entry by hand instead of via auto-detection.

        .EXAMPLE
        Start-FastCruise -ReportPath 'S:\FastCruise' -ReportFileName 'FastCruise.csv' -Verbose
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [String]$ConfigPath = $Script:DefaultConfigPath,
        [Parameter(Position = 1)]
        [String]$ReportPath,
        [Parameter(Position = 2)]
        [String]$ReportFileName,
        [Parameter(Position = 3)]
        [Switch]$ManualInput
    )
    Begin
    {
        $Config = Read-FastCruiseConfig -ConfigPath $ConfigPath
        if (-not (Get-FCLogPath))
        {
            $null = Initialize-FCLog -LogDirectory $Config.Logging.LogDirectory -MinimumLevel $Config.Logging.MinimumLevel
        }
        Write-FCLog -Message 'Starting Fast Cruise.' -Level Info

        if (-not $ReportPath)
        {
            $ReportPath = $Config.ReportRoot
        }
        if (-not $ReportFileName)
        {
            $ReportFileName = $Config.ReportFileName
        }

        # Map the shared network drive if it isn't already available.
        if (-not (Test-Path -Path ('{0}\' -f $Config.NetworkDriveLetter)))
        {
            Write-FCLog -Message ('Mapping {0} to {1}' -f $Config.NetworkDriveLetter, $Config.NetworkSharePath) -Level Warning
            net.exe use $Config.NetworkDriveLetter $Config.NetworkSharePath
        }

        $ComputerName = $env:COMPUTERNAME
        $YearMonth = Get-Date -Format 'yyyy-MMMM'
        $ReportFileName = $ReportFileName.Replace('.', ('_{0}.' -f $YearMonth))

        $LocalReportPath = $Config.LocalReportPath
        if (-not (Test-Path -Path $LocalReportPath))
        {
            Write-FCLog -Message 'Creating local report file.' -Level Verbose
            $LocalDirectory = Split-Path -Path $LocalReportPath -Parent
            if ($LocalDirectory -and -not (Test-Path -Path $LocalDirectory))
            {
                $null = New-Item -Path $LocalDirectory -ItemType Directory -Force
            }
            $null = New-Item -Path $LocalReportPath -ItemType File -Force
        }

        try
        {
            # Confirms the workstation can see the domain, which also implies
            # the network report path should be reachable.
            [void][System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()
            Write-FCLog -Message ('Authentication Server: {0}' -f $env:LOGONSERVER) -Level Verbose
            if (-not (Test-Path -Path $ReportPath))
            {
                $null = New-Item -Path $ReportPath -ItemType Directory -Force
            }
        }
        catch
        {
            Write-FCLog -Message ('Network path not available. Using local workstation: {0}' -f $ComputerName) -Level Warning
            $ReportPath = $env:TEMP
        }

        $Report = Join-Path -Path $ReportPath -ChildPath $ReportFileName
        Write-FCLog -Message ('Report path: {0}' -f $Report) -Level Verbose

        # Every column FastCruise could possibly write is declared here, up
        # front, with a placeholder value. Different runs take different
        # branches below (domain vs. workgroup, manual vs. automatic,
        # whichever software/application checks are configured) - without a
        # fixed schema, rows appended to the same shared CSV would end up
        # with different column sets, which silently misaligns the report
        # (and is exactly what Export-Csv -Append refuses to do on
        # Windows PowerShell when column sets differ).
        $ComputerStat = [Ordered]@{
            'Date'                 = "$(Get-Date)"
            'ComputerName'         = $ComputerName
            'SerialNumber'         = 'N/A'
            'MacAddress'           = 'N/A'
            'UserName'             = $env:USERNAME
            'Manufacturer'         = 'N/A'
            'Model'                = 'N/A'
            'Domain'               = 'N/A'
            'WorkGroup'            = 'N/A'
            'WSUS Search Success'  = 'N/A'
            'WSUS Install Success' = 'N/A'
        }
        foreach ($SoftwareItem in $Config.SoftwareChecks)
        {
            $ComputerStat["$SoftwareItem Version"] = 'N/A'
        }
        foreach ($TestName in $Config.ApplicationTests.Keys)
        {
            $ComputerStat["$TestName Test"] = 'N/A'
        }
        $ComputerStat['Department'] = 'N/A'
        $ComputerStat['Building']   = 'N/A'
        $ComputerStat['Room']       = 'N/A'
        $ComputerStat['Desk']       = 'N/A'
        $ComputerStat['Phone']      = 'N/A'
        $ComputerStat['Notes']      = ''

        $LatestStatus = $null
        $LocationVerified = $false
    } #End BEGIN region
    Process
    {
        if (-not $ManualInput)
        {
            try
            {
                $WsusResults = (New-Object -ComObject 'Microsoft.Update.AutoUpdate').Results
                $ComputerStat['WSUS Search Success']  = $WsusResults.LastSearchSuccessDate
                $ComputerStat['WSUS Install Success'] = $WsusResults.LastInstallationSuccessDate
            }
            catch
            {
                Write-FCLog -Message ('Unable to read WSUS status: {0}' -f $_.Exception.Message) -Level Warning
            }

            Write-FCLog -Message 'Collecting workstation details.' -Level Verbose
            $ComputerStat['MacAddress']   = Get-MacAddress
            $ComputerStat['SerialNumber'] = Get-WorkstationInfo -Info serialnumber
            $ComputerStat['Manufacturer'] = Get-WorkstationInfo -Info Manufacturer
            $ComputerStat['Model']        = Get-WorkstationInfo -Info Model

            if (Get-WorkstationInfo -Info PartOfDomain)
            {
                $ComputerStat['Domain'] = Get-WorkstationInfo -Info Domain
            }
            else
            {
                $ComputerStat['WorkGroup'] = Get-WorkstationInfo -Info Workgroup
                $Report = Join-Path -Path $env:TEMP -ChildPath 'FastCruise.csv'
                Write-FCLog -Message ('This computer is not attached to the domain. Reporting to {0}' -f $Report) -Level Warning
            }

            foreach ($SoftwareItem in $Config.SoftwareChecks)
            {
                $ComputerStat["$SoftwareItem Version"] = Get-InstalledSoftware -SoftwareName $SoftwareItem -SelectParameter DisplayVersion
            }

            $StatusLookupPath = if ((Get-Item -Path $LocalReportPath).Length -gt 0) { $LocalReportPath } else { $Report }
            $LatestStatus = Get-LastComputerStatus -LastCruiseStatus $StatusLookupPath

            if ($Config.RunApplicationTests)
            {
                foreach ($TestName in $Config.ApplicationTests.Keys)
                {
                    $Test = $Config.ApplicationTests[$TestName]
                    $ComputerStat["$TestName Test"] = Start-ApplicationTest -WaitTest -TestFile $Test.TestFile -TestProgram $Test.TestProgram -ProcessName $Test.ProcessName
                }
            }

            if ($LatestStatus)
            {
                $LocationSummary = @"

ComputerName: (Asset Tag)
- $($LatestStatus.ComputerName)

Serial Number:
- $($ComputerStat.SerialNumber)

Department:
- $($LatestStatus.Department)

Building:
- $($LatestStatus.Building)

Room:
- $($LatestStatus.Room)

Desk:
- $($LatestStatus.Desk)

Phone:
- $($LatestStatus.Phone)

"@
                $LocationVerified = Show-ConfirmDialog -Message $LocationSummary -Title 'Is this location still correct?'
            }
        }
        else
        {
            $ComputerStat['ComputerName'] = Show-InputDialog -Message 'ComputerName: (Asset Tag)' -Title 'ComputerName' -DefaultValue 'D1234567'
            $ComputerStat['SerialNumber'] = 'Manual Input'
            foreach ($TestName in $Config.ApplicationTests.Keys)
            {
                $ComputerStat["$TestName Test"] = 'Manual Input'
            }
        }

        if ($LocationVerified)
        {
            $ComputerStat['Department'] = $LatestStatus.Department
            $ComputerStat['Building']   = $LatestStatus.Building
            $ComputerStat['Room']       = $LatestStatus.Room
            $ComputerStat['Desk']       = $LatestStatus.Desk
            $ComputerStat['Phone']      = $LatestStatus.Phone
        }
        else
        {
            $Location = Get-ComputerLocation -LocationConfigFile $Config.LocationConfigFile -DeskLabels $Config.DeskLabels
            Write-FCLog -Message ('Computer Description: {0}-{1}-{2}{3}' -f $Location.Department, $Location.Building, $Location.Room, $Location.Desk) -Level Verbose
            $ComputerStat['Department'] = $Location.Department
            $ComputerStat['Building']   = $Location.Building
            $ComputerStat['Room']       = $Location.Room
            $ComputerStat['Desk']       = $Location.Desk

            $Phone = ''
            while ($Phone -notmatch $Config.PhoneNumberPattern)
            {
                $Phone = Show-InputDialog -Message 'Nearest Phone Number:' -DefaultValue '757-555-1234'
                if ($Phone -eq '')
                {
                    break
                }
            }
            $ComputerStat['Phone'] = $Phone
        }

        $Notes = Show-InputDialog -Message 'Notes about this cruise:' -DefaultValue 'Related to the computer'
        if ($Notes -eq 'Related to the computer')
        {
            $Notes = ''
        }
        $ComputerStat['Notes'] = $Notes

        if ($ComputerStat.Desk -in $Config.FacilityIssueDesks)
        {
            Get-FacilityIssues -RoomStatusFile $Config.FacilityIssueFile -LatestStatus $ComputerStat
        }
    } #End PROCESS region
    End
    {
        Export-FastCruiseRecord -Path $Report -Record $ComputerStat
        Export-FastCruiseRecord -Path $LocalReportPath -Record $ComputerStat

        Write-FCLog -Message 'Fast Cruise record saved.' -Level Info
        Write-FCLog -Message ('Local file: {0}' -f $LocalReportPath) -Level Info

        [PSCustomObject]$ComputerStat | Format-Table

        Write-Host -Object 'Fast Cruise shipmates:'
        Import-FastCruiseRecord -Path $Report |
            Select-Object -Last 4 -Property Date, Username, Department, Building, Room, Phone |
            Format-Table
    } #End END region
}
