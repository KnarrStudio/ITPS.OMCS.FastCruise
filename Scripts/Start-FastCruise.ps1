﻿#requires -Version 3.0 -Modules NetAdapter

if(-not (Test-Path -Path 'S:\'))
{
  Clear-Host
  Write-Warning -Message 'Yo. Mapping your S: Drive'
  Write-Host 'Net Use S: \\localhost\Folder-1' -ForegroundColor Cyan
  net.exe Use S: \\localhost\Folder-1 
}

# Edit the Variables
$SoftwareChecks = @('Axway', 'Mozilla Firefox', 'McAfee Agent', 'Java') 

[Object[]]$Script:Desk = @('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q')

$jsonFilePath = "S:\ComputerLocation.json"

#Edit the splats to customize the script
$FastCruiseSplat = @{
  FastCruiseReportPath = 'S:\FastCruise'
  FastCruiseFile       = 'FastCruise.csv'
  Verbose              = $true
}
$ManualInputSplat = @{
  FastCruiseReportPath = 'S:\FastCruise'
  FastCruiseFile       = 'FastCruise.csv'
  ManualInput          = $true
  Verbose              = $true
}

$PDFApplicationTestSplat = @{
  TestFile    = 'S:\Information-Systems\Scripts\FastCruise\FastCruiseTestFile.pdf'
  TestProgram = "${env:ProgramFiles(x86)}\Adobe\Acrobat 2015\Acrobat\Acrobat.exe"
  ProcessName = 'Acrobat'
}
$PowerPointApplicationTestSplat = @{
  TestFile    = 'S:\Information-Systems\Scripts\FastCruise\FastCruiseTestFile.pptx'
  TestProgram = "${env:ProgramFiles(x86)}\Microsoft Office\Office16\POWERPNT.EXE"
  ProcessName = 'POWERPNT'
}
$FacilityIssuesSplat = @{
  RoomStatusFile = 'S:\FC-Facility_Issue\Facility_Issue_Report.txt'
  Verbose        = $true
}


function Start-FastCruise
{
  param
  (
    [Parameter(Mandatory, Position = 0)]
    [String]$FastCruiseReportPath,
    [Parameter(Mandatory, Position = 1)]
    [ValidateScript({
          If($_ -match '.csv')
          {
            $true
          }
          Else
          {
            Throw 'Input file needs to be CSV'
          }
    })][String]$FastCruiseFile,
    [Parameter(Mandatory = $false, Position = 1)]
    [Switch]$ManualInput
  )
  Begin
  {
    Write-Verbose -Message 'Setup Variables'
    #$LocationVerification = $null
    $ComputerName = $env:COMPUTERNAME
    Write-Verbose -Message 'Setup Report' 
    $YearMonth = Get-Date -Format yyyy-MMMM
    $FastCruiseFile = [String]$($FastCruiseFile.Replace('.',('_{0}.' -f $YearMonth)))
    $LocalCruiseFile = 'C:\temp\FastCruise\FastCruiseFile.csv'
    if(-not (Test-Path -Path $LocalCruiseFile))
    {
      Write-Verbose -Message 'Creating Local File.'
      $null = New-Item -Path $LocalCruiseFile -ItemType File -Force
    }
    try
    {
      # Check if computer is connected to domain network
      [void]::([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain())
      Write-Output -InputObject ('Authentication Server: {0}' -f $env:LOGONSERVER)
      if(-not (Test-Path -Path $FastCruiseReportPath))
      {
        Write-Verbose -Message 'Path not found.  Creating the Directory now.'
        $null = New-Item -Path $FastCruiseReportPath -ItemType Directory -Force
      }
    }
    catch
    {
      Write-Verbose -Message ('Local Workstation: {0}' -f $ComputerName)
      Write-Warning  -Message ('Network path not available') 
      #Write-Output -InputObject ('{0}' -f $FastCruiseReport)
      $FastCruiseReportPath = $env:TEMP
      #$FastCruiseReport = ('{0}\{1}' -f $FastCruiseReportPath, $FastCruiseFile)
    }
    $FastCruiseReport = ('{0}\{1}' -f $FastCruiseReportPath, $FastCruiseFile)
    Write-Verbose -Message ('Report Path: {0}' -f $FastCruiseReportPath)
    Write-Verbose -Message ('Testing the Report Path: {0}' -f $FastCruiseReport)
    if(-not (Test-Path -Path $FastCruiseReport))
    {
      Write-Verbose -Message 'Test Failed.  Creating the File now.'
      $null = New-Item -Path $FastCruiseReport -ItemType File -Force
    } 
    # Variables
    $Phone = $null
    Write-Verbose -Message 'Get-Content of Json File'
    try
    {
      $Script:PhysicalLocations = Get-Content -Path $jsonFile -ErrorAction Stop | ConvertFrom-Json 
      Write-Verbose -Message 'Physical Locations'
      $PhysicalLocations
    }
    catch
    {
      $PhysicalLocations = $null
    }
    <#bookmark Vb Form #>
    function Script:Show-VbForm    
    {
      <#
          .SYNOPSIS
          Creates and displays the form used to make selections or inupt data.
      #>
      [cmdletbinding(DefaultParameterSetName = 'Message')]
      param(
        [Parameter(Position = 0,ParameterSetName = 'Message')]
        [Switch]$YesNoBox,
        [Parameter(Position = 0,ParameterSetName = 'Input')]
        [Switch]$InputBox,
        [Parameter(Mandatory,Position = 1)]
        [string]$Message,
        [Parameter(Position = 2)]
        [string]$TitleBar = 'Fast Cruise',
        [Parameter(Position = 3,ParameterSetName = 'Input')]
        [string]$DefaultValue
      )
      Write-Verbose -Message ('Enter Function: {0}' -f $PSCmdlet.MyInvocation.MyCommand.Name)
      Add-Type -AssemblyName Microsoft.VisualBasic
      switch($PSBoundParameters.Keys){
        'InputBox'
        {
          $Response = [Microsoft.VisualBasic.Interaction]::InputBox($Message, $TitleBar, $DefaultValue)
        }
        'YesNoBox'
        {
          $Response = [Microsoft.VisualBasic.Interaction]::MsgBox($Message, 'YesNo, SystemModal, MsgBoxSetForeground', $TitleBar)
        }
      }
      #Write-host $Response
      $Response
    } # End VbForm-Function
    <#bookmark Application Test #>
    function Start-ApplicationTest    
    {
      <#
          .SYNOPSIS
          Tests the applications passed to see if it will start.
      #>
      param
      (
        [Parameter(Mandatory, Position = 0)]
        [Switch]$WaitTest,
        [Parameter(Mandatory, Position = 1)]
        [string]$TestFile,
        [Parameter(Mandatory, Position = 2)]
        [string]$TestProgram,
        [Parameter(Mandatory, Position = 3)]
        [string]$ProcessName
      )
      $DescriptionLists = [Ordered]@{
        FunctionResult = 'Good', 'Failed'
      }
      Write-Verbose -Message ('Enter Function: {0}' -f $PSCmdlet.MyInvocation.MyCommand.Name)
      try
      {
        Write-Verbose -Message ('Attempting to open {0} with {1}' -f $TestFile, $ProcessName)
        #Start-Process -FilePath $TestProgram -ArgumentList $TestFile
        Start-Process -FilePath $TestFile
      }
      Catch
      {
        Write-Verbose -Message 'TestResult: Failed'
        $TestResult = $DescriptionLists.FunctionResult[1]
        # get error record
        $ErrorMessage  = $_.exception.message
        Write-Verbose -Message ('Error Message: {0}' -f $ErrorMessage)
      }
      if($WaitTest)
      {
        Write-Host -Object ('The Fast Cruise Script will continue after {0} has been closed.' -f $ProcessName) -BackgroundColor Red -ForegroundColor Yellow
        Write-Verbose -Message ('Wait-Process: {0}' -f $ProcessName)
        Wait-Process -Name $ProcessName
      }
      $TestResult = $DescriptionLists.FunctionResult | Out-GridView -Title $ProcessName -OutputMode Single
      Return $TestResult
    } # End ApplicationTest-Function
    <#bookmark Get Computer status last recorded #>
    function Get-LastComputerStatus    
    {
      <#
          .SYNOPSIS
          Return the last status of system based on what was in the current Fast Cruise Report
      #>
      param
      (
        [Parameter(Mandatory, Position = 0)]
        [String]$LastCruiseStatus
      )
      Write-Verbose -Message ('Enter Function: {0}' -f $PSCmdlet.MyInvocation.MyCommand.Name)
      Write-Verbose -Message 'Importing the Fast Cruise Report'
      $CompImport = Import-Csv -Path $LastCruiseStatus
      # Select last status of system.
      Write-Verbose -Message "Getting last status of workstation: $env:COMPUTERNAME"
      try
      {
        $LatestStatus = $CompImport |
        Where-Object -FilterScript {
          $PSItem.ComputerName -eq $env:COMPUTERNAME
        } |
        Select-Object -Last 1 
        if($LatestStatus -eq $null)
        {
          Write-Output -InputObject 'Unable to find an existing record for this system.'
          $Script:Ans = 'NoHistory'
        }
      }
      Catch
      {
        # get error record
        $ErrorMessage  = $_.exception.message
        Write-Verbose -Message ('Error Message: {0}' -f $ErrorMessage)
      }
      Return $LatestStatus
    } # End ComputerStatus-Function
    <#bookmark Computer Location #>
    function Get-ComputerLocation     
    {
      <#
          .SYNOPSIS
          Get-ComputerLocation of workstation
      #>
      param
      (
        [Parameter(Mandatory = $false, Position = 0)]
        [AllowNull()]
        [String]$jsonFilePath
      )
      function Convert-JSONToHash
      {
        param(
          [AllowNull()]
          [Object]$root
        )
        $hash = @{}
        $keys = $root |
        Get-Member -MemberType NoteProperty |
        Select-Object -ExpandProperty Name
        $keys | ForEach-Object -Process {
          $key = $_
          $obj = $root.$($_)
          if($obj -match '@{')
          {
            $nesthash = Convert-JSONToHash -root $obj
            $hash.add($key,$nesthash)
          }
          else
          {
            $hash.add($key,$obj)
          }
        }
        return $hash
      }
      if(Test-Path -Path $jsonFilePath -ErrorAction SilentlyContinue)
      {
        Write-Verbose -Message 'Using JSON File'
        $location = Convert-JSONToHash -root $(Get-Content -Path $jsonFilePath -ErrorAction SilentlyContinue | ConvertFrom-Json)
        [string]$Script:LclDept = $location.Department.keys | Out-GridView -Title 'Department' -OutputMode Single
        [string]$Script:LclBuild = $location.Department[$LclDept].Building.Keys | Out-GridView -Title 'Building' -OutputMode Single
        [string]$Script:LclRm = $location.Department[$LclDept].Building[$LclBuild].Room | Out-GridView -Title 'Room' -OutputMode Single
        [string]$Script:LclDesk = $Desk | Out-GridView -Title 'Desk' -OutputMode Single
      }
      else
      {
        Write-Verbose -Message 'Unable to find or use JSON File'
        [string]$Script:LclDept = Show-VbForm -InputBox -Message 'Department: Produce, Bakery, Dairy' -TitleBar 'Department' -DefaultValue 'Other'
        [string]$Script:LclBuild = Show-VbForm -InputBox -Message 'Building: Office-4, Bay-34' -TitleBar 'Building' -DefaultValue 'Office'
        [string]$Script:LclRm = Show-VbForm -InputBox -Message 'Room Number:' -TitleBar 'Room' -DefaultValue 1
        [string]$Script:LclDesk = $Desk | Out-GridView -Title 'Desk' -OutputMode Single
      }
    } # End Location-Function
    <#bookmark Get Installed Software #>
    Function Get-InstalledSoftware    
    {
      [cmdletbinding(SupportsPaging)]
      Param(
        [Parameter(ValueFromPipeline, Position = 0)]
        [AllowNull()]
        [String[]]$SoftwareName,
        [ValidateSet('DisplayName','DisplayVersion')] 
        [AllowNull()]
        [String]$SelectParameter
      )
      Begin { 
        Write-Verbose -Message ('Enter Function: {0}' -f $PSCmdlet.MyInvocation.MyCommand.Name)
        $SoftwareOutput = @()
        $InstalledSoftware = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*)#, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*)
        #$InstalledSoftware = (Get-ItemProperty -Path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*)
      }
      Process {
        Try 
        {
          if($SoftwareName -eq $null) 
          {
            $SoftwareOutput = $InstalledSoftware |
            Select-Object -Property Installdate, DisplayVersion, DisplayName #, UninstallString 
          }
          Else 
          {
            foreach($Item in $SoftwareName)
            {
              $SoftwareOutput += $InstalledSoftware |
              Where-Object -Property DisplayName -Match -Value $Item |
              Select-Object -Property Installdate, DisplayVersion, DisplayName #, UninstallString 
            }
          }
        }
        Catch 
        {
          # get error record
          $ErrorMessage  = $_.exception.message
          Write-Verbose -Message ('Error Message: {0}' -f $ErrorMessage)
        }
      }
      End{ 
        Switch ($SelectParameter){
          'DisplayName' 
          {
            $SoftwareOutput.displayname
          }
          'DisplayVersion' 
          {
            $SoftwareOutput.DisplayVersion
          }
          default 
          {
            $SoftwareOutput
          }
        }
      }
    } # End InstalledSoftware-Function
    <#bookmark Workstation Information #>
    function Get-WorkstationInfo    
    {
      param(
        [Parameter(Mandatory = $true,Position = 0)]
        [ValidateSet('Manufacturer','Model','Name','PrimaryOwnerName','Domain','serialnumber','PartOfDomain','Workgroup')] 
        [String]$Info
      )
      Write-Verbose -Message ('Enter Function: {0}' -f $PSCmdlet.MyInvocation.MyCommand.Name)
      if($Info -eq 'serialnumber')
      {
        (Get-WmiObject -Class win32_SystemEnclosure).serialnumber
      }
      else
      {
        (Get-WmiObject -Class:Win32_ComputerSystem).$Info
      }
    } # End Workstation Information-Function
    <#bookmark Get MAC Address #>
    function Get-MacAddress     
    {
      param(
        [Parameter(Position = 0)]
        [Switch]$LastFour
      )
      $MacAddress = (Get-NetAdapter -Physical | Where-Object -Property status -EQ -Value 'Up').macaddress
      if($LastFour)
      {
        $MacInfo = (($MacAddress.Split('-',5))[4]).replace('-',':')
      }
      else
      {
        $MacInfo = $MacAddress
      }
      $MacInfo
    } # End MacAddress-Function
    <#bookmark Facility Issues #>
    function Get-FacilityIssues
    {
      <#
          .SYNOPSIS
          Get-RoomStatus allows a place to put notes about the room.  
          Such as bad A/C or lights not working
      #>
      param
      (
        [Parameter(Mandatory, Position = 0)]
        [String]$RoomStatusFile,
        [Parameter(Mandatory, Position = 1)]
        [Object]$LatestStatus
      )
      $DblLine = ('=' * 20)
      $dateFormat = 'MMMM-dd'
      $CurrentDay = Get-Date -UFormat %A
      Switch ($CurrentDay){
        Monday 
        {
          $CurrentWeek = (Get-Date).AddDays(0).ToString($dateFormat)
        }
        Tuesday 
        {
          $CurrentWeek = (Get-Date).AddDays(-1).ToString($dateFormat)
        }
        Wednesday 
        {
          $CurrentWeek  = (Get-Date).AddDays(-2).ToString($dateFormat)
        }
        Thursday 
        {
          $CurrentWeek = (Get-Date).AddDays(-3).ToString($dateFormat)
        }
        Friday 
        {
          $CurrentWeek = (Get-Date).AddDays(-4).ToString($dateFormat)
        }
        Saturday 
        {
          $CurrentWeek = (Get-Date).AddDays(-5).ToString($dateFormat)
        }
        Sunday 
        {
          $CurrentWeek = (Get-Date).AddDays(-6).ToString($dateFormat)
        }
      }
      $FacilityIssueReport = [String]$($RoomStatusFile.Replace('.',('-WeekOf_{0}.' -f $CurrentWeek)))
      if(-not (Test-Path -Path $FacilityIssueReport ))
      {
        Write-Verbose -Message ('Creating the Issue Report for the week of {0}' -f $CurrentWeek)
        $null = New-Item -Path $FacilityIssueReport -ItemType File -Force
        'Fast Cruise Room Descrepancy Report ' | Out-File -FilePath $FacilityIssueReport -Append
      }
      $FacilityIssuesHeader = (@'
{2}
{3} - Building: {0}  Room: {1} 

'@ -f $LatestStatus.Building, $LatestStatus.Room, $DblLine, (Get-Date))
      $FacilityIssuesHeader | Out-File -FilePath $FacilityIssueReport -Append


      $DefaultInput = 'None Found'
      $DoNotWrite = @(' ', 'Exit' )
      do
      {
        # Write-Host ('Default: {0}' -f $DefaultInput)
        # Write-Host ('Do not write: {0}' -f $DoNotWrite)
  
        if($RoomIssue -ne $DefaultInput)
        {
          $RoomIssue = Show-VbForm -InputBox -Message 'Enter any issues with the room.  These will be sent to the facilities department.  Type "EXIT" to exit' -TitleBar 'Facility Issue Report' -DefaultValue $DefaultInput
      
          if ($RoomIssue -notin $DoNotWrite)
          {
            ('- {0}' -f $RoomIssue) | Out-File -FilePath $FacilityIssueReport -Append
            if($RoomIssue -match $DefaultInput)
            {
              $DoNotWrite += $RoomIssue
            }
          }
        }
        else
        {
          $RoomIssue = 'Exit'
        }
        #Write-Host ('Do not write: {0}' -f $DoNotWrite)
      }
      While($RoomIssue -notin $DoNotWrite)
    } <# END Facility Issues #>
    <#bookmark ComputerStat Hashtable #>
    Write-Verbose -Message 'Setting up the ComputerStat hash'
    $ComputerStat = [ordered]@{
      'Date'               = "$(Get-Date)"
      'ComputerName'       = "$env:COMPUTERNAME"
      'SerialNumber'       = 'N/A'
      'MacAddress'         = 'N/A'
      'UserName'           = "$env:USERNAME"
      'WSUS Search Success' = 'N/A'
      'WSUS Install Success' = 'N/A'
      'Department'         = 'N/A'
      'Building'           = 'N/A'
      'Room'               = 'N/A'
      'Desk'               = 'N/A'
    }
  } #End BEGIN region
  Process
  {
    if($ManualInput -eq $false)
    {
      <#bookmark Windows Updates #> 
      $LatestWSUSupdate = (New-Object -ComObject 'Microsoft.Update.AutoUpdate'). Results 
      $ComputerStat['WSUS Search Success'] = $LatestWSUSupdate.LastSearchSuccessDate
      $ComputerStat['WSUS Install Success'] = $LatestWSUSupdate.LastInstallationSuccessDate
      <#bookmark Get-MacAddress #>
      Write-Verbose -Message 'Getting Mac Address'
      $ComputerStat['MacAddress'] = Get-MacAddress
      Write-Verbose -Message 'Getting Serial Number'
      $ComputerStat['SerialNumber'] = Get-WorkstationInfo -Info serialnumber
      Write-Verbose -Message 'Getting Manufacturer'
      $ComputerStat['Manufacturer'] = Get-WorkstationInfo -Info Manufacturer
      Write-Verbose -Message 'Getting Model'
      $ComputerStat['Model'] = Get-WorkstationInfo -Info Model
      $PartOfDomain = Get-WorkstationInfo -Info PartOfDomain
      if($PartOfDomain -eq $true)
      {
        Write-Verbose -Message 'Getting Domain'
        $ComputerStat['Domain'] = Get-WorkstationInfo -Info Domain
      }
      else
      {
        Write-Verbose -Message 'Getting Workgroup'
        $ComputerStat['WorkGroup'] = Get-WorkstationInfo -Info Workgroup
        $FastCruiseReport = "$env:TEMP\FastCruise.csv"
        Write-Warning  -Message ('This computer is not attached to the domain') 
        Write-Output -InputObject ('{0}' -f $FastCruiseReport)
      }
      <#bookmark Software Versions #>
      #$ComputerStat['VmWare Version']  = Get-InstalledSoftware -SoftwareName 'Vmware' -SelectParameter DisplayVersion
      
      foreach($SoftwareItem in $SoftwareChecks)
      {
        $ComputerStat["$SoftwareItem Version"] = Get-InstalledSoftware -SoftwareName $SoftwareItem -SelectParameter DisplayVersion
      }
      if($LocalCruiseFile.Length -gt 0)
      {
        Write-Verbose -Message 'Getting Last Status recorded locally'
        $Script:LatestStatus = (Get-LastComputerStatus -LastCruiseStatus $LocalCruiseFile)
      }
      else
      {
        $Script:LatestStatus = (Get-LastComputerStatus -LastCruiseStatus $FastCruiseReport)
      }
      <#bookmark Location Verification #>
      $ComputerLocation = (@'

ComputerName: (Assest Tag)
- {0}

Serial Number:
- {6}

Department:
- {1}

Building:
- {2}

Room:
- {3}

Desk:
- {4}

Phone
- {5}

          
'@ -f $LatestStatus.ComputerName, $LatestStatus.Department, $LatestStatus.Building, $LatestStatus.Room, $LatestStatus.Desk, $LatestStatus.Phone, $LatestStatus.SerialNumber)

      <#bookmark Application Test #> 
      $FunctionTest = 'No' #Show-VbForm -YesNoBox -Message 'Perform Applicaion Tests (MS Office and Adobe)?' 
      if($FunctionTest -eq 'Yes')
      {
        $ComputerStat['Adobe Test'] = Start-ApplicationTest -WaitTest @PDFApplicationTestSplat
        $ComputerStat['MS Office Test'] = Start-ApplicationTest -WaitTest @PowerPointApplicationTestSplat
      }
      Else
      {
        Write-Verbose -Message 'TestResult: Bypassed'
        $TestResult = 'Bypassed'
        $ComputerStat['MS Office Test'] = $TestResult
        $ComputerStat['Adobe Test'] = $TestResult
      }
      $LocationVerification = Show-VbForm -YesNoBox -Message $ComputerLocation
    }
    if($ManualInput -eq $true)
    {
      $LocationVerification = 'No'
      $ComputerStat['ComputerName'] = Show-VbForm -InputBox -Message 'ComputerName: (Assest Tag)' -TitleBar 'ComputerName' -DefaultValue 'D1234567'
      $ComputerStat['SerialNumber'] = 'Manual Input'
      $TestResult = 'Manual Input'
      $ComputerStat['MS Office Test'] = $TestResult
      $ComputerStat['Adobe Test'] = $TestResult
    }
    if($LocationVerification -eq 'No')
    {
      Get-ComputerLocation -jsonFilePath $jsonFilePath
      Write-Verbose -Message ('Computer Description: ABC-DEF-{0}-{1}-{2}{3}' -f $LclDept, $LclBuild, $LclRm, $LclDesk)
      $ComputerStat['Department'] = $LclDept 
      $ComputerStat['Building'] = $LclBuild
      $ComputerStat['Room'] = $LclRm
      $ComputerStat['Desk'] = $LclDesk
    }
    else
    {
      $ComputerStat['Department'] = $($LatestStatus.Department)
      $ComputerStat['Building'] = $($LatestStatus.Building)
      $ComputerStat['Room'] = $($LatestStatus.Room)
      $ComputerStat['Desk'] = $($LatestStatus.Desk)
      $ComputerStat['Phone'] = $($LatestStatus.Phone)
    }
    if($LocationVerification -eq 'No')
    {
      <#bookmark Local phone number #> 
      $RegexPhone = '^\d{3}-\d{3}-\d{4}'
      While($Phone -notmatch $RegexPhone)
      {
        $Phone = Show-VbForm -InputBox -Message 'Nearest Phone Number: ' -DefaultValue '757-555-1234'
        if($Phone -eq '')
        {
          Break
        }
      }
      $ComputerStat['Phone'] = $Phone
    }
    <#bookmark Fast cruise notes #>
    [string]$Notes = Show-VbForm -InputBox -Message 'Notes about this cruise: ' -DefaultValue 'Related to the computer'
    if($Notes -eq 'Related to the computer')
    {
      $Notes = ''
    }
    $ComputerStat['Notes'] = $Notes

    <#bookmark Facility Test #>
    if('AB' -match $ComputerStat.Desk)
    {
      Get-FacilityIssues @FacilityIssuesSplat -LatestStatus $ComputerStat
    }
     
  } #End PROCESS region
  END
  {    
    $ComputerStat |
    ForEach-Object -Process {
      [pscustomobject]$_
    } |
    Export-Csv -Path $FastCruiseReport -NoTypeInformation -Append -Force
    
    $ComputerStat |
    ForEach-Object -Process {
      [pscustomobject]$_
    } |
    Export-Csv -Path $LocalCruiseFile -NoTypeInformation -Force

    Write-Output -InputObject 'The information recorded'
    Write-Output -InputObject ('Local File: {0}' -f $LocalCruiseFile)

    $ComputerStat | Format-Table
    <#bookmark Fast cruising shipmates #>
    Write-Output -InputObject 'Fast Cruise shipmates'
    Import-Csv -Path $FastCruiseReport |
    Select-Object -Last 4 -Property Date, Username, Department, Building, Room, Phone |
    Format-Table 
  } #End END region
}
# This is what calls the script to run
Start-FastCruise @FastCruiseSplat # Make sure you have updated and completed the "Splats" at the top of the script
### Everything below this line is only to provide the menu that you see after running the script.  Everything below this line can be deleted without impacting the functionality of the script.
<#bookmark ASCII Menu #>
function Show-AsciiMenu 
{
  <#
      .SYNOPSIS
      Create a simple menu.
  #>
  [CmdletBinding()]
  param
  (
    [string]$Title = 'Title',
    [String[]]$MenuItems = 'None',
    [string]$TitleColor = 'Red',
    [string]$LineColor = 'Yellow',
    [string]$MenuItemColor = 'Cyan'
  )
  Begin{
    # Set Variables
    $i = 1
    $Tab = "`t"
    $VertLine = '║'
    function Write-HorizontalLine
    {
      param
      (
        [Parameter(Position = 0)]
        [string]
        $DrawLine = 'Top'
      )
      Switch ($DrawLine) {
        Top 
        {
          Write-Host -Object ('╔{0}╗' -f $HorizontalLine) -ForegroundColor $LineColor
        }
        Middle 
        {
          Write-Host -Object ('╠{0}╣' -f $HorizontalLine) -ForegroundColor $LineColor
        }
        Bottom 
        {
          Write-Host -Object ('╚{0}╝' -f $HorizontalLine) -ForegroundColor $LineColor
        }
      }
    }
    function Get-Padding
    {
      param
      (
        [Parameter(Mandatory, Position = 0)]
        [int]$Multiplier 
      )
      "`0"*$Multiplier
    }
    function Write-MenuTitle
    {
      Write-Host -Object ('{0}{1}' -f $VertLine, $TextPadding) -NoNewline -ForegroundColor $LineColor
      Write-Host -Object ($Title) -NoNewline -ForegroundColor $TitleColor
      if($TotalTitlePadding % 2 -eq 1)
      {
        $TextPadding = Get-Padding -Multiplier ($TitlePaddingCount + 1)
      }
      Write-Host -Object ('{0}{1}' -f $TextPadding, $VertLine) -ForegroundColor $LineColor
    }
    function Write-MenuItems
    {
      foreach($menuItem in $MenuItems)
      {
        $number = $i++
        $ItemPaddingCount = $TotalLineWidth - $menuItem.Length - 6 #This number is needed to offset the Tab, space and 'dot'
        $ItemPadding = Get-Padding -Multiplier $ItemPaddingCount
        Write-Host -Object $VertLine  -NoNewline -ForegroundColor $LineColor
        Write-Host -Object ('{0}{1}. {2}{3}' -f $Tab, $number, $menuItem, $ItemPadding) -NoNewline -ForegroundColor $LineColor
        Write-Host -Object $VertLine -ForegroundColor $LineColor
      }
    }
  }
  Process
  {
    $TitleCount = $Title.Length
    $LongestMenuItemCount = ($MenuItems | Measure-Object -Maximum -Property Length).Maximum
    Write-Debug -Message ('LongestMenuItemCount = {0}' -f $LongestMenuItemCount)
    if  ($TitleCount -gt $LongestMenuItemCount)
    {
      $ItemWidthCount = $TitleCount
    }
    else
    {
      $ItemWidthCount = $LongestMenuItemCount
    }
    if($ItemWidthCount % 2 -eq 1)
    {
      $ItemWidth = $ItemWidthCount + 1
    }
    else
    {
      $ItemWidth = $ItemWidthCount
    }
    Write-Debug -Message ('Item Width = {0}' -f $ItemWidth)
    $TotalLineWidth = $ItemWidth + 10
    Write-Debug -Message ('Total Line Width = {0}' -f $TotalLineWidth)
    $TotalTitlePadding = $TotalLineWidth - $TitleCount
    Write-Debug -Message ('Total Title Padding  = {0}' -f $TotalTitlePadding)
    $TitlePaddingCount = [math]::Floor($TotalTitlePadding / 2)
    Write-Debug -Message ('Title Padding Count = {0}' -f $TitlePaddingCount)
    $HorizontalLine = '═'*$TotalLineWidth
    $TextPadding = Get-Padding -Multiplier $TitlePaddingCount
    Write-Debug -Message ('Text Padding Count = {0}' -f $TextPadding.Length)
    Write-HorizontalLine -DrawLine Top
    Write-MenuTitle
    Write-HorizontalLine -DrawLine Middle
    Write-MenuItems
    Write-HorizontalLine -DrawLine Bottom
  }
  End
  {}
}
do
{
  #Show-AsciiMenu -Title 'THIS IS THE TITLE' -MenuItems 'Exchange Server', 'Active Directory', 'Sytem Center Configuration Manager', 'Lync Server' -TitleColor Red  -MenuItemColor green
  Show-AsciiMenu -Title 'EXIT STRATAGY' -MenuItems 'Retart-Fastcruise', 'Manually Add Computer', 'Record a Facility Issue', 'Restart Computer', 'Quit to Prompt' #-Debug
  $Blueberry = Read-Host -Prompt 'Select Number'
  switch($Blueberry)
  {
    1 
    {
      Clear-Host #Clears the console.  This shouldn't be needed once the script can be run directly from PS
      Write-Host -Object "`n`n"
      Start-FastCruise @FastCruiseSplat # Make sure you have updated and completed the "Splats" at the top of the script
      $Blueberry = $null
    }
    2
    {
      Start-FastCruise @ManualInputSplat
    }
    3 
    {
      #Write-Host 'Not Operational' -ForegroundColor Cyan
      Get-FacilityIssues @FacilityIssuesSplat -LatestStatus $LatestStatus
    }
    4 
    {
      Restart-Computer
    }
    5
    {
      Break
    }
  }
}
Until ($Blueberry -eq 5)
